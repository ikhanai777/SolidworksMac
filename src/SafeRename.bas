Attribute VB_Name = "SafeRename"
'===============================================================================
' SafeRename.bas  -  SOLIDWORKS VBA macro
'
' Renames part / sub-assembly files that are used inside the active assembly
' WITHOUT breaking references in the assemblies, drawings and derived parts
' that use them.
'
' How it works (short version):
'   1. You select one or more components in the active assembly (in the
'      FeatureManager tree or in the graphics area) and run the macro.
'   2. For every selected file you type a new file name.
'   3. The macro scans a folder (default: the top assembly's folder and all
'      sub-folders) for every .SLDASM / .SLDDRW / .SLDPRT that references the
'      file (where-used search via GetDocumentDependencies2 - nothing is
'      opened during the scan).
'   4. All those dependents are opened, so SOLIDWORKS holds them in memory.
'   5. The file is renamed in memory (ModelDocExtension.RenameDocument, the
'      same engine as "Rename Component > Rename Document" in the UI; with a
'      Save As fallback for older versions).
'   6. The renamed file and every dependent are saved, so each one now stores
'      the NEW path.
'   7. Every dependent is re-read from disk and verified to point at the new
'      file and no longer at the old one. A log file is written.
'   8. Only when everything verified OK you are offered to move the old file
'      into a "_RenameBackup" folder (default answer keeps it untouched).
'      Optionally, a drawing with the same name as the old file is renamed too.
'
' Requirements: SOLIDWORKS 2018 or newer recommended (RenameDocument API).
'               Older versions use the Save As fallback.
' Not for files managed in SOLIDWORKS PDM - use the vault's own Rename there.
'===============================================================================
Option Explicit

'------------------------------ user settings ----------------------------------
' Use Save As (not copy) when RenameDocument is unavailable or fails.
Private Const ALLOW_SAVEAS_FALLBACK As Boolean = True
' Offer to rename a drawing that has the same name as the renamed model.
Private Const OFFER_DRAWING_RENAME As Boolean = True
' Close the dependents that the macro had to open, when it is done.
Private Const CLOSE_OPENED_DOCS As Boolean = True
' Name of the folder (next to the old file) that receives the old file.
Private Const BACKUP_FOLDER_NAME As String = "_RenameBackup"
' Write a log file next to the top assembly.
Private Const WRITE_LOG_FILE As Boolean = True

'------------------------------ API constants ----------------------------------
' (Defined locally so the macro does not depend on the swconst reference.)
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3
Private Const swOpenDocOptions_Silent As Long = 1
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swSaveAsOptions_SaveReferenced As Long = 4
Private Const swSaveAsCurrentVersion As Long = 0
Private Const swDontRebuild As Long = 1
Private Const swComponentSuppressed As Long = 0
Private Const swComponentLightweight As Long = 1
Private Const swComponentFullyResolved As Long = 2
Private Const swComponentFullyLightweight As Long = 4

'------------------------------ module state -----------------------------------
Private swApp As SldWorks.SldWorks
Private fso As Object
Private logText As String
Private depCache As Object          ' file path -> dependency array
Private skippedText As String       ' components ignored during selection

'===============================================================================
Sub main()
    Set swApp = Application.SldWorks
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set depCache = CreateObject("Scripting.Dictionary")
    depCache.CompareMode = vbTextCompare
    logText = ""

    Dim topModel As SldWorks.ModelDoc2
    Set topModel = swApp.ActiveDoc

    If topModel Is Nothing Then
        MsgBox "Open an assembly and select the component(s) to rename.", vbExclamation, "Safe Rename"
        Exit Sub
    End If
    If topModel.GetType <> swDocASSEMBLY Then
        MsgBox "The active document must be an assembly.", vbExclamation, "Safe Rename"
        Exit Sub
    End If
    If topModel.GetPathName = "" Then
        MsgBox "Save the assembly once before renaming its components.", vbExclamation, "Safe Rename"
        Exit Sub
    End If

    LogMsg "Safe Rename started " & Now
    LogMsg "Top assembly: " & topModel.GetPathName

    '--- 1. collect selected components (unique by file) -----------------------
    Dim jobs As Object                      ' old path -> Component2
    Set jobs = CollectSelectedComponents(topModel)
    If jobs.Count = 0 Then
        MsgBox "Select at least one component (part or sub-assembly) in the assembly, then run the macro again." & _
               vbCrLf & vbCrLf & GetSkippedText(), vbExclamation, "Safe Rename"
        Exit Sub
    End If

    '--- 2. ask for new names --------------------------------------------------
    Dim oldPaths() As String, newNames() As String, comps() As Object
    Dim nJobs As Long
    If Not AskNewNames(jobs, oldPaths, newNames, comps, nJobs) Then Exit Sub
    If nJobs = 0 Then
        MsgBox "Nothing to rename.", vbInformation, "Safe Rename"
        Exit Sub
    End If

    '--- 3. where to look for dependents ---------------------------------------
    Dim searchRoot As String
    searchRoot = AskSearchRoot(fso.GetParentFolderName(topModel.GetPathName))
    If searchRoot = "" Then Exit Sub
    LogMsg "Search folder: " & searchRoot

    SetStatus "Safe Rename: listing SOLIDWORKS files in " & searchRoot & " ..."
    Dim allFiles As Collection
    Set allFiles = New Collection
    ListSwFiles fso.GetFolder(searchRoot), allFiles
    LogMsg "SOLIDWORKS files found in search folder: " & allFiles.Count

    '--- 4. where-used search --------------------------------------------------
    Dim dependents() As Collection
    ReDim dependents(1 To nJobs)
    Dim i As Long, summary As String
    For i = 1 To nJobs
        Set dependents(i) = FindDependents(oldPaths(i), allFiles)
        summary = summary & fso.GetFileName(oldPaths(i)) & "  ->  " & newNames(i) & "." & _
                  fso.GetExtensionName(oldPaths(i)) & vbCrLf & _
                  "      referenced by " & dependents(i).Count & " file(s) in the search folder" & _
                  ListPreview(dependents(i), 6) & vbCrLf
    Next i

    Dim answer As VbMsgBoxResult
    answer = MsgBox("The following will be renamed:" & vbCrLf & vbCrLf & summary & vbCrLf & _
                    "All referencing files will be opened and SAVED (the top assembly is saved too)." & vbCrLf & _
                    "Make sure nobody else has these files open and that you have a backup." & vbCrLf & vbCrLf & _
                    "Continue?", vbQuestion + vbOKCancel, "Safe Rename - confirm")
    If answer <> vbOK Then
        LogMsg "Cancelled by user at confirmation."
        FinishLog topModel
        Exit Sub
    End If

    '--- 5. do the work --------------------------------------------------------
    Dim okCount As Long, report As String, ok As Boolean
    For i = 1 To nJobs
        ok = ProcessJob(topModel, comps(i), oldPaths(i), newNames(i), dependents(i))
        If ok Then
            okCount = okCount + 1
            report = report & "OK      " & fso.GetFileName(oldPaths(i)) & " -> " & newNames(i) & vbCrLf
        Else
            report = report & "FAILED  " & fso.GetFileName(oldPaths(i)) & " (see log)" & vbCrLf
        End If
    Next i

    SetStatus ""
    Dim logPath As String
    logPath = FinishLog(topModel)

    MsgBox okCount & " of " & nJobs & " file(s) renamed and verified." & vbCrLf & vbCrLf & report & _
           IIf(logPath <> "", vbCrLf & "Log: " & logPath, ""), _
           IIf(okCount = nJobs, vbInformation, vbExclamation), "Safe Rename - done"
End Sub

'===============================================================================
' Selection
'===============================================================================
Private Function CollectSelectedComponents(topModel As SldWorks.ModelDoc2) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    skippedText = ""

    Dim selMgr As SldWorks.SelectionMgr
    Set selMgr = topModel.SelectionManager

    Dim n As Long, i As Long, p As String
    Dim comp As SldWorks.Component2
    n = selMgr.GetSelectedObjectCount2(-1)

    For i = 1 To n
        Set comp = Nothing
        On Error Resume Next
        Set comp = selMgr.GetSelectedObjectsComponent4(i, -1)
        On Error GoTo 0
        If Not comp Is Nothing Then
            p = comp.GetPathName
            If comp.IsVirtual Then
                Skip comp.Name2 & ": virtual component (no file to rename - use Rename Component)"
            ElseIf p = "" Then
                Skip comp.Name2 & ": component has no file path"
            ElseIf comp.GetSuppression2 = swComponentSuppressed Then
                Skip comp.Name2 & ": component is suppressed - unsuppress it first"
            ElseIf Not result.Exists(p) Then
                result.Add p, comp
            End If
        End If
    Next i

    Set CollectSelectedComponents = result
End Function

Private Sub Skip(ByVal msg As String)
    skippedText = skippedText & "  - " & msg & vbCrLf
    LogMsg "Skipped: " & msg
End Sub

Private Function GetSkippedText() As String
    If skippedText <> "" Then GetSkippedText = "Skipped:" & vbCrLf & skippedText
End Function

'===============================================================================
' Ask for new names and validate them
'===============================================================================
Private Function AskNewNames(jobs As Object, oldPaths() As String, newNames() As String, _
                             comps() As Object, nJobs As Long) As Boolean
    Dim keys As Variant, k As Long
    Dim oldPath As String, oldBase As String, ext As String, folder As String
    Dim newName As String, newPath As String, problem As String
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    keys = jobs.keys
    ReDim oldPaths(1 To jobs.Count)
    ReDim newNames(1 To jobs.Count)
    ReDim comps(1 To jobs.Count)
    nJobs = 0

    For k = 0 To UBound(keys)
        oldPath = keys(k)
        oldBase = fso.GetBaseName(oldPath)
        ext = fso.GetExtensionName(oldPath)
        folder = fso.GetParentFolderName(oldPath)
        newName = oldBase

        Do
            newName = InputBox("Current file:" & vbCrLf & oldPath & vbCrLf & vbCrLf & _
                               "Enter the NEW file name (without ." & ext & ")." & vbCrLf & _
                               "Leave empty / Cancel to skip this file.", _
                               "Safe Rename - new name", newName)
            newName = Trim$(newName)
            If newName = "" Then Exit Do

            ' strip an extension the user may have typed
            If LCase$(Right$(newName, Len(ext) + 1)) = "." & LCase$(ext) Then
                newName = Left$(newName, Len(newName) - Len(ext) - 1)
            End If
            newPath = fso.BuildPath(folder, newName & "." & ext)

            problem = ""
            If newName Like "*[\/:*?""<>|]*" Then
                problem = "The name contains invalid characters  \ / : * ? "" < > |"
            ElseIf StrComp(newName, oldBase, vbBinaryCompare) = 0 Then
                problem = "The new name is the same as the current name."
            ElseIf StrComp(newName, oldBase, vbTextCompare) = 0 Then
                problem = "Changing only upper/lower case is not supported by Windows file renaming." & vbCrLf & _
                          "Rename to a temporary name first, then to the final name."
            ElseIf fso.FileExists(newPath) Then
                problem = "A file with that name already exists:" & vbCrLf & newPath
            ElseIf Not swApp.GetOpenDocumentByName(newPath) Is Nothing Then
                problem = "A document with that name is already open in SOLIDWORKS."
            ElseIf used.Exists(newPath) Then
                problem = "You already used this name for another file in this run."
            ElseIf IsReadOnlyFile(oldPath) Then
                problem = "The current file is read-only (checked in to PDM or write-protected)."
            End If

            If problem = "" Then Exit Do
            If MsgBox(problem & vbCrLf & vbCrLf & "Try another name?", vbExclamation + vbRetryCancel, _
                      "Safe Rename") <> vbRetry Then
                newName = ""
                Exit Do
            End If
        Loop

        If newName <> "" Then
            nJobs = nJobs + 1
            oldPaths(nJobs) = oldPath
            newNames(nJobs) = newName
            Set comps(nJobs) = jobs(oldPath)
            used.Add newPath, True
            LogMsg "Planned: " & oldPath & "  ->  " & newPath
        Else
            LogMsg "Skipped by user: " & oldPath
        End If
    Next k

    AskNewNames = True
End Function

Private Function IsReadOnlyFile(ByVal path As String) As Boolean
    On Error Resume Next
    IsReadOnlyFile = (fso.GetFile(path).Attributes And 1) <> 0
End Function

'===============================================================================
' Search folder
'===============================================================================
Private Function AskSearchRoot(ByVal defaultFolder As String) As String
    Dim r As VbMsgBoxResult
    r = MsgBox("Where should the macro look for assemblies and drawings that use the file(s)?" & vbCrLf & vbCrLf & _
               defaultFolder & vbCrLf & "(and all sub-folders)" & vbCrLf & vbCrLf & _
               "Yes      = use this folder" & vbCrLf & _
               "No       = choose a different (e.g. higher-level project) folder" & vbCrLf & _
               "Cancel  = abort", vbQuestion + vbYesNoCancel, "Safe Rename - search folder")
    If r = vbCancel Then Exit Function
    If r = vbYes Then
        AskSearchRoot = defaultFolder
        Exit Function
    End If

    Dim sh As Object, f As Object
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0, "Folder to search (sub-folders included) for files that reference the renamed file(s):", _
                               &H1 Or &H40, 0)
    If f Is Nothing Then Exit Function
    AskSearchRoot = f.Self.path
End Function

Private Sub ListSwFiles(folder As Object, result As Collection)
    Dim f As Object, sf As Object, ext As String
    On Error Resume Next            ' skip folders we have no access to
    For Each f In folder.Files
        If Left$(f.Name, 2) <> "~$" Then
            ext = LCase$(fso.GetExtensionName(f.Name))
            If ext = "sldasm" Or ext = "slddrw" Or ext = "sldprt" Then result.Add f.path
        End If
    Next f
    For Each sf In folder.SubFolders
        If StrComp(sf.Name, BACKUP_FOLDER_NAME, vbTextCompare) <> 0 Then ListSwFiles sf, result
    Next sf
End Sub

'===============================================================================
' Where-used
'===============================================================================
Private Function GetDeps(ByVal filePath As String) As Variant
    If depCache.Exists(filePath) Then
        GetDeps = depCache(filePath)
        Exit Function
    End If
    Dim v As Variant
    On Error Resume Next
    ' Traverse = False (direct references only), Search = True (resolve with search rules)
    v = swApp.GetDocumentDependencies2(filePath, False, True, False)
    On Error GoTo 0
    depCache.Add filePath, v
    GetDeps = v
End Function

' Returns True if filePath directly references target (full-path match).
' nameOnly is set when only the file NAME matches (same name, other folder).
Private Function RefersTo(ByVal filePath As String, ByVal target As String, ByRef nameOnly As Boolean) As Boolean
    Dim deps As Variant, i As Long, s As String, targetName As String
    nameOnly = False
    deps = GetDeps(filePath)
    If Not IsArray(deps) Then Exit Function
    targetName = fso.GetFileName(target)
    For i = LBound(deps) To UBound(deps)
        s = CStr(deps(i))
        If InStr(s, "\") > 0 Then
            If StrComp(s, target, vbTextCompare) = 0 Then
                RefersTo = True
                Exit Function
            ElseIf StrComp(fso.GetFileName(s), targetName, vbTextCompare) = 0 Then
                nameOnly = True
            End If
        End If
    Next i
End Function

Private Function FindDependents(ByVal target As String, allFiles As Collection) As Collection
    Dim result As New Collection, f As Variant, nameOnly As Boolean, i As Long
    For Each f In allFiles
        i = i + 1
        If i Mod 25 = 0 Then
            SetStatus "Safe Rename: where-used search for " & fso.GetFileName(target) & " (" & i & "/" & allFiles.Count & ")"
            DoEvents
        End If
        If StrComp(f, target, vbTextCompare) <> 0 Then
            If RefersTo(f, target, nameOnly) Then
                result.Add f
                LogMsg "  Dependent of " & fso.GetFileName(target) & ": " & f
            ElseIf nameOnly Then
                LogMsg "  WARNING: " & f & " references a DIFFERENT file with the same name " & _
                    fso.GetFileName(target) & " - not modified."
            End If
        End If
    Next f
    Set FindDependents = result
End Function

'===============================================================================
' Rename one file
'===============================================================================
Private Function ProcessJob(topModel As SldWorks.ModelDoc2, comp As SldWorks.Component2, _
                            ByVal oldPath As String, ByVal newName As String, _
                            dependents As Collection) As Boolean
    Dim folder As String, ext As String, newPath As String
    folder = fso.GetParentFolderName(oldPath)
    ext = fso.GetExtensionName(oldPath)
    newPath = fso.BuildPath(folder, newName & "." & ext)

    LogMsg ""
    LogMsg "=== " & oldPath & "  ->  " & newPath

    Dim opened As New Collection            ' docs this macro opened
    Dim toSave As Object                    ' path -> ModelDoc2 that must be saved
    Set toSave = CreateObject("Scripting.Dictionary")
    toSave.CompareMode = vbTextCompare

    Dim d As Variant, doc As SldWorks.ModelDoc2, errs As Long, warns As Long

    '--- open every dependent so SOLIDWORKS updates it in memory ----------------
    For Each d In dependents
        SetStatus "Safe Rename: opening " & fso.GetFileName(d)
        Set doc = swApp.GetOpenDocumentByName(d)
        If doc Is Nothing Then
            errs = 0: warns = 0
            Set doc = swApp.OpenDoc6(d, DocTypeFromPath(d), swOpenDocOptions_Silent, "", errs, warns)
            If doc Is Nothing Then
                LogMsg "  ERROR: could not open dependent " & d & " (error " & errs & "). Rename aborted for this file."
                CloseOpened opened
                Exit Function
            End If
            opened.Add CStr(d)
            LogMsg "  Opened " & d
        End If
        If IsReadOnlyDoc(doc) Then
            LogMsg "  ERROR: dependent is read-only / opened read-only: " & d & ". Rename aborted for this file."
            CloseOpened opened
            Exit Function
        End If
        If Not toSave.Exists(CStr(d)) Then toSave.Add CStr(d), doc
    Next d

    '--- back to the top assembly ----------------------------------------------
    errs = 0
    swApp.ActivateDoc3 topModel.GetTitle, False, swDontRebuild, errs

    '--- parents of the file inside the active assembly (even outside search folder)
    Dim swAssy As SldWorks.AssemblyDoc, vComps As Variant, k As Long
    Dim c As SldWorks.Component2, parent As SldWorks.Component2, pDoc As SldWorks.ModelDoc2
    Set swAssy = topModel
    vComps = swAssy.GetComponents(False)
    If IsArray(vComps) Then
        For k = LBound(vComps) To UBound(vComps)
            Set c = vComps(k)
            If StrComp(c.GetPathName, oldPath, vbTextCompare) = 0 Then
                Select Case c.GetSuppression2
                    Case swComponentSuppressed
                        LogMsg "  WARNING: suppressed instance " & c.Name2 & " - its reference will be checked during verification."
                    Case swComponentLightweight, swComponentFullyLightweight
                        c.SetSuppression2 swComponentFullyResolved
                End Select
                Set parent = c.GetParent
                If parent Is Nothing Then
                    Set pDoc = topModel
                Else
                    Set pDoc = parent.GetModelDoc2
                End If
                If Not pDoc Is Nothing Then
                    If Not toSave.Exists(pDoc.GetPathName) Then toSave.Add pDoc.GetPathName, pDoc
                End If
            End If
        Next k
    End If

    Dim compModel As SldWorks.ModelDoc2
    Set compModel = comp.GetModelDoc2
    If compModel Is Nothing Then
        LogMsg "  ERROR: the component's document is not loaded. Resolve the component and try again."
        CloseOpened opened
        Exit Function
    End If

    '--- rename in memory -------------------------------------------------------
    SetStatus "Safe Rename: renaming " & fso.GetFileName(oldPath)
    Dim renamed As Boolean, rc As Long, extObj As Object
    topModel.ClearSelection2 True
    If comp.Select4(False, Nothing, False) Then
        rc = -1
        Set extObj = topModel.Extension         ' late bound: compiles on old versions too
        On Error Resume Next
        rc = extObj.RenameDocument(newName)
        If Err.Number <> 0 Then
            LogMsg "  RenameDocument not available (" & Err.Description & ")."
            rc = -1
        End If
        On Error GoTo 0
        renamed = (rc = 0)
        If renamed Then
            LogMsg "  Renamed in memory with RenameDocument."
        ElseIf rc <> -1 Then
            LogMsg "  RenameDocument returned error code " & rc & " (swRenameDocumentError_e)."
        End If
    Else
        LogMsg "  Could not select the component for RenameDocument."
    End If
    topModel.ClearSelection2 True

    If Not renamed And ALLOW_SAVEAS_FALLBACK Then
        errs = 0: warns = 0
        renamed = compModel.Extension.SaveAs(newPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errs, warns)
        If renamed Then
            LogMsg "  Renamed with Save As fallback."
        Else
            LogMsg "  ERROR: Save As fallback failed (error " & errs & ")."
        End If
    End If

    If Not renamed Then
        LogMsg "  ERROR: rename failed - nothing was saved for this file."
        CloseOpened opened
        Exit Function
    End If

    '--- save: top assembly (+ referenced), renamed file, then every dependent --
    SetStatus "Safe Rename: saving"
    topModel.ForceRebuild3 True
    errs = 0: warns = 0
    topModel.SetSaveFlag
    If Not topModel.Save3(swSaveAsOptions_Silent + swSaveAsOptions_SaveReferenced, errs, warns) Then
        LogMsg "  WARNING: saving top assembly reported error " & errs
    End If

    If Not fso.FileExists(newPath) Then
        errs = 0: warns = 0
        compModel.SetSaveFlag
        compModel.Save3 swSaveAsOptions_Silent, errs, warns
    End If

    Dim pass As Long, key As Variant
    For pass = 1 To 2                       ' models first, drawings last
        For Each key In toSave.keys
            Set doc = toSave(key)
            If (pass = 1 And doc.GetType <> swDocDRAWING) Or (pass = 2 And doc.GetType = swDocDRAWING) Then
                errs = 0: warns = 0
                doc.SetSaveFlag
                If doc.Save3(swSaveAsOptions_Silent, errs, warns) Then
                    LogMsg "  Saved " & key
                Else
                    LogMsg "  ERROR saving " & key & " (error " & errs & ")"
                End If
            End If
        Next key
    Next pass

    '--- verify from disk -------------------------------------------------------
    SetStatus "Safe Rename: verifying"
    Dim allOk As Boolean
    allOk = True
    If fso.FileExists(newPath) Then
        LogMsg "  New file exists on disk: " & newPath
    Else
        LogMsg "  ERROR: new file not found on disk: " & newPath
        allOk = False
    End If

    Dim savedPaths As Collection
    Set savedPaths = New Collection
    For Each key In toSave.keys
        savedPaths.Add CStr(toSave(key).GetPathName)  ' current path (in case a parent was renamed earlier)
    Next key
    For Each d In savedPaths
        If depCache.Exists(d) Then depCache.Remove d
        If StrComp(d, newPath, vbTextCompare) <> 0 Then
            If Not VerifyReference(CStr(d), fso.GetFileName(oldPath), fso.GetFileName(newPath)) Then allOk = False
        End If
    Next d

    '--- clean up ---------------------------------------------------------------
    If CLOSE_OPENED_DOCS Then CloseOpened opened

    If allOk Then
        LogMsg "  VERIFIED: all " & savedPaths.Count & " referencing file(s) now point to " & fso.GetFileName(newPath)
        If OFFER_DRAWING_RENAME Then RenameMatchingDrawings oldPath, newPath, dependents
        HandleOldFile oldPath, newPath
    Else
        LogMsg "  Verification FAILED - old file kept untouched. Check the entries above."
    End If

    ProcessJob = allOk
End Function

Private Function VerifyReference(ByVal filePath As String, ByVal oldName As String, ByVal newName As String) As Boolean
    Dim deps As Variant, i As Long, s As String, hasNew As Boolean, hasOld As Boolean
    On Error Resume Next
    ' Search = False: read the paths exactly as they are stored in the file
    deps = swApp.GetDocumentDependencies2(filePath, False, False, False)
    On Error GoTo 0
    If IsArray(deps) Then
        For i = LBound(deps) To UBound(deps)
            s = CStr(deps(i))
            If InStr(s, "\") > 0 Then
                If StrComp(fso.GetFileName(s), newName, vbTextCompare) = 0 Then hasNew = True
                If StrComp(fso.GetFileName(s), oldName, vbTextCompare) = 0 Then hasOld = True
            End If
        Next i
    End If
    If hasNew And Not hasOld Then
        LogMsg "  Verified " & filePath
        VerifyReference = True
    ElseIf hasNew And hasOld Then
        LogMsg "  ERROR: " & filePath & " references BOTH the new and the old file (suppressed instance?)."
    Else
        LogMsg "  ERROR: " & filePath & " does not reference the new file."
    End If
End Function

'===============================================================================
' Optional post steps
'===============================================================================
Private Sub RenameMatchingDrawings(ByVal oldPath As String, ByVal newPath As String, dependents As Collection)
    Dim d As Variant, oldBase As String, newDrw As String, doc As SldWorks.ModelDoc2
    oldBase = fso.GetBaseName(oldPath)
    For Each d In dependents
        If LCase$(fso.GetExtensionName(d)) = "slddrw" _
           And StrComp(fso.GetBaseName(d), oldBase, vbTextCompare) = 0 _
           And StrComp(fso.GetParentFolderName(d), fso.GetParentFolderName(oldPath), vbTextCompare) = 0 Then

            newDrw = fso.BuildPath(fso.GetParentFolderName(d), fso.GetBaseName(newPath) & ".SLDDRW")
            If fso.FileExists(newDrw) Then
                LogMsg "  Drawing not renamed, target exists: " & newDrw
            ElseIf MsgBox("Also rename the drawing?" & vbCrLf & vbCrLf & d & vbCrLf & "  ->  " & newDrw, _
                          vbQuestion + vbYesNo, "Safe Rename - drawing") = vbYes Then
                ' nothing references a drawing, so a plain file rename is safe
                Set doc = swApp.GetOpenDocumentByName(d)
                If Not doc Is Nothing Then swApp.CloseDoc doc.GetTitle
                Set doc = swApp.GetOpenDocumentByName(d)
                If Not doc Is Nothing Then
                    LogMsg "  Drawing still open, not renamed: " & d
                Else
                    On Error Resume Next
                    fso.MoveFile d, newDrw
                    If Err.Number = 0 Then
                        LogMsg "  Drawing renamed: " & d & " -> " & newDrw
                    Else
                        LogMsg "  ERROR renaming drawing " & d & ": " & Err.Description
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next d
End Sub

Private Sub HandleOldFile(ByVal oldPath As String, ByVal newPath As String)
    If Not fso.FileExists(oldPath) Then
        LogMsg "  Old file no longer exists on disk (renamed by SOLIDWORKS)."
        Exit Sub
    End If
    If Not swApp.GetOpenDocumentByName(oldPath) Is Nothing Then
        LogMsg "  Old file is still open in SOLIDWORKS - left in place."
        Exit Sub
    End If

    If MsgBox(fso.GetFileName(oldPath) & " was renamed to " & fso.GetFileName(newPath) & _
              " and all references found were verified." & vbCrLf & vbCrLf & _
              "The old file is still on disk. Move it into the """ & BACKUP_FOLDER_NAME & """ folder?" & vbCrLf & vbCrLf & _
              "Yes = move to backup folder" & vbCrLf & _
              "No  = keep it where it is (safest if files OUTSIDE the search folder may still use it)", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Safe Rename - old file") <> vbYes Then
        LogMsg "  Old file kept: " & oldPath
        Exit Sub
    End If

    Dim bdir As String, target As String
    bdir = fso.BuildPath(fso.GetParentFolderName(oldPath), BACKUP_FOLDER_NAME)
    On Error Resume Next
    If Not fso.FolderExists(bdir) Then fso.CreateFolder bdir
    target = fso.BuildPath(bdir, fso.GetFileName(oldPath))
    If fso.FileExists(target) Then
        target = fso.BuildPath(bdir, fso.GetBaseName(oldPath) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & fso.GetExtensionName(oldPath))
    End If
    fso.MoveFile oldPath, target
    If Err.Number = 0 Then
        LogMsg "  Old file moved to " & target
    Else
        LogMsg "  ERROR moving old file: " & Err.Description
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' Helpers
'===============================================================================
Private Function DocTypeFromPath(ByVal path As String) As Long
    Select Case LCase$(fso.GetExtensionName(path))
        Case "sldprt": DocTypeFromPath = swDocPART
        Case "sldasm": DocTypeFromPath = swDocASSEMBLY
        Case "slddrw": DocTypeFromPath = swDocDRAWING
    End Select
End Function

Private Function IsReadOnlyDoc(doc As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    IsReadOnlyDoc = doc.IsOpenedReadOnly Or doc.IsOpenedViewOnly
End Function

Private Sub CloseOpened(opened As Collection)
    Dim d As Variant, doc As SldWorks.ModelDoc2
    For Each d In opened
        Set doc = swApp.GetOpenDocumentByName(d)
        If Not doc Is Nothing Then
            swApp.CloseDoc doc.GetTitle
            LogMsg "  Closed " & d
        End If
    Next d
End Sub

Private Function ListPreview(items As Collection, ByVal maxItems As Long) As String
    Dim i As Long, s As String
    For i = 1 To items.Count
        If i > maxItems Then
            s = s & vbCrLf & "         ... and " & (items.Count - maxItems) & " more (see log)"
            Exit For
        End If
        s = s & vbCrLf & "         " & fso.GetFileName(items(i))
    Next i
    ListPreview = s
End Function

Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    swApp.Frame.SetStatusBarText text
End Sub

Private Sub LogMsg(ByVal text As String)
    logText = logText & text & vbCrLf
    Debug.Print text
End Sub

Private Function FinishLog(topModel As SldWorks.ModelDoc2) As String
    If Not WRITE_LOG_FILE Then Exit Function
    Dim p As String, ts As Object
    p = fso.BuildPath(fso.GetParentFolderName(topModel.GetPathName), _
                      "SafeRename_Log_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt")
    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True)
    ts.Write logText
    ts.Close
    If Err.Number = 0 Then FinishLog = p
End Function
