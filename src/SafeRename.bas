Attribute VB_Name = "SafeRename"
'===============================================================================
' SafeRename.bas  -  SOLIDWORKS VBA macro (engine + entry point)
'
' Renames part / sub-assembly files used inside the active assembly WITHOUT
' breaking references in the assemblies, drawings and derived parts that use
' them. The user interface lives in the frmSafeRename UserForm; this module
' does the work.
'
' How a rename works:
'   1. Where-used search: every .SLDASM / .SLDDRW / .SLDPRT in a folder is
'      checked with GetDocumentDependencies2 (nothing is opened for the scan).
'   2. All files that reference the target are opened, so SOLIDWORKS holds
'      them in memory and updates their references.
'   3. The file is renamed in memory with ModelDocExtension.RenameDocument
'      (same engine as "Rename Component > Rename Document", SW 2018+), with a
'      Save As fallback for older versions.
'   4. The renamed file and every dependent are saved -> they store the NEW path.
'   5. Every dependent is re-read from disk and verified to point to the new
'      file and no longer to the old one. Everything is logged.
'   6. Only when verification passed, optional clean-up: rename the drawing of
'      the same name, move the old file into a "_RenameBackup" folder.
'
' Not for files managed in SOLIDWORKS PDM - use the vault's own Rename there.
'===============================================================================
Option Explicit

'------------------------------ options (set by the dialog) --------------------
Public Opt_SaveAsFallback As Boolean    ' use Save As when RenameDocument fails
Public Opt_RenameDrawing As Boolean     ' rename drawing with the same name too
Public Opt_CloseOpened As Boolean       ' close dependents the macro opened
Public Opt_BackupOldFile As Boolean     ' move old file to BACKUP_FOLDER_NAME
Public Opt_WriteLog As Boolean          ' write a log file next to the assembly

Public Const BACKUP_FOLDER_NAME As String = "_RenameBackup"

'------------------------------ API constants ----------------------------------
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

'------------------------------ shared state -----------------------------------
Public swApp As SldWorks.SldWorks
Public gUI As Object                    ' the open dialog (receives log/progress)
Private fso As Object
Private logText As String
Private depCache As Object              ' file path -> dependency array
Private skippedText As String           ' components ignored during selection

'===============================================================================
' Entry point
'===============================================================================
Sub main()
    Set swApp = Application.SldWorks
    frmSafeRename.Show vbModeless
End Sub

Public Sub InitEngine()
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set depCache = CreateObject("Scripting.Dictionary")
    depCache.CompareMode = vbTextCompare
    logText = ""
    Opt_SaveAsFallback = True
    Opt_RenameDrawing = True
    Opt_CloseOpened = True
    Opt_BackupOldFile = False
    Opt_WriteLog = True
End Sub

' Forget cached where-used results (call before a new scan).
Public Sub ClearDependencyCache()
    depCache.RemoveAll
End Sub

'===============================================================================
' Active assembly and selection
'===============================================================================
Public Function GetTopAssembly(ByRef problem As String) As SldWorks.ModelDoc2
    Dim m As SldWorks.ModelDoc2
    problem = ""
    Set m = swApp.ActiveDoc
    If m Is Nothing Then
        problem = "Open an assembly first."
    ElseIf m.GetType <> swDocASSEMBLY Then
        problem = "The active document must be an assembly."
    ElseIf m.GetPathName = "" Then
        problem = "Save the assembly once before renaming its components."
    Else
        Set GetTopAssembly = m
    End If
End Function

' Returns a Dictionary: file path -> Component2 (one entry per file).
Public Function CollectSelectedComponents(topModel As SldWorks.ModelDoc2) As Object
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
                SkipComp comp.Name2 & ": virtual component (no file to rename - use Rename Component)"
            ElseIf p = "" Then
                SkipComp comp.Name2 & ": component has no file path"
            ElseIf comp.GetSuppression2 = swComponentSuppressed Then
                SkipComp comp.Name2 & ": component is suppressed - unsuppress it first"
            ElseIf Not result.Exists(p) Then
                result.Add p, comp
            End If
        End If
    Next i

    Set CollectSelectedComponents = result
End Function

Private Sub SkipComp(ByVal msg As String)
    skippedText = skippedText & "  - " & msg & vbCrLf
    LogMsg "Skipped: " & msg
End Sub

Public Function GetSkippedText() As String
    GetSkippedText = skippedText
End Function

'===============================================================================
' Name validation. Strips a typed extension from newName. Returns "" when OK,
' otherwise a description of the problem. usedPaths: new paths already taken.
'===============================================================================
Public Function ValidateNewName(ByVal oldPath As String, ByRef newName As String, usedPaths As Object) As String
    Dim ext As String, oldBase As String, newPath As String
    ext = fso.GetExtensionName(oldPath)
    oldBase = fso.GetBaseName(oldPath)

    newName = Trim$(newName)
    If LCase$(Right$(newName, Len(ext) + 1)) = "." & LCase$(ext) Then
        newName = Left$(newName, Len(newName) - Len(ext) - 1)
    End If
    newPath = NewPathFor(oldPath, newName)

    If newName = "" Then
        ValidateNewName = "Enter a new name."
    ElseIf newName Like "*[\/:*?""<>|]*" Then
        ValidateNewName = "The name contains invalid characters  \ / : * ? "" < > |"
    ElseIf StrComp(newName, oldBase, vbBinaryCompare) = 0 Then
        ValidateNewName = "The new name is the same as the current name."
    ElseIf StrComp(newName, oldBase, vbTextCompare) = 0 Then
        ValidateNewName = "Changing only upper/lower case is not supported - rename to a temporary name first."
    ElseIf fso.FileExists(newPath) Then
        ValidateNewName = "A file with that name already exists in the folder."
    ElseIf Not swApp.GetOpenDocumentByName(newPath) Is Nothing Then
        ValidateNewName = "A document with that name is already open in SOLIDWORKS."
    ElseIf Not usedPaths Is Nothing Then
        If usedPaths.Exists(newPath) Then ValidateNewName = "This name is used for another file in the list."
    End If
    If ValidateNewName = "" And IsReadOnlyFile(oldPath) Then
        ValidateNewName = "The current file is read-only (checked in to PDM or write-protected)."
    End If
End Function

Public Function NewPathFor(ByVal oldPath As String, ByVal newName As String) As String
    NewPathFor = fso.BuildPath(fso.GetParentFolderName(oldPath), newName & "." & fso.GetExtensionName(oldPath))
End Function

Private Function IsReadOnlyFile(ByVal path As String) As Boolean
    On Error Resume Next
    IsReadOnlyFile = (fso.GetFile(path).Attributes And 1) <> 0
End Function

'===============================================================================
' Where-used search
'===============================================================================
Public Function ListSwFiles(ByVal rootFolder As String, ByVal recurse As Boolean) As Collection
    Dim result As New Collection
    If fso.FolderExists(rootFolder) Then AddSwFiles fso.GetFolder(rootFolder), recurse, result
    Set ListSwFiles = result
End Function

Private Sub AddSwFiles(folder As Object, ByVal recurse As Boolean, result As Collection)
    Dim f As Object, sf As Object, ext As String
    On Error Resume Next            ' skip folders we have no access to
    For Each f In folder.Files
        If Left$(f.Name, 2) <> "~$" Then
            ext = LCase$(fso.GetExtensionName(f.Name))
            If ext = "sldasm" Or ext = "slddrw" Or ext = "sldprt" Then result.Add f.path
        End If
    Next f
    If recurse Then
        For Each sf In folder.SubFolders
            If StrComp(sf.Name, BACKUP_FOLDER_NAME, vbTextCompare) <> 0 Then AddSwFiles sf, True, result
        Next sf
    End If
End Sub

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

' True if filePath directly references target (full-path match).
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

' jobIndex / jobCount are only used for the progress bar.
Public Function FindDependents(ByVal target As String, allFiles As Collection, _
                               ByVal jobIndex As Long, ByVal jobCount As Long) As Collection
    Dim result As New Collection, f As Variant, nameOnly As Boolean, i As Long
    LogMsg "Where-used: " & target
    For Each f In allFiles
        i = i + 1
        If i Mod 10 = 0 Or i = allFiles.Count Then
            SetStatus "Searching references to " & fso.GetFileName(target) & "  (" & i & " / " & allFiles.Count & ")", _
                      ((jobIndex - 1) + i / allFiles.Count) / jobCount
        End If
        If StrComp(f, target, vbTextCompare) <> 0 Then
            If RefersTo(f, target, nameOnly) Then
                result.Add f
                LogMsg "  used by " & f
            ElseIf nameOnly Then
                LogMsg "  WARNING: " & f & " uses a DIFFERENT file with the same name - it will not be changed."
            End If
        End If
    Next f
    LogMsg "  " & result.Count & " referencing file(s)"
    Set FindDependents = result
End Function

'===============================================================================
' Rename one file. Returns True when renamed AND verified.
'===============================================================================
Public Function ProcessJob(topModel As SldWorks.ModelDoc2, comp As SldWorks.Component2, _
                           ByVal oldPath As String, ByVal newName As String, _
                           dependents As Collection) As Boolean
    Dim newPath As String
    newPath = NewPathFor(oldPath, newName)

    LogMsg ""
    LogMsg "=== " & oldPath & "  ->  " & newPath

    Dim opened As New Collection            ' docs this macro opened
    Dim toSave As Object                    ' path -> ModelDoc2 that must be saved
    Set toSave = CreateObject("Scripting.Dictionary")
    toSave.CompareMode = vbTextCompare

    Dim d As Variant, doc As SldWorks.ModelDoc2, errs As Long, warns As Long, n As Long

    '--- open every dependent so SOLIDWORKS updates it in memory ----------------
    For Each d In dependents
        n = n + 1
        SetStatus "Opening " & fso.GetFileName(d) & "  (" & n & " / " & dependents.Count & ")", 0.05 + 0.35 * n / dependents.Count
        Set doc = swApp.GetOpenDocumentByName(d)
        If doc Is Nothing Then
            errs = 0: warns = 0
            Set doc = swApp.OpenDoc6(d, DocTypeFromPath(d), swOpenDocOptions_Silent, "", errs, warns)
            If doc Is Nothing Then
                LogMsg "  ERROR: could not open " & d & " (error " & errs & "). Nothing renamed for this file."
                CloseOpened opened
                Exit Function
            End If
            opened.Add CStr(d)
            LogMsg "  Opened " & d
        End If
        If IsReadOnlyDoc(doc) Then
            LogMsg "  ERROR: " & d & " is read-only. Nothing renamed for this file."
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
                        LogMsg "  WARNING: suppressed instance " & c.Name2 & " - its reference is checked during verification."
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
    SetStatus "Renaming " & fso.GetFileName(oldPath), 0.45
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

    If Not renamed And Opt_SaveAsFallback Then
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
    SetStatus "Saving " & fso.GetFileName(topModel.GetPathName), 0.55
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
    n = 0
    For pass = 1 To 2                       ' models first, drawings last
        For Each key In toSave.keys
            Set doc = toSave(key)
            If (pass = 1 And doc.GetType <> swDocDRAWING) Or (pass = 2 And doc.GetType = swDocDRAWING) Then
                n = n + 1
                SetStatus "Saving " & fso.GetFileName(key) & "  (" & n & " / " & toSave.Count & ")", 0.55 + 0.3 * n / toSave.Count
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
    SetStatus "Verifying references on disk", 0.9
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
    If Opt_CloseOpened Then CloseOpened opened

    If allOk Then
        LogMsg "  VERIFIED: all " & savedPaths.Count & " referencing file(s) now point to " & fso.GetFileName(newPath)
        If Opt_RenameDrawing Then RenameMatchingDrawings oldPath, newPath, dependents
        HandleOldFile oldPath
    Else
        LogMsg "  Verification FAILED - old file kept untouched. Check the entries above."
    End If

    SetStatus "Done: " & fso.GetFileName(oldPath), 1
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
' Optional clean-up (only after successful verification)
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
            Else
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

Private Sub HandleOldFile(ByVal oldPath As String)
    If Not fso.FileExists(oldPath) Then
        LogMsg "  Old file no longer exists on disk (renamed by SOLIDWORKS)."
        Exit Sub
    End If
    If Not Opt_BackupOldFile Then
        LogMsg "  Old file kept: " & oldPath
        Exit Sub
    End If
    If Not swApp.GetOpenDocumentByName(oldPath) Is Nothing Then
        LogMsg "  Old file is still open in SOLIDWORKS - left in place."
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

' fraction: 0..1 for the progress bar, negative = leave the bar as it is.
Public Sub SetStatus(ByVal text As String, Optional ByVal fraction As Double = -1)
    On Error Resume Next
    swApp.Frame.SetStatusBarText IIf(text = "", "", "Safe Rename: " & text)
    If Not gUI Is Nothing Then gUI.ShowProgress text, fraction
    DoEvents
End Sub

Public Sub LogMsg(ByVal text As String)
    logText = logText & text & vbCrLf
    Debug.Print text
    On Error Resume Next
    If Not gUI Is Nothing Then gUI.AppendLog text
End Sub

' Writes the log next to the top assembly. Returns its path ("" if not written).
Public Function SaveLog(topModel As SldWorks.ModelDoc2) As String
    If Not Opt_WriteLog Or logText = "" Then Exit Function
    Dim p As String, ts As Object
    p = fso.BuildPath(fso.GetParentFolderName(topModel.GetPathName), _
                      "SafeRename_Log_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt")
    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True)
    ts.Write logText
    ts.Close
    If Err.Number = 0 Then SaveLog = p
    logText = ""
End Function
