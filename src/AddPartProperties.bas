'===============================================================================
' AddPartProperties.bas  -  SOLIDWORKS VBA macro, single module, no UserForm
'
' Adds the custom properties DESCRIPTION, Material, Weight and Size to every
' part used in the active assembly (sub-assemblies included). All questions
' are asked with InputBox / MsgBox.
'
'   - Only local part files in the assembly's folder (and its sub-folders) are
'     changed. Virtual components, PDM files, read-only files, Toolbox parts,
'     suppressed and envelope components are skipped.
'   - Properties go only to the Custom tab (configuration name "").
'   - Weight is linked to SW-Mass and converted to kg when the part's mass unit
'     is not kg. The result is checked against the part's real mass.
'   - Optionally all old properties on the Custom tab are deleted first. They
'     are logged, and restored if writing the part fails.
'
' Install: Tools > Macro > New, save as AddPartProperties.swp. In the VBA editor
' select all the code of the new module, delete it and paste this whole file
' (or File > Import File... this .bas). Then run main.
' Late bound: no extra references are needed.
'
' See docs/AddPartProperties-spec.md.
'===============================================================================
Option Explicit

Private Const TITLE As String = "Add Part Properties"

'------------------------------ property names ---------------------------------
Private Const PROP_DESC As String = "DESCRIPTION"
Private Const PROP_MAT As String = "Material"
Private Const PROP_WEIGHT As String = "Weight"
Private Const PROP_SIZE As String = "Size"

'------------------------------ options (set by the dialogs) -------------------
Private Opt_AddDesc As Boolean
Private Opt_AddMat As Boolean
Private Opt_AddWeight As Boolean
Private Opt_AddSize As Boolean
Private Opt_DeleteOld As Boolean        ' delete every Custom-tab property first
Private Opt_Overwrite As Boolean        ' replace existing values (Keep mode only)
Private Opt_Save As Boolean             ' save each part after writing

Private MatCommon As String             ' material for all parts
Private SizeSameForAll As Boolean       ' True: SizeAll for every part
Private SizeAll As String

'------------------------------ parts (1-based, PartCount entries) --------------
Private PartCount As Long
Private pPath() As String               ' full path of the part file
Private pComp() As Object               ' one instance (Component2) in the top assembly
Private pWhere() As String              ' sub-assembly path, for display
Private pConfig() As String             ' referenced configuration
Private pOldDesc() As String            ' existing values ("" = none)
Private pOldMat() As String
Private pOldSize() As String
Private pHasDesc() As Boolean
Private pHasMat() As Boolean
Private pHasSize() As Boolean
Private pHasWeight() As Boolean
Private pPropCount() As Long            ' number of properties on the Custom tab
Private pDesc() As String               ' value typed for the part
Private pDescSet() As Boolean           ' False = keep existing (skipped)
Private pSize() As String
Private pSizeSet() As Boolean
Private pMatExc() As Boolean            ' True = exception, uses pMat
Private pMat() As String
Private pResult() As String             ' result after writing

'------------------------------ API constants ----------------------------------
Private Const swDocASSEMBLY As Long = 2
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swComponentSuppressed As Long = 0
Private Const swComponentLightweight As Long = 1
Private Const swComponentFullyResolved As Long = 2
Private Const swComponentFullyLightweight As Long = 4
Private Const swToolboxStandardPart As Long = 1
Private Const CI_TEXT As Long = 30              ' swCustomInfoText
Private Const CP_DELETE_AND_ADD As Long = 1     ' swCustomPropertyDeleteAndAdd
Private Const ADD_OK As Long = 0                ' swCustomInfoAddResult_AddedOrChanged

'------------------------------ shared state -----------------------------------
Private swApp As Object
Private TopDoc As Object                ' the active assembly
Private fso As Object
Private logText As String
Private skippedText As String
Private skippedCount As Long
Private pdmFolders As Object            ' folder -> True/False (cache)
Private pdmVault As Object
Private pdmVaultTried As Boolean

'===============================================================================
' Entry point
'===============================================================================
Sub main()
    Dim problem As String
    Set swApp = Application.SldWorks
    InitEngine
    Set TopDoc = GetTopAssembly(problem)
    If TopDoc Is Nothing Then
        MsgBox problem, vbExclamation, TITLE
        Exit Sub
    End If

    SetStatus "Collecting parts..."
    CollectParts TopDoc
    SetStatus ""
    If PartCount = 0 Then
        MsgBox "No part in this assembly can be changed." & vbCrLf & vbCrLf & _
               IIf(skippedText = "", "The assembly has no parts.", "Skipped:" & vbCrLf & Clip(skippedText, 800)), _
               vbExclamation, TITLE
        Exit Sub
    End If

    If Not AskProperties() Then Exit Sub
    If Not AskDeleteOld() Then Exit Sub
    If Not (Opt_AddDesc Or Opt_AddMat Or Opt_AddWeight Or Opt_AddSize Or Opt_DeleteOld) Then
        MsgBox "Nothing to do: no property chosen and old properties are kept.", vbInformation, TITLE
        Exit Sub
    End If
    If Opt_AddMat Then
        If Not AskMaterial() Then Exit Sub
    End If
    If Opt_AddSize Then
        If Not AskSize() Then Exit Sub
    End If
    If Opt_AddDesc Or (Opt_AddSize And Not SizeSameForAll) Then
        If Not AskPerPart() Then Exit Sub
    End If
    If Not AskOverwrite() Then Exit Sub
    If Not ConfirmSummary() Then Exit Sub

    ApplyAll
End Sub

Private Sub InitEngine()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pdmFolders = CreateObject("Scripting.Dictionary")
    pdmFolders.CompareMode = vbTextCompare
    Set pdmVault = Nothing
    pdmVaultTried = False
    logText = ""
    Opt_AddDesc = True
    Opt_AddMat = True
    Opt_AddWeight = True
    Opt_AddSize = True
    Opt_DeleteOld = False
    Opt_Overwrite = True
    Opt_Save = True
    MatCommon = ""
    SizeSameForAll = True
    SizeAll = "-"
End Sub

'===============================================================================
' Dialogs. Each returns False when the user cancels (nothing is changed then).
'===============================================================================
' 1. Which of the four properties to add.
Private Function AskProperties() As Boolean
    Dim prompt As String, ans As String
    prompt = PartCount & " part file(s) found (each file once, sub-assemblies included)." & vbCrLf
    If skippedCount > 0 Then prompt = prompt & skippedCount & " skipped:" & vbCrLf & Clip(skippedText, 300)
    prompt = prompt & vbCrLf & "Which properties to add? Type their numbers:" & vbCrLf & _
             "  1 = DESCRIPTION  (typed for each part)" & vbCrLf & _
             "  2 = Material  (one value, with exceptions)" & vbCrLf & _
             "  3 = Weight  (SW-Mass in kg, automatic)" & vbCrLf & _
             "  4 = Size  (typed)" & vbCrLf & _
             "1234 = all four.  0 = none (only delete old properties)."
    Do
        ans = InputBox(prompt, TITLE, "1234")
        If StrPtr(ans) = 0 Then Exit Function
        ans = Trim$(ans)
        If ans <> "" And Not ans Like "*[!0-4 ,]*" Then Exit Do
        MsgBox "Type digits 1 to 4, for example 1234 or 23. Type 0 for none.", vbExclamation, TITLE
    Loop
    Opt_AddDesc = InStr(ans, "1") > 0
    Opt_AddMat = InStr(ans, "2") > 0
    Opt_AddWeight = InStr(ans, "3") > 0
    Opt_AddSize = InStr(ans, "4") > 0
    AskProperties = True
End Function

' 2. Keep or delete the old Custom-tab properties.
Private Function AskDeleteOld() As Boolean
    Dim i As Long, n As Long, r As VbMsgBoxResult
    For i = 1 To PartCount
        If pPropCount(i) > 0 Then n = n + 1
    Next i
    If n = 0 Then
        Opt_DeleteOld = False
        AskDeleteOld = True
        Exit Function
    End If
    r = MsgBox(n & " of the " & PartCount & " parts already have custom properties on the Custom tab." & vbCrLf & vbCrLf & _
               "Delete ALL existing custom properties first?" & vbCrLf & vbCrLf & _
               "Yes = delete all old properties, then add the chosen ones" & vbCrLf & _
               "        (cannot be undone - the deleted values are written to the log)" & vbCrLf & _
               "No = keep existing properties" & vbCrLf & _
               "Cancel = stop" & vbCrLf & vbCrLf & _
               "Configuration-specific properties are never changed.", _
               vbYesNoCancel + vbQuestion + vbDefaultButton2, TITLE)
    If r = vbCancel Then Exit Function
    If r = vbYes Then
        If MsgBox("All properties on the Custom tab of " & n & " part(s) will be DELETED." & vbCrLf & _
                  "Are you sure?", vbOKCancel + vbExclamation + vbDefaultButton2, TITLE) <> vbOK Then Exit Function
    End If
    Opt_DeleteOld = (r = vbYes)
    AskDeleteOld = True
End Function

' 3. Material: one value for all parts, optionally with exceptions.
Private Function AskMaterial() As Boolean
    Dim v As String, r As VbMsgBoxResult
    Do
        v = InputBox("Material for all parts (written as typed):", TITLE, MatCommon)
        If StrPtr(v) = 0 Then Exit Function
        v = Trim$(v)
        If v <> "" Then Exit Do
        MsgBox "Type the material for all parts.", vbExclamation, TITLE
    Loop
    MatCommon = v

    r = MsgBox("Material = """ & MatCommon & """" & vbCrLf & vbCrLf & _
               "Yes = apply to all " & PartCount & " parts" & vbCrLf & _
               "No = apply to all parts, with exceptions (pick parts that get a different material)" & vbCrLf & _
               "Cancel = stop", vbYesNoCancel + vbQuestion, TITLE)
    If r = vbCancel Then Exit Function
    If r = vbNo Then
        If Not AskExceptions() Then Exit Function
    End If
    AskMaterial = True
End Function

' Lists the parts page by page; the user types the numbers of the exceptions,
' then the material of each exception.
Private Function AskExceptions() As Boolean
    Const PER_PAGE As Long = 15
    Dim pages As Long, pg As Long, first As Long, last As Long, i As Long
    Dim prompt As String, ans As String, picked() As Boolean, nPicked As Long, bad As String

    ReDim picked(1 To PartCount)
    pages = (PartCount + PER_PAGE - 1) \ PER_PAGE
    pg = 1
    Do While pg <= pages
        first = (pg - 1) * PER_PAGE + 1
        last = first + PER_PAGE - 1
        If last > PartCount Then last = PartCount
        prompt = "Material exceptions - parts " & first & " to " & last & " of " & PartCount & ":" & vbCrLf
        For i = first To last
            prompt = prompt & IIf(picked(i), " *", "  ") & i & "  " & Clip(fso.GetFileName(pPath(i)), 40) & vbCrLf
        Next i
        prompt = Clip(prompt, 850) & vbCrLf & "Type the numbers of the parts with a different material " & _
                 "(e.g. 3, 7, 9-11), or leave empty and press OK for the next page."
        ans = InputBox(prompt, TITLE & " - page " & pg & " of " & pages, "")
        If StrPtr(ans) = 0 Then Exit Function
        bad = MarkNumbers(ans, picked)
        If bad <> "" Then
            MsgBox "Not understood: " & bad & vbCrLf & "Type part numbers like 3, 7, 9-11.", vbExclamation, TITLE
        Else
            pg = pg + 1
        End If
    Loop

    Dim lastVal As String, n As Long, k As Long, v As String
    For i = 1 To PartCount
        If picked(i) Then nPicked = nPicked + 1
    Next i
    If nPicked = 0 Then
        AskExceptions = True
        Exit Function
    End If

    lastVal = MatCommon
    For i = 1 To PartCount
        If picked(i) Then
            k = k + 1
            SelectPart TopDoc, i, True
            v = InputBox("Exception " & k & " of " & nPicked & ": " & fso.GetFileName(pPath(i)) & _
                         IIf(pWhere(i) = "", "", vbCrLf & "In: " & pWhere(i)) & vbCrLf & vbCrLf & _
                         "Material for this part (empty = use """ & MatCommon & """):", TITLE, lastVal)
            If StrPtr(v) = 0 Then
                EndSelection
                Exit Function
            End If
            v = Trim$(v)
            If v <> "" And v <> MatCommon Then
                pMatExc(i) = True
                pMat(i) = v
                lastVal = v
                n = n + 1
            End If
        End If
    Next i
    EndSelection
    LogMsg n & " material exception(s)."
    AskExceptions = True
End Function

' Marks the numbers in text ("3, 7 9-11") in picked(). Returns the parts it did
' not understand ("" = all fine).
Private Function MarkNumbers(ByVal text As String, picked() As Boolean) As String
    Dim tokens() As String, t As String, k As Long, j As Long, dash As Long
    Dim a As Long, b As Long, bad As String
    text = Replace(Replace(text, ";", " "), ",", " ")
    tokens = Split(Trim$(text), " ")
    For k = LBound(tokens) To UBound(tokens)
        t = Trim$(tokens(k))
        If t <> "" Then
            a = 0: b = 0
            dash = InStr(t, "-")
            If dash > 1 And dash < Len(t) Then
                If IsDigits(Left$(t, dash - 1)) And IsDigits(Mid$(t, dash + 1)) Then
                    a = CLng(Left$(t, dash - 1))
                    b = CLng(Mid$(t, dash + 1))
                End If
            ElseIf IsDigits(t) Then
                a = CLng(t)
                b = a
            End If
            If a < 1 Or b < a Or b > UBound(picked) Then
                bad = bad & " " & t
            Else
                For j = a To b
                    picked(j) = True
                Next j
            End If
        End If
    Next k
    MarkNumbers = Trim$(bad)
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim k As Long
    If s = "" Or Len(s) > 6 Then Exit Function
    For k = 1 To Len(s)
        If Mid$(s, k, 1) < "0" Or Mid$(s, k, 1) > "9" Then Exit Function
    Next k
    IsDigits = True
End Function

' 4. Size: same value for all parts, or typed per part.
Private Function AskSize() As Boolean
    Dim r As VbMsgBoxResult, v As String
    r = MsgBox("Size" & vbCrLf & vbCrLf & _
               "Yes = same value for all parts (default ""-"")" & vbCrLf & _
               "No = type the size for each part" & vbCrLf & _
               "Cancel = stop", vbYesNoCancel + vbQuestion, TITLE)
    If r = vbCancel Then Exit Function
    SizeSameForAll = (r = vbYes)
    If SizeSameForAll Then
        v = InputBox("Size for all parts:", TITLE, "-")
        If StrPtr(v) = 0 Then Exit Function
        SizeAll = Trim$(v)
        If SizeAll = "" Then SizeAll = "-"
    End If
    AskSize = True
End Function

' 5. One question per part for DESCRIPTION (and Size when typed per part).
'    Empty = keep the existing value, "<" = back, Cancel = skip the rest or stop.
Private Function AskPerPart() As Boolean
    Dim i As Long, act As String
    i = 1
    Do While i <= PartCount
        SelectPart TopDoc, i, True
        act = "next"
        If Opt_AddDesc Then act = AskPartValue(i, PROP_DESC)
        If act = "next" And Opt_AddSize And Not SizeSameForAll Then act = AskPartValue(i, PROP_SIZE)
        Select Case act
            Case "next": i = i + 1
            Case "back": If i > 1 Then i = i - 1
            Case "rest": Exit Do
            Case "stop"
                EndSelection
                Exit Function
        End Select
    Loop
    EndSelection
    AskPerPart = True
End Function

' Asks one value of part i. Returns "next", "back", "rest" (skip remaining) or "stop".
Private Function AskPartValue(ByVal i As Long, ByVal prop As String) As String
    Dim prompt As String, cur As String, existing As String, has As Boolean, v As String, r As VbMsgBoxResult, j As Long
    If prop = PROP_DESC Then
        cur = pDesc(i): existing = pOldDesc(i): has = pHasDesc(i)
    Else
        cur = pSize(i): existing = pOldSize(i): has = pHasSize(i)
    End If
    prompt = "Part " & i & " of " & PartCount & ":  " & fso.GetFileName(pPath(i)) & vbCrLf & _
             IIf(pWhere(i) = "", "", "In: " & pWhere(i) & vbCrLf) & _
             "Configuration: " & pConfig(i) & vbCrLf & _
             "Existing " & prop & ": " & IIf(has, """" & Clip(existing, 120) & """", "(none)") & vbCrLf & vbCrLf & _
             "Type the " & prop & " and press OK." & vbCrLf & _
             "Empty = keep the existing value.   < = back to the previous part." & vbCrLf & _
             "Cancel = skip the remaining parts, or stop."
    Do
        v = InputBox(prompt, TITLE & " - " & prop, cur)
        If StrPtr(v) <> 0 Then Exit Do
        r = MsgBox("Yes = skip the remaining parts (keep their " & prop & ") and continue" & vbCrLf & _
                   "No = stop the macro, nothing is changed" & vbCrLf & _
                   "Cancel = back to this part", vbYesNoCancel + vbQuestion, TITLE)
        If r = vbYes Then
            For j = i To PartCount
                KeepPart j
            Next j
            AskPartValue = "rest"
            Exit Function
        ElseIf r = vbNo Then
            AskPartValue = "stop"
            Exit Function
        End If
    Loop

    v = Trim$(v)
    If v = "<" Then
        AskPartValue = "back"
        Exit Function
    End If
    If prop = PROP_DESC Then
        pDescSet(i) = (v <> "")
        pDesc(i) = IIf(v = "", pOldDesc(i), v)
    ElseIf v = "" Then
        ' keep the existing size; a part without one gets "-"
        pSizeSet(i) = Not pHasSize(i)
        pSize(i) = IIf(pHasSize(i), pOldSize(i), "-")
    Else
        pSizeSet(i) = True
        pSize(i) = v
    End If
    AskPartValue = "next"
End Function

Private Sub KeepPart(ByVal i As Long)
    pDescSet(i) = False
    pDesc(i) = pOldDesc(i)
    pSizeSet(i) = False
    pSize(i) = IIf(pHasSize(i), pOldSize(i), "-")
End Sub

' 6. Overwrite existing values? Only asked in Keep mode when there are any.
Private Function AskOverwrite() As Boolean
    Dim i As Long, n As Long, r As VbMsgBoxResult
    AskOverwrite = True
    Opt_Overwrite = True
    If Opt_DeleteOld Then Exit Function
    For i = 1 To PartCount
        If (Opt_AddDesc And pDescSet(i) And pHasDesc(i)) Or (Opt_AddMat And pHasMat(i)) _
           Or (Opt_AddSize And pHasSize(i)) Or (Opt_AddWeight And pHasWeight(i)) Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    r = MsgBox(n & " part(s) already have some of these properties." & vbCrLf & vbCrLf & _
               "Yes = overwrite the existing values" & vbCrLf & _
               "No = keep the existing values, only add missing properties" & vbCrLf & _
               "Cancel = stop", vbYesNoCancel + vbQuestion, TITLE)
    If r = vbCancel Then AskOverwrite = False
    Opt_Overwrite = (r = vbYes)
End Function

' 7. Summary of what will be written; OK writes.
Private Function ConfirmSummary() As Boolean
    Dim s As String, i As Long, shown As Long, ln As String
    Const MAX_LINES As Long = 12

    s = "Write to " & PartCount & " part(s) and save them?" & vbCrLf
    If Opt_DeleteOld Then s = s & "ALL old Custom-tab properties are deleted first." & vbCrLf
    s = s & "Add: " & IIf(Opt_AddDesc, PROP_DESC & "  ", "") & IIf(Opt_AddMat, PROP_MAT & "  ", "") & _
            IIf(Opt_AddWeight, PROP_WEIGHT & " (kg)  ", "") & IIf(Opt_AddSize, PROP_SIZE, "") & vbCrLf
    If Not Opt_DeleteOld Then s = s & "Existing values: " & IIf(Opt_Overwrite, "overwritten", "kept") & vbCrLf
    s = s & vbCrLf

    For i = 1 To PartCount
        If shown >= MAX_LINES Then
            s = s & "... and " & (PartCount - shown) & " more (all listed in the log)" & vbCrLf
            Exit For
        End If
        ln = Clip(fso.GetFileName(pPath(i)), 30) & ":"
        If Opt_AddDesc Then ln = ln & "  D=" & Clip(PlannedText(i, PROP_DESC), 25)
        If Opt_AddMat Then ln = ln & "  M=" & Clip(PlannedText(i, PROP_MAT), 15)
        If Opt_AddSize Then ln = ln & "  S=" & Clip(PlannedText(i, PROP_SIZE), 12)
        s = s & ln & vbCrLf
        shown = shown + 1
    Next i

    LogMsg ""
    LogMsg "Planned:"
    For i = 1 To PartCount
        LogMsg "  " & fso.GetFileName(pPath(i)) & " | " & PROP_DESC & ": " & PlannedText(i, PROP_DESC) & _
               " | " & PROP_MAT & ": " & PlannedText(i, PROP_MAT) & " | " & PROP_WEIGHT & ": " & PlannedText(i, PROP_WEIGHT) & _
               " | " & PROP_SIZE & ": " & PlannedText(i, PROP_SIZE) & " | " & PlannedNote(i)
    Next i

    ConfirmSummary = (MsgBox(Clip(s, 1000), vbOKCancel + IIf(Opt_DeleteOld, vbExclamation, vbQuestion), TITLE) = vbOK)
End Function

' 8. Write every part, rebuild, save the log, report.
Private Sub ApplyAll()
    Dim i As Long, okCount As Long, failCount As Long, logPath As String, fails As String

    LogMsg ""
    LogMsg "Started " & Now
    LogMsg "Add: " & IIf(Opt_AddDesc, PROP_DESC & " ", "") & IIf(Opt_AddMat, PROP_MAT & " ", "") & _
           IIf(Opt_AddWeight, PROP_WEIGHT & " ", "") & IIf(Opt_AddSize, PROP_SIZE, "") & _
           " | old properties: " & IIf(Opt_DeleteOld, "DELETE", "keep") & _
           " | overwrite: " & Opt_Overwrite & " | save: " & Opt_Save
    For i = 1 To PartCount
        SetStatus "Writing " & fso.GetFileName(pPath(i)) & "  (" & i & " / " & PartCount & ")"
        If WritePart(i) Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
            fails = fails & "  - " & fso.GetFileName(pPath(i)) & ": " & pResult(i) & vbCrLf
        End If
    Next i
    SetStatus ""

    On Error Resume Next
    TopDoc.EditRebuild3                 ' refresh the BOM
    On Error GoTo 0
    LogMsg ""
    LogMsg "Finished " & Now & ": " & okCount & " updated, " & failCount & " failed, " & skippedCount & " skipped."
    logPath = SaveLog(TopDoc)

    If MsgBox(okCount & " part(s) updated, " & failCount & " failed, " & skippedCount & " skipped." & vbCrLf & _
              IIf(fails = "", "", vbCrLf & "Failed:" & vbCrLf & Clip(fails, 500)) & _
              IIf(logPath = "", "", vbCrLf & "Log: " & logPath & vbCrLf & vbCrLf & "Open the log?"), _
              IIf(failCount > 0, vbExclamation, vbInformation) + IIf(logPath = "", vbOKOnly, vbYesNo), TITLE) = vbYes Then
        Shell "notepad.exe """ & logPath & """", vbNormalFocus
    End If
End Sub

Private Sub EndSelection()
    On Error Resume Next
    TopDoc.ClearSelection2 True
    TopDoc.ViewZoomtofit2
End Sub

Private Function Clip(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then Clip = Left$(s, n - 3) & "..." Else Clip = s
End Function

'===============================================================================
' Active assembly and its parts
'===============================================================================
Private Function GetTopAssembly(ByRef problem As String) As Object
    Dim m As Object
    problem = ""
    Set m = swApp.ActiveDoc
    If m Is Nothing Then
        problem = "Open an assembly first."
    ElseIf m.GetType <> swDocASSEMBLY Then
        problem = "Open an assembly first. The active document must be an assembly."
    ElseIf m.GetPathName = "" Then
        problem = "Save the assembly once before adding properties to its parts."
    Else
        Set GetTopAssembly = m
    End If
End Function

' Fills the p...() arrays with every part file that will be processed.
Private Sub CollectParts(topModel As Object)
    Dim seen As Object, pending As Object, vComps As Variant, k As Long, c As Object
    Dim p As String, asmFolder As String, reason As String, doc As Object

    PartCount = 0
    skippedText = ""
    skippedCount = 0
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set pending = CreateObject("Scripting.Dictionary")      ' file -> skipped instance
    pending.CompareMode = vbTextCompare
    asmFolder = fso.GetParentFolderName(topModel.GetPathName)
    LogMsg "Assembly: " & topModel.GetPathName
    LogMsg "Folder:   " & asmFolder

    Dim swAssy As Object
    Set swAssy = topModel
    vComps = swAssy.GetComponents(False)            ' all levels
    If Not IsArray(vComps) Then Exit Sub

    For k = LBound(vComps) To UBound(vComps)
        Set c = vComps(k)
        p = c.GetPathName
        reason = InstanceSkipReason(c)
        If LCase$(fso.GetExtensionName(p)) <> "sldprt" Then
            ' sub-assembly: not processed itself, its parts are in the list too
        ElseIf seen.Exists(p) Then
            ' same file, other instance
        ElseIf reason <> "" Then
            ' this instance cannot be used; another instance of the file may be
            If Not pending.Exists(p) Then pending.Add p, c.Name2 & vbTab & reason
        Else
            seen.Add p, True
            reason = SkipReason(p, asmFolder)
            If reason <> "" Then
                Skip c.Name2, p, reason
            Else
                Select Case c.GetSuppression2
                    Case swComponentLightweight, swComponentFullyLightweight
                        c.SetSuppression2 swComponentFullyResolved
                End Select
                Set doc = c.GetModelDoc2
                If doc Is Nothing Then
                    Skip c.Name2, p, "part could not be loaded"
                ElseIf IsReadOnlyDoc(doc) Then
                    Skip c.Name2, p, "opened read-only"
                ElseIf ToolboxType(doc) = swToolboxStandardPart Then
                    Skip c.Name2, p, "Toolbox part"
                Else
                    AddPart c, p, doc
                End If
            End If
        End If
    Next k

    Dim key As Variant, parts() As String
    For Each key In pending.keys
        If Not seen.Exists(key) Then
            parts = Split(pending(key), vbTab)
            Skip parts(0), CStr(key), parts(1)
        End If
    Next key

    LogMsg PartCount & " part(s) to process, " & skippedCount & " skipped."
End Sub

' Why this instance cannot be used, or "".
Private Function InstanceSkipReason(c As Object) As String
    If c.IsVirtual Then
        InstanceSkipReason = "virtual component"
    ElseIf c.GetSuppression2 = swComponentSuppressed Then
        InstanceSkipReason = "suppressed"
    ElseIf IsEnvelope(c) Then
        InstanceSkipReason = "envelope component"
    End If
End Function

' Why part file p is not processed, or "" when it is.
Private Function SkipReason(ByVal p As String, ByVal asmFolder As String) As String
    If StrComp(Left$(p, Len(asmFolder) + 1), asmFolder & "\", vbTextCompare) <> 0 Then
        SkipReason = "outside the assembly folder"
    ElseIf Not fso.FileExists(p) Then
        SkipReason = "file not found"
    ElseIf IsPdmFile(p) Then
        SkipReason = "file is in a SOLIDWORKS PDM vault"
    ElseIf IsReadOnlyFile(p) Then
        SkipReason = "file is read-only (checked in to PDM or write-protected)"
    End If
End Function

Private Sub AddPart(c As Object, ByVal p As String, doc As Object)
    Dim i As Long, cpm As Object, names As Variant
    PartCount = PartCount + 1
    i = PartCount
    ReDim Preserve pPath(1 To i): ReDim Preserve pComp(1 To i)
    ReDim Preserve pWhere(1 To i): ReDim Preserve pConfig(1 To i)
    ReDim Preserve pOldDesc(1 To i): ReDim Preserve pOldMat(1 To i): ReDim Preserve pOldSize(1 To i)
    ReDim Preserve pHasDesc(1 To i): ReDim Preserve pHasMat(1 To i)
    ReDim Preserve pHasSize(1 To i): ReDim Preserve pHasWeight(1 To i)
    ReDim Preserve pPropCount(1 To i)
    ReDim Preserve pDesc(1 To i): ReDim Preserve pDescSet(1 To i)
    ReDim Preserve pSize(1 To i): ReDim Preserve pSizeSet(1 To i)
    ReDim Preserve pMatExc(1 To i): ReDim Preserve pMat(1 To i)
    ReDim Preserve pResult(1 To i)

    pPath(i) = p
    Set pComp(i) = c
    pWhere(i) = ParentPath(c)
    pConfig(i) = c.ReferencedConfiguration

    Set cpm = doc.Extension.CustomPropertyManager("")
    pHasDesc(i) = ReadProp(cpm, PROP_DESC, pOldDesc(i))
    pHasMat(i) = ReadProp(cpm, PROP_MAT, pOldMat(i))
    pHasSize(i) = ReadProp(cpm, PROP_SIZE, pOldSize(i))
    pHasWeight(i) = ReadProp(cpm, PROP_WEIGHT, "")
    names = cpm.GetNames
    If IsArray(names) Then pPropCount(i) = UBound(names) - LBound(names) + 1

    pDesc(i) = pOldDesc(i)
    pDescSet(i) = False
    pSize(i) = IIf(pHasSize(i), pOldSize(i), "-")
    pSizeSet(i) = False
    pMatExc(i) = False
    pMat(i) = ""
    pResult(i) = ""
End Sub

Private Sub Skip(ByVal compName As String, ByVal p As String, ByVal reason As String)
    skippedCount = skippedCount + 1
    skippedText = skippedText & "  - " & IIf(p = "", compName, fso.GetFileName(p)) & ": " & reason & vbCrLf
    LogMsg "Skipped " & IIf(p = "", compName, p) & ": " & reason
End Sub

' "SubAsm-1 / SubSub-1" for a part inside sub-assemblies, "" at top level.
Private Function ParentPath(c As Object) As String
    Dim parent As Object, s As String
    Set parent = c.GetParent
    Do While Not parent Is Nothing
        s = parent.Name2 & IIf(s = "", "", " / " & s)
        Set parent = parent.GetParent
    Loop
    ParentPath = s
End Function

'===============================================================================
' Skip tests
'===============================================================================
Private Function IsEnvelope(c As Object) As Boolean
    On Error Resume Next
    IsEnvelope = c.IsEnvelope
End Function

Private Function IsReadOnlyFile(ByVal path As String) As Boolean
    On Error Resume Next
    IsReadOnlyFile = (fso.GetFile(path).Attributes And 1) <> 0
End Function

Private Function IsReadOnlyDoc(doc As Object) As Boolean
    On Error Resume Next
    IsReadOnlyDoc = doc.IsOpenedReadOnly Or doc.IsOpenedViewOnly
End Function

Private Function ToolboxType(doc As Object) As Long
    Dim ext As Object
    On Error Resume Next
    Set ext = doc.Extension                 ' late bound: compiles on every version
    ToolboxType = ext.ToolboxPartType
End Function

' True when the file is inside a SOLIDWORKS PDM vault view. Two tests:
'   1. the PDM client API, when installed, knows the vault of a path;
'   2. a vault view's root folder has a desktop.ini that names the PDM
'      (Conisio) shell extension.
' Read-only files are skipped separately, which already covers checked-in files.
Private Function IsPdmFile(ByVal path As String) As Boolean
    Dim vaultName As String
    If Not pdmVaultTried Then
        pdmVaultTried = True
        On Error Resume Next
        Set pdmVault = CreateObject("ConisioLib.EdmVault")
        On Error GoTo 0
    End If
    If Not pdmVault Is Nothing Then
        On Error Resume Next
        vaultName = pdmVault.GetVaultNameFromPath(path)
        On Error GoTo 0
        If vaultName <> "" Then IsPdmFile = True: Exit Function
    End If
    IsPdmFile = IsPdmFolder(fso.GetParentFolderName(path))
End Function

Private Function IsPdmFolder(ByVal folder As String) As Boolean
    Dim result As Boolean, ini As String, ts As Object, text As String
    If folder = "" Then Exit Function
    If pdmFolders.Exists(folder) Then IsPdmFolder = pdmFolders(folder): Exit Function

    ini = fso.BuildPath(folder, "desktop.ini")
    If fso.FileExists(ini) Then
        On Error Resume Next
        Set ts = fso.OpenTextFile(ini, 1, False, -2)    ' ForReading, system default (ANSI/Unicode)
        text = ts.ReadAll
        ts.Close
        On Error GoTo 0
        text = LCase$(text)
        result = InStr(text, "conisio") > 0 Or InStr(text, "epdm") > 0
    End If
    If Not result And Not fso.GetFolder(folder).IsRootFolder Then
        result = IsPdmFolder(fso.GetParentFolderName(folder))
    End If
    pdmFolders.Add folder, result
    IsPdmFolder = result
End Function

'===============================================================================
' Planned values (shown in the summary before writing)
'===============================================================================
' What will happen to property prop of part i. Returns the text for the Review
' summary: the value to write, "(unchanged)", "(kept: ...)", "(deleted)" or "".
Private Function PlannedText(ByVal i As Long, ByVal prop As String) As String
    Dim ticked As Boolean, has As Boolean, oldVal As String, newVal As String, writes As Boolean
    Select Case prop
        Case PROP_DESC
            ticked = Opt_AddDesc: has = pHasDesc(i): oldVal = pOldDesc(i)
            writes = pDescSet(i): newVal = pDesc(i)
        Case PROP_MAT
            ticked = Opt_AddMat: has = pHasMat(i): oldVal = pOldMat(i)
            writes = True: newVal = MaterialFor(i)
        Case PROP_SIZE
            ticked = Opt_AddSize: has = pHasSize(i): oldVal = pOldSize(i)
            writes = True: newVal = SizeFor(i, writes)
        Case PROP_WEIGHT
            ticked = Opt_AddWeight: has = pHasWeight(i)
            writes = True: newVal = "SW-Mass in kg"
    End Select

    If Not ticked Then
        If Opt_DeleteOld And has Then PlannedText = "(deleted)"
    ElseIf Not writes Then
        PlannedText = IIf(has, "(unchanged)", "(not set)")
    ElseIf has And Not Opt_DeleteOld And Not Opt_Overwrite Then
        PlannedText = "(kept" & IIf(oldVal = "", "", ": " & oldVal) & ")"
    Else
        PlannedText = newVal
    End If
End Function

Private Function PlannedNote(ByVal i As Long) As String
    Dim n As Long, s As String
    If Opt_DeleteOld Then
        n = pPropCount(i)
        If n > 0 Then s = n & " old propert" & IIf(n = 1, "y", "ies") & " deleted"
    ElseIf Opt_Overwrite Then
        If (Opt_AddDesc And pDescSet(i) And pHasDesc(i)) Or (Opt_AddMat And pHasMat(i)) _
           Or (Opt_AddSize And pHasSize(i)) Or (Opt_AddWeight And pHasWeight(i)) Then
            s = "existing values overwritten"
        End If
    End If
    PlannedNote = s
End Function

Private Function MaterialFor(ByVal i As Long) As String
    MaterialFor = IIf(pMatExc(i), pMat(i), MatCommon)
End Function

' Size value for part i. writes = False when the user skipped the part.
Private Function SizeFor(ByVal i As Long, ByRef writes As Boolean) As String
    If SizeSameForAll Then
        SizeFor = SizeAll
        writes = True
    Else
        SizeFor = pSize(i)
        writes = pSizeSet(i)
    End If
    If Trim$(SizeFor) = "" Then SizeFor = "-"
End Function

'===============================================================================
' Write one part. Returns True when everything was written (and saved).
'===============================================================================
Private Function WritePart(ByVal i As Long) As Boolean
    Dim doc As Object, cpm As Object
    Dim ok As Boolean, writes As Boolean, v As String
    Dim delNames() As String, delTypes() As Long, delVals() As String, delCount As Long

    LogMsg ""
    LogMsg "=== " & pPath(i)
    Set doc = pComp(i).GetModelDoc2
    If doc Is Nothing Then
        LogMsg "  ERROR: the part is no longer loaded."
        pResult(i) = "FAILED - not loaded"
        Exit Function
    End If
    Set cpm = doc.Extension.CustomPropertyManager("")
    ok = True

    '--- delete old properties (logged first, so they can be restored) ---------
    If Opt_DeleteOld Then
        Dim names As Variant, k As Long, nm As String, raw As String
        names = cpm.GetNames
        If IsArray(names) Then
            For k = LBound(names) To UBound(names)
                nm = CStr(names(k))
                raw = ""
                ReadProp cpm, nm, raw
                delCount = delCount + 1
                ReDim Preserve delNames(1 To delCount)
                ReDim Preserve delTypes(1 To delCount)
                ReDim Preserve delVals(1 To delCount)
                delNames(delCount) = nm
                delTypes(delCount) = cpm.GetType2(nm)
                delVals(delCount) = raw
                LogMsg "  Deleted old property: " & nm & " (type " & delTypes(delCount) & ") = " & raw
                cpm.Delete2 nm
            Next k
        End If
        ' Skip in the per-part dialog keeps the old value: write it back.
        If Opt_AddDesc And Not pDescSet(i) And pHasDesc(i) Then ok = ok And SetText(cpm, PROP_DESC, pOldDesc(i))
        If Opt_AddSize And Not SizeSameForAll And Not pSizeSet(i) And pHasSize(i) Then _
            ok = ok And SetText(cpm, PROP_SIZE, pOldSize(i))
    End If

    '--- write the ticked properties -------------------------------------------
    If Opt_AddDesc And pDescSet(i) Then
        If CanWrite(cpm, PROP_DESC) Then ok = ok And SetText(cpm, PROP_DESC, pDesc(i))
    End If
    If Opt_AddMat Then
        If CanWrite(cpm, PROP_MAT) Then ok = ok And SetText(cpm, PROP_MAT, MaterialFor(i))
    End If
    If Opt_AddSize Then
        v = SizeFor(i, writes)
        If writes Then
            If CanWrite(cpm, PROP_SIZE) Then ok = ok And SetText(cpm, PROP_SIZE, v)
        End If
    End If
    If Opt_AddWeight Then
        If CanWrite(cpm, PROP_WEIGHT) Then ok = ok And SetWeight(doc, cpm, pPath(i))
    End If

    '--- failure: put the deleted properties back, do not save -----------------
    If Not ok Then
        If delCount > 0 Then
            Dim cur As Variant
            cur = cpm.GetNames
            If IsArray(cur) Then
                For k = LBound(cur) To UBound(cur)
                    cpm.Delete2 CStr(cur(k))
                Next k
            End If
            For k = 1 To delCount
                cpm.Add3 delNames(k), delTypes(k), delVals(k), CP_DELETE_AND_ADD
            Next k
            LogMsg "  Old properties restored. The part was not saved."
        End If
        pResult(i) = "FAILED - see log"
        Exit Function
    End If

    '--- save -------------------------------------------------------------------
    If Opt_Save Then
        Dim errs As Long, warns As Long
        doc.SetSaveFlag
        If doc.Save3(swSaveAsOptions_Silent, errs, warns) Then
            LogMsg "  Saved."
        Else
            LogMsg "  ERROR: save failed (error " & errs & "). The changes are only in memory."
            pResult(i) = "FAILED - not saved"
            Exit Function
        End If
    End If
    pResult(i) = IIf(Opt_Save, "Updated", "Updated (not saved)")
    WritePart = True
End Function

' False (and logged) when the property exists and must be kept.
Private Function CanWrite(cpm As Object, ByVal prop As String) As Boolean
    If Opt_DeleteOld Or Opt_Overwrite Then
        CanWrite = True
    ElseIf ReadProp(cpm, prop, "") Then
        LogMsg "  " & prop & ": exists, kept (overwrite is off)."
    Else
        CanWrite = True
    End If
End Function

' Writes a Text property and reads it back.
Private Function SetText(cpm As Object, ByVal prop As String, ByVal value As String) As Boolean
    Dim rc As Long, back As String
    value = Trim$(value)
    rc = cpm.Add3(prop, CI_TEXT, value, CP_DELETE_AND_ADD)
    If rc <> ADD_OK Then
        LogMsg "  ERROR: " & prop & " could not be written (code " & rc & ")."
        Exit Function
    End If
    ReadProp cpm, prop, back
    If back <> value Then
        LogMsg "  ERROR: " & prop & " reads back as """ & back & """, expected """ & value & """."
        Exit Function
    End If
    LogMsg "  " & prop & " = " & value
    SetText = True
End Function

'===============================================================================
' Weight: "SW-Mass@<part>.SLDPRT", converted to kg
'
' SW-Mass is shown in the part's mass unit. The unit is found by comparing the
' resolved SW-Mass with the real mass in kg (MassProperty.Mass is always kg).
' For a non-kg part an equation such as "SW-Mass@Part.SLDPRT"/1000 is written
' and checked; both the plain form and the "=" form of the equation are tried.
'===============================================================================
Private Function SetWeight(doc As Object, cpm As Object, ByVal path As String) As Boolean
    Dim link As String, massKg As Double, shown As String
    Dim factor As Double, unitName As String, op As String, candidates As Variant, k As Long, expr As String

    link = """SW-Mass@" & fso.GetFileName(path) & """"
    If Not WriteWeight(cpm, link) Then Exit Function

    massKg = RealMassKg(doc)
    shown = ResolvedValue(cpm, PROP_WEIGHT)
    If massKg <= 0 Then
        LogMsg "  Weight = " & link & "  (mass is 0 - check the part's material; unit could not be checked)"
        SetWeight = True
        Exit Function
    End If

    UnitOf ToNumber(shown), Decimals(shown), massKg, factor, unitName, op
    If factor = 1 Then
        LogMsg "  Weight = " & link & "  -> " & shown & " kg"
        SetWeight = True
        Exit Function
    End If
    If factor = 0 Then
        LogMsg "  WARNING: Weight = " & link & " shows " & shown & ", which does not match the mass " & _
               Format$(massKg, "0.000") & " kg in any known unit. Check the part's mass unit."
        SetWeight = True
        Exit Function
    End If

    candidates = Array(link & op, "=" & link & op)
    For k = LBound(candidates) To UBound(candidates)
        expr = candidates(k)
        If WriteWeight(cpm, expr) Then
            shown = ResolvedValue(cpm, PROP_WEIGHT)
            If IsClose(ToNumber(shown), massKg, Decimals(shown)) Then
                LogMsg "  Weight = " & expr & "  -> " & shown & " kg (part mass unit: " & unitName & ")"
                SetWeight = True
                Exit Function
            End If
        End If
    Next k

    ' The conversion did not evaluate: keep the plain link and say so.
    WriteWeight cpm, link
    LogMsg "  WARNING: the part's mass unit is " & unitName & " and the kg conversion equation did not evaluate." & _
           " Weight is linked to SW-Mass and shows " & unitName & "." & _
           " Set Tools > Options > Document Properties > Units > Mass to kilograms in this part."
    SetWeight = True
End Function

Private Function WriteWeight(cpm As Object, ByVal expr As String) As Boolean
    Dim rc As Long
    rc = cpm.Add3(PROP_WEIGHT, CI_TEXT, expr, CP_DELETE_AND_ADD)
    If rc <> ADD_OK Then
        LogMsg "  ERROR: Weight could not be written (code " & rc & ")."
    Else
        WriteWeight = True
    End If
End Function

' Finds which unit the displayed mass is in. factor = displayed / kg
' (1 = kg, 0 = unknown); op is the expression suffix that converts to kg.
Private Sub UnitOf(ByVal shown As Double, ByVal dec As Long, ByVal massKg As Double, _
                   ByRef factor As Double, ByRef unitName As String, ByRef op As String)
    Dim f As Variant, names As Variant, ops As Variant, k As Long
    f = Array(1#, 1000#, 1000000#, 2.20462262184878, 35.2739619495804)
    names = Array("kg", "g", "mg", "lb", "oz")
    ops = Array("", "/1000", "/1000000", "*0.45359237", "*0.028349523125")
    factor = 0
    For k = LBound(f) To UBound(f)
        If IsClose(shown, massKg * f(k), dec) Then
            factor = f(k): unitName = names(k): op = ops(k)
            Exit Sub
        End If
    Next k
End Sub

' True when shown (rounded to dec decimals) matches expected within rounding
' and 0.5 %.
Private Function IsClose(ByVal shown As Double, ByVal expected As Double, ByVal dec As Long) As Boolean
    Dim tol As Double
    tol = 0.5 * 10 ^ (-dec) + Abs(expected) * 0.005 + 0.0000001
    IsClose = Abs(shown - expected) <= tol
End Function

Private Function RealMassKg(doc As Object) As Double
    Dim ext As Object, mp As Object
    On Error Resume Next
    Set ext = doc.Extension
    Set mp = ext.CreateMassProperty
    If Not mp Is Nothing Then RealMassKg = mp.Mass
End Function

Private Function ResolvedValue(cpm As Object, ByVal prop As String) As String
    Dim raw As String, resolved As String, wasResolved As Boolean
    On Error Resume Next
    cpm.Get5 prop, False, raw, resolved, wasResolved
    ResolvedValue = Trim$(resolved)
End Function

' Locale-independent number parsing ("2.50" or "2,50").
Private Function ToNumber(ByVal s As String) As Double
    ToNumber = Val(Replace(Trim$(s), ",", "."))
End Function

Private Function Decimals(ByVal s As String) As Long
    Dim p As Long
    s = Replace(Trim$(s), ",", ".")
    p = InStr(s, ".")
    If p > 0 Then Decimals = Len(s) - p
End Function

'===============================================================================
' Helpers
'===============================================================================
' Reads the raw value (text expression) of a property. Returns True if it exists.
Private Function ReadProp(cpm As Object, ByVal prop As String, ByRef value As String) As Boolean
    Dim raw As String, resolved As String, wasResolved As Boolean, names As Variant, k As Long
    names = cpm.GetNames
    If Not IsArray(names) Then Exit Function
    For k = LBound(names) To UBound(names)
        If StrComp(CStr(names(k)), prop, vbTextCompare) = 0 Then
            cpm.Get5 CStr(names(k)), False, raw, resolved, wasResolved
            value = raw
            ReadProp = True
            Exit Function
        End If
    Next k
End Function

Private Sub SelectPart(topModel As Object, ByVal i As Long, ByVal zoom As Boolean)
    On Error Resume Next
    topModel.ClearSelection2 True
    If pComp(i).Select4(False, Nothing, False) Then
        If zoom Then topModel.ViewZoomToSelection
    End If
    topModel.GraphicsRedraw2
End Sub

Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    swApp.Frame.SetStatusBarText IIf(text = "", "", "Add Part Properties: " & text)
    DoEvents
End Sub

Private Sub LogMsg(ByVal text As String)
    logText = logText & text & vbCrLf
    Debug.Print text
End Sub

' Writes the log next to the top assembly. Returns its path ("" if not written).
Private Function SaveLog(topModel As Object) As String
    If logText = "" Then Exit Function
    Dim p As String, ts As Object
    p = fso.BuildPath(fso.GetParentFolderName(topModel.GetPathName), _
                      "AddPartProperties_" & Format$(Now, "yyyymmdd_hhnnss") & ".log")
    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True)
    ts.Write logText
    ts.Close
    If Err.Number = 0 Then SaveLog = p
    logText = ""
End Function
