'===============================================================================
' frmSafeRename - dialog of the Safe Rename macro.
'
' All controls are created in code (UserForm_Initialize), so this code can also
' be pasted into an EMPTY UserForm named "frmSafeRename" if you build the macro
' by hand. The work itself is done by the SafeRename module.
'===============================================================================
Option Explicit

Private Const TITLE As String = "Safe Rename"

' controls with events
Private WithEvents lstJobs As MSForms.ListBox
Private WithEvents txtNewName As MSForms.TextBox
Private WithEvents btnAddSel As MSForms.CommandButton
Private WithEvents btnRemove As MSForms.CommandButton
Private WithEvents btnClear As MSForms.CommandButton
Private WithEvents txtFolder As MSForms.TextBox
Private WithEvents btnBrowse As MSForms.CommandButton
Private WithEvents btnAsmFolder As MSForms.CommandButton
Private WithEvents chkSub As MSForms.CheckBox
Private WithEvents btnScan As MSForms.CommandButton
Private WithEvents btnRename As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private WithEvents btnOpenLog As MSForms.CommandButton

' other controls
Private lblExt As MSForms.Label
Private lblNameHint As MSForms.Label
Private lblScan As MSForms.Label
Private chkDrawing As MSForms.CheckBox
Private chkClose As MSForms.CheckBox
Private chkFallback As MSForms.CheckBox
Private chkLog As MSForms.CheckBox
Private optKeep As MSForms.OptionButton
Private optBackup As MSForms.OptionButton
Private lblStatus As MSForms.Label
Private pbBack As MSForms.Label
Private pbFill As MSForms.Label
Private txtLog As MSForms.TextBox

' rename jobs (1-based, jCount entries)
Private jPath() As String       ' current full path
Private jNew() As String        ' new name typed by the user (no extension)
Private jComp() As SldWorks.Component2   ' component in the top assembly
Private jDeps() As Collection   ' where-used result (Nothing = not searched)
Private jStatus() As String     ' "", "Working...", "Renamed", "FAILED - see log"
Private jCount As Long

Private topModel As SldWorks.ModelDoc2
Private fso As Object
Private scanDone As Boolean
Private busy As Boolean
Private suppress As Boolean
Private lastLogPath As String

'===============================================================================
' Form life cycle
'===============================================================================
Private Sub UserForm_Initialize()
    InitEngine
    Set fso = CreateObject("Scripting.FileSystemObject")
    BuildUI
    Set gUI = Me

    chkSub.Value = True
    chkDrawing.Value = Opt_RenameDrawing
    chkClose.Value = Opt_CloseOpened
    chkFallback.Value = Opt_SaveAsFallback
    chkLog.Value = Opt_WriteLog
    optKeep.Value = Not Opt_BackupOldFile
    optBackup.Value = Opt_BackupOldFile

    AddSelection False
    LoadSelectedJob
    SetBusy False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If busy Then Cancel = True
End Sub

Private Sub UserForm_Terminate()
    Set gUI = Nothing
End Sub

'===============================================================================
' Layout
'===============================================================================
Private Sub BuildUI()
    Dim f As Object, c As Object

    Me.Caption = "Safe Rename - rename parts and assemblies without breaking references"
    Me.BackColor = RGB(243, 243, 243)
    Me.Width = 624
    Me.Height = 538

    '--- header ------------------------------------------------------------------
    Set c = NewCtl(Me, "Forms.Label.1", 0, 0, 624, 44)
    c.BackColor = RGB(31, 56, 100)
    Set c = NewCtl(Me, "Forms.Label.1", 12, 4, 400, 22, "Safe Rename")
    c.BackStyle = 0: c.ForeColor = RGB(255, 255, 255): c.Font.Size = 14: c.Font.Bold = True
    Set c = NewCtl(Me, "Forms.Label.1", 12, 26, 600, 14, _
                   "Every assembly, drawing and derived part that uses the file is updated, saved and verified.")
    c.BackStyle = 0: c.ForeColor = RGB(200, 215, 235)

    '--- 1. files ------------------------------------------------------------------
    Set f = NewFrame("1.  Files to rename", 8, 50, 600, 156)
    ColumnHeader f, 9, "Current file"
    ColumnHeader f, 199, "New file name"
    ColumnHeader f, 389, "Used by"
    ColumnHeader f, 469, "Status"
    Set lstJobs = NewCtl(f, "Forms.ListBox.1", 6, 15, 584, 70)
    lstJobs.ColumnCount = 4
    lstJobs.ColumnWidths = "190 pt;190 pt;80 pt;110 pt"
    lstJobs.BackColor = RGB(255, 255, 255)

    NewCtl f, "Forms.Label.1", 6, 94, 84, 14, "New file name:"
    Set txtNewName = NewCtl(f, "Forms.TextBox.1", 90, 91, 250, 18)
    Set lblExt = NewCtl(f, "Forms.Label.1", 344, 94, 70, 14)
    lblExt.ForeColor = RGB(110, 110, 110)
    Set lblNameHint = NewCtl(f, "Forms.Label.1", 414, 89, 176, 24)
    lblNameHint.Font.Size = 8
    lblNameHint.WordWrap = True

    Set btnAddSel = NewCtl(f, "Forms.CommandButton.1", 6, 115, 160, 22, "Add selected components")
    Set btnRemove = NewCtl(f, "Forms.CommandButton.1", 172, 115, 76, 22, "Remove")
    Set btnClear = NewCtl(f, "Forms.CommandButton.1", 254, 115, 76, 22, "Clear list")
    Set c = NewCtl(f, "Forms.Label.1", 338, 115, 252, 24, _
                   "Select parts or sub-assemblies in the FeatureManager tree or graphics area, then click Add.")
    c.ForeColor = RGB(110, 110, 110): c.Font.Size = 8: c.WordWrap = True

    '--- 2. search folder ------------------------------------------------------------
    Set f = NewFrame("2.  Where-used search folder", 8, 212, 600, 66)
    Set txtFolder = NewCtl(f, "Forms.TextBox.1", 6, 5, 394, 18)
    Set btnBrowse = NewCtl(f, "Forms.CommandButton.1", 406, 4, 86, 20, "Browse...")
    Set btnAsmFolder = NewCtl(f, "Forms.CommandButton.1", 498, 4, 92, 20, "Assembly folder")
    Set chkSub = NewCtl(f, "Forms.CheckBox.1", 6, 28, 120, 18, "Include sub-folders")
    Set lblScan = NewCtl(f, "Forms.Label.1", 132, 31, 268, 14, "Not searched yet.")
    lblScan.ForeColor = RGB(110, 110, 110)
    Set btnScan = NewCtl(f, "Forms.CommandButton.1", 406, 27, 184, 22, "Find where-used")

    '--- 3. options -------------------------------------------------------------------
    Set f = NewFrame("3.  Options", 8, 284, 600, 70)
    Set chkDrawing = NewCtl(f, "Forms.CheckBox.1", 6, 3, 292, 17, "Also rename the drawing that has the same name")
    Set chkClose = NewCtl(f, "Forms.CheckBox.1", 6, 20, 292, 17, "Close documents the macro had to open")
    Set chkFallback = NewCtl(f, "Forms.CheckBox.1", 6, 37, 292, 17, "Use Save As if Rename Document is not available")
    Set optKeep = NewCtl(f, "Forms.OptionButton.1", 304, 3, 290, 17, "Keep the old file where it is (safest)")
    Set optBackup = NewCtl(f, "Forms.OptionButton.1", 304, 20, 290, 17, _
                           "Move the old file to """ & BACKUP_FOLDER_NAME & """ after verification")
    Set chkLog = NewCtl(f, "Forms.CheckBox.1", 304, 37, 290, 17, "Write a log file next to the assembly")

    '--- progress + log ------------------------------------------------------------------
    Set lblStatus = NewCtl(Me, "Forms.Label.1", 8, 360, 600, 14, "Ready.")
    Set pbBack = NewCtl(Me, "Forms.Label.1", 8, 376, 600, 8)
    pbBack.BackColor = RGB(222, 222, 222)
    Set pbFill = NewCtl(Me, "Forms.Label.1", 8, 376, 1, 8)
    pbFill.BackColor = RGB(0, 120, 215)
    pbFill.Visible = False

    Set txtLog = NewCtl(Me, "Forms.TextBox.1", 8, 390, 600, 84)
    txtLog.MultiLine = True
    txtLog.ScrollBars = 2               ' vertical
    txtLog.Locked = True
    txtLog.Font.Name = "Consolas"
    txtLog.Font.Size = 8
    txtLog.BackColor = RGB(255, 255, 255)

    '--- buttons -------------------------------------------------------------------------
    Set btnOpenLog = NewCtl(Me, "Forms.CommandButton.1", 8, 482, 100, 26, "Open log file")
    Set btnRename = NewCtl(Me, "Forms.CommandButton.1", 374, 482, 136, 26, "Rename and verify")
    btnRename.BackColor = RGB(0, 120, 215)
    btnRename.ForeColor = RGB(255, 255, 255)
    btnRename.Font.Bold = True
    Set btnClose = NewCtl(Me, "Forms.CommandButton.1", 516, 482, 92, 26, "Close")
    btnClose.Cancel = True
End Sub

Private Function NewCtl(container As Object, ByVal progId As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single, Optional ByVal caption As String = "") As Object
    Dim c As Object
    Set c = container.Controls.Add(progId)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    If caption <> "" Then c.Caption = caption
    On Error GoTo 0
    Set NewCtl = c
End Function

Private Function NewFrame(ByVal caption As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single) As Object
    Dim f As Object
    Set f = NewCtl(Me, "Forms.Frame.1", l, t, w, h, caption)
    f.Font.Bold = True
    f.BackColor = RGB(243, 243, 243)
    Set NewFrame = f
End Function

Private Sub ColumnHeader(container As Object, ByVal l As Single, ByVal caption As String)
    Dim c As Object
    Set c = NewCtl(container, "Forms.Label.1", l, 2, 180, 12, caption)
    c.Font.Size = 8
    c.Font.Bold = True
    c.ForeColor = RGB(90, 90, 90)
End Sub

'===============================================================================
' Called by the SafeRename module
'===============================================================================
Public Sub ShowProgress(ByVal text As String, Optional ByVal fraction As Double = -1)
    lblStatus.Caption = text
    lblStatus.ForeColor = RGB(30, 30, 30)
    If fraction >= 0 Then
        If fraction > 1 Then fraction = 1
        pbFill.Width = IIf(pbBack.Width * fraction < 1, 1, pbBack.Width * fraction)
        pbFill.Visible = True
    End If
    Me.Repaint
End Sub

Public Sub AppendLog(ByVal text As String)
    txtLog.Text = txtLog.Text & text & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

'===============================================================================
' 1. Files to rename
'===============================================================================
Private Sub btnAddSel_Click()
    AddSelection True
End Sub

Private Sub AddSelection(ByVal showMessages As Boolean)
    Dim problem As String, m As SldWorks.ModelDoc2, sel As Object, k As Variant, added As Long, firstNew As Long

    Set m = GetTopAssembly(problem)
    If m Is Nothing Then
        ShowWarning problem
        If showMessages Then MsgBox problem, vbExclamation, TITLE
        Exit Sub
    End If

    If Not topModel Is Nothing Then
        If StrComp(topModel.GetPathName, m.GetPathName, vbTextCompare) <> 0 Then ClearJobs
    End If
    Set topModel = m
    If txtFolder.Text = "" Then txtFolder.Text = fso.GetParentFolderName(m.GetPathName)

    Set sel = CollectSelectedComponents(m)
    For Each k In sel.keys
        If FindJob(CStr(k)) = 0 Then
            jCount = jCount + 1
            ReDim Preserve jPath(1 To jCount)
            ReDim Preserve jNew(1 To jCount)
            ReDim Preserve jComp(1 To jCount)
            ReDim Preserve jDeps(1 To jCount)
            ReDim Preserve jStatus(1 To jCount)
            jPath(jCount) = CStr(k)
            jNew(jCount) = fso.GetBaseName(k)
            Set jComp(jCount) = sel(k)
            Set jDeps(jCount) = Nothing
            jStatus(jCount) = ""
            added = added + 1
            If firstNew = 0 Then firstNew = jCount
        End If
    Next k

    If added > 0 Then scanDone = False
    RefreshList
    If firstNew > 0 Then lstJobs.ListIndex = firstNew - 1

    If added > 0 Then
        ShowProgress added & " file(s) added. Type the new name for each file.", -1
    ElseIf GetSkippedText() <> "" Then
        ShowWarning "Some selected components cannot be renamed - see the log."
        If showMessages Then MsgBox "These components cannot be renamed:" & vbCrLf & GetSkippedText(), vbExclamation, TITLE
    ElseIf showMessages Then
        MsgBox "Nothing new is selected." & vbCrLf & vbCrLf & _
               "Select one or more parts or sub-assemblies in the assembly (FeatureManager tree or graphics area), " & _
               "then click 'Add selected components'.", vbInformation, TITLE
    End If
End Sub

Private Function FindJob(ByVal path As String) As Long
    Dim i As Long
    For i = 1 To jCount
        If StrComp(jPath(i), path, vbTextCompare) = 0 Then FindJob = i: Exit Function
    Next i
End Function

Private Sub btnRemove_Click()
    Dim idx As Long, i As Long
    idx = lstJobs.ListIndex + 1
    If idx < 1 Then Exit Sub
    For i = idx To jCount - 1
        jPath(i) = jPath(i + 1)
        jNew(i) = jNew(i + 1)
        Set jComp(i) = jComp(i + 1)
        Set jDeps(i) = jDeps(i + 1)
        jStatus(i) = jStatus(i + 1)
    Next i
    jCount = jCount - 1
    If jCount = 0 Then
        ClearJobs
    Else
        ReDim Preserve jPath(1 To jCount)
        ReDim Preserve jNew(1 To jCount)
        ReDim Preserve jComp(1 To jCount)
        ReDim Preserve jDeps(1 To jCount)
        ReDim Preserve jStatus(1 To jCount)
    End If
    RefreshList
    If jCount > 0 Then lstJobs.ListIndex = IIf(idx > jCount, jCount, idx) - 1
    LoadSelectedJob
End Sub

Private Sub btnClear_Click()
    ClearJobs
    RefreshList
    LoadSelectedJob
End Sub

Private Sub ClearJobs()
    jCount = 0
    Erase jPath, jNew, jComp, jDeps, jStatus
    scanDone = False
End Sub

Private Sub RefreshList()
    Dim i As Long, keep As Long
    suppress = True
    keep = lstJobs.ListIndex
    lstJobs.Clear
    For i = 1 To jCount
        lstJobs.AddItem fso.GetFileName(jPath(i))
        lstJobs.List(i - 1, 1) = NormName(i) & "." & fso.GetExtensionName(jPath(i))
        If jDeps(i) Is Nothing Then
            lstJobs.List(i - 1, 2) = "-"
        Else
            lstJobs.List(i - 1, 2) = jDeps(i).Count & " file(s)"
        End If
        lstJobs.List(i - 1, 3) = jStatus(i)
    Next i
    If keep >= 0 And keep < jCount Then lstJobs.ListIndex = keep
    suppress = False
End Sub

Private Sub lstJobs_Click()
    If Not suppress Then LoadSelectedJob
End Sub

Private Sub LoadSelectedJob()
    Dim idx As Long
    idx = lstJobs.ListIndex + 1
    suppress = True
    If idx < 1 Then
        txtNewName.Text = ""
        txtNewName.Enabled = False
        lblExt.Caption = ""
        lblNameHint.Caption = ""
    Else
        txtNewName.Enabled = (jStatus(idx) <> "Renamed")
        txtNewName.Text = jNew(idx)
        lblExt.Caption = "." & fso.GetExtensionName(jPath(idx))
        ShowProgress jPath(idx), -1
    End If
    suppress = False
    If idx >= 1 Then ValidateRow idx
End Sub

Private Sub txtNewName_Change()
    Dim idx As Long
    If suppress Then Exit Sub
    idx = lstJobs.ListIndex + 1
    If idx < 1 Then Exit Sub
    jNew(idx) = txtNewName.Text
    lstJobs.List(idx - 1, 1) = NormName(idx) & "." & fso.GetExtensionName(jPath(idx))
    ValidateRow idx
End Sub

' Shows whether the name of row idx is valid. Returns the problem ("" = OK).
Private Function ValidateRow(ByVal idx As Long) As String
    Dim nm As String, problem As String
    If jStatus(idx) = "Renamed" Then
        lblNameHint.Caption = "Renamed."
        lblNameHint.ForeColor = RGB(16, 124, 16)
        Exit Function
    End If
    nm = jNew(idx)
    problem = ValidateNewName(jPath(idx), nm, UsedPaths(idx))
    If problem = "" Then
        lblNameHint.Caption = "Name OK"
        lblNameHint.ForeColor = RGB(16, 124, 16)
    Else
        lblNameHint.Caption = problem
        lblNameHint.ForeColor = RGB(196, 43, 28)
    End If
    ValidateRow = problem
End Function

' New paths of all pending rows except skipIdx.
Private Function UsedPaths(ByVal skipIdx As Long) As Object
    Dim d As Object, i As Long, p As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For i = 1 To jCount
        If i <> skipIdx And jStatus(i) <> "Renamed" Then
            p = NewPathFor(jPath(i), NormName(i))
            If Not d.Exists(p) Then d.Add p, True
        End If
    Next i
    Set UsedPaths = d
End Function

' New name of row i, trimmed and without a typed extension.
Private Function NormName(ByVal i As Long) As String
    Dim nm As String, ext As String
    nm = Trim$(jNew(i))
    ext = "." & LCase$(fso.GetExtensionName(jPath(i)))
    If LCase$(Right$(nm, Len(ext))) = ext Then nm = Left$(nm, Len(nm) - Len(ext))
    NormName = nm
End Function

'===============================================================================
' 2. Where-used search
'===============================================================================
Private Sub btnBrowse_Click()
    Dim sh As Object, f As Object
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0, "Folder to search for assemblies and drawings that use the file(s):", &H1 Or &H40, 0)
    If Not f Is Nothing Then txtFolder.Text = f.Self.path
End Sub

Private Sub btnAsmFolder_Click()
    If Not topModel Is Nothing Then txtFolder.Text = fso.GetParentFolderName(topModel.GetPathName)
End Sub

Private Sub txtFolder_Change()
    InvalidateScan
End Sub

Private Sub chkSub_Click()
    InvalidateScan
End Sub

Private Sub InvalidateScan()
    Dim i As Long
    If Not scanDone Then Exit Sub
    scanDone = False
    For i = 1 To jCount
        Set jDeps(i) = Nothing
    Next i
    lblScan.Caption = "Search folder changed - search again."
    RefreshList
End Sub

Private Sub btnScan_Click()
    DoScan
End Sub

Private Function DoScan() As Boolean
    Dim files As Collection, i As Long
    If jCount = 0 Then
        MsgBox "Add at least one component first.", vbInformation, TITLE
        Exit Function
    End If
    If Not fso.FolderExists(txtFolder.Text) Then
        MsgBox "Choose an existing search folder.", vbExclamation, TITLE
        Exit Function
    End If

    SetBusy True
    ClearDependencyCache
    ShowProgress "Listing SOLIDWORKS files in " & txtFolder.Text & " ...", 0
    Set files = ListSwFiles(txtFolder.Text, chkSub.Value)
    LogMsg "Search folder: " & txtFolder.Text & IIf(chkSub.Value, " (with sub-folders)", "") & _
           " - " & files.Count & " SOLIDWORKS file(s)"

    For i = 1 To jCount
        If jStatus(i) <> "Renamed" Then Set jDeps(i) = FindDependents(jPath(i), files, i, jCount)
    Next i

    scanDone = True
    lblScan.Caption = files.Count & " SOLIDWORKS file(s) searched."
    RefreshList
    ShowProgress "Where-used search finished.", 1
    SetBusy False
    DoScan = True
End Function

'===============================================================================
' Rename
'===============================================================================
Private Sub btnRename_Click()
    Dim i As Long, nm As String, problem As String, used As Object, pending As Long, summary As String

    If busy Then Exit Sub
    For i = 1 To jCount
        If jStatus(i) <> "Renamed" Then pending = pending + 1
    Next i
    If pending = 0 Then
        MsgBox "Add the components to rename first: select them in the assembly, then click 'Add selected components'.", _
               vbInformation, TITLE
        Exit Sub
    End If

    ' validate every pending name
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    For i = 1 To jCount
        If jStatus(i) <> "Renamed" Then
            nm = jNew(i)
            problem = ValidateNewName(jPath(i), nm, used)
            If problem <> "" Then
                lstJobs.ListIndex = i - 1
                MsgBox fso.GetFileName(jPath(i)) & ":" & vbCrLf & problem, vbExclamation, TITLE
                Exit Sub
            End If
            jNew(i) = nm
            used.Add NewPathFor(jPath(i), nm), True
        End If
    Next i

    If Not scanDone Then
        If Not DoScan() Then Exit Sub
    End If

    ' confirm
    For i = 1 To jCount
        If jStatus(i) <> "Renamed" Then
            summary = summary & "   " & fso.GetFileName(jPath(i)) & "  ->  " & jNew(i) & "." & _
                      fso.GetExtensionName(jPath(i)) & "   (used by " & jDeps(i).Count & ")" & vbCrLf
        End If
    Next i
    If Len(summary) > 700 Then summary = Left$(summary, 700) & "   ..." & vbCrLf
    If MsgBox("Rename " & pending & " file(s)?" & vbCrLf & vbCrLf & summary & vbCrLf & _
              "All referencing files are opened and SAVED, including the top assembly." & vbCrLf & _
              "Make sure nobody else has these files open and that you have a backup.", _
              vbQuestion + vbOKCancel, TITLE) <> vbOK Then Exit Sub

    ' options
    Opt_RenameDrawing = chkDrawing.Value
    Opt_CloseOpened = chkClose.Value
    Opt_SaveAsFallback = chkFallback.Value
    Opt_WriteLog = chkLog.Value
    Opt_BackupOldFile = optBackup.Value

    SetBusy True
    LogMsg ""
    LogMsg "Rename started " & Now & " - top assembly: " & topModel.GetPathName

    Dim okCount As Long, failCount As Long, oldPath As String
    For i = 1 To jCount
        If jStatus(i) <> "Renamed" Then
            jStatus(i) = "Working..."
            RefreshList
            oldPath = jPath(i)
            If ProcessJob(topModel, jComp(i), oldPath, jNew(i), jDeps(i)) Then
                okCount = okCount + 1
                jStatus(i) = "Renamed"
                jPath(i) = NewPathFor(oldPath, jNew(i))
            Else
                failCount = failCount + 1
                jStatus(i) = "FAILED - see log"
            End If
            Set jDeps(i) = Nothing
            RefreshList
        End If
    Next i

    scanDone = False
    lblScan.Caption = "Not searched yet."
    lastLogPath = SaveLog(topModel)
    SetBusy False
    LoadSelectedJob

    If failCount = 0 Then
        ShowProgress okCount & " file(s) renamed and verified.", 1
        MsgBox okCount & " file(s) renamed. All references found were updated and verified." & _
               IIf(lastLogPath <> "", vbCrLf & vbCrLf & "Log: " & lastLogPath, ""), vbInformation, TITLE
    Else
        ShowWarning okCount & " renamed, " & failCount & " failed - see the log."
        MsgBox okCount & " file(s) renamed, " & failCount & " FAILED." & vbCrLf & _
               "Failed files were left untouched where possible - check the log." & _
               IIf(lastLogPath <> "", vbCrLf & vbCrLf & "Log: " & lastLogPath, ""), vbExclamation, TITLE
    End If
End Sub

'===============================================================================
' Misc
'===============================================================================
Private Sub btnOpenLog_Click()
    If lastLogPath = "" Or Not fso.FileExists(lastLogPath) Then
        MsgBox "No log file has been written yet. It is created after a rename.", vbInformation, TITLE
    Else
        Shell "notepad.exe """ & lastLogPath & """", vbNormalFocus
    End If
End Sub

Private Sub btnClose_Click()
    If Not busy Then Unload Me
End Sub

Private Sub SetBusy(ByVal b As Boolean)
    busy = b
    btnAddSel.Enabled = Not b
    btnRemove.Enabled = Not b
    btnClear.Enabled = Not b
    btnBrowse.Enabled = Not b
    btnAsmFolder.Enabled = Not b
    btnScan.Enabled = Not b
    btnRename.Enabled = Not b
    btnClose.Enabled = Not b
    btnOpenLog.Enabled = Not b
    lstJobs.Enabled = Not b
    txtFolder.Enabled = Not b
    Me.MousePointer = IIf(b, 11, 0)     ' fmMousePointerHourGlass / default
    Me.Repaint
End Sub

Private Sub ShowWarning(ByVal text As String)
    lblStatus.Caption = text
    lblStatus.ForeColor = RGB(196, 43, 28)
End Sub
