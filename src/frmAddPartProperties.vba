'===============================================================================
' frmAddPartProperties - wizard of the Add Part Properties macro.
'
' All controls are created in code (UserForm_Initialize), so this code can also
' be pasted into an EMPTY UserForm named "frmAddPartProperties" if you build the
' macro by hand. The work itself is done by the AddPartProperties module.
'
' Pages: Start -> Material -> Exceptions -> Size -> one page per part -> Review.
' Pages for properties that are not ticked on the Start page are left out.
'===============================================================================
Option Explicit

Private Const TITLE As String = "Add Part Properties"

Private Const PG_START As Long = 1
Private Const PG_MAT As Long = 2
Private Const PG_EXC As Long = 3
Private Const PG_SIZE As Long = 4
Private Const PG_PART As Long = 5
Private Const PG_REVIEW As Long = 6
Private Const PAGE_COUNT As Long = 6

Private Const RED As Long = &H1C2BC4         ' RGB(196, 43, 28)
Private Const GREY As Long = &H6E6E6E        ' RGB(110, 110, 110)

' shared buttons
Private WithEvents btnBack As MSForms.CommandButton
Private WithEvents btnNext As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

' start page
Private WithEvents btnShowSkipped As MSForms.CommandButton
Private WithEvents chkDesc As MSForms.CheckBox
Private WithEvents chkMat As MSForms.CheckBox
Private WithEvents chkWeight As MSForms.CheckBox
Private WithEvents chkSize As MSForms.CheckBox
Private WithEvents optKeep As MSForms.OptionButton
Private WithEvents optDelete As MSForms.OptionButton
Private lblDelWarn As MSForms.Label

' material + exceptions pages
Private txtMat As MSForms.TextBox
Private optMatAll As MSForms.OptionButton
Private optMatExc As MSForms.OptionButton
Private WithEvents lstExc As MSForms.ListBox
Private txtExcMat As MSForms.TextBox
Private WithEvents btnExcSet As MSForms.CommandButton
Private WithEvents btnExcReset As MSForms.CommandButton
Private WithEvents btnExcNone As MSForms.CommandButton

' size page
Private WithEvents optSizeAll As MSForms.OptionButton
Private WithEvents optSizePer As MSForms.OptionButton
Private txtSizeAll As MSForms.TextBox

' per-part page
Private lblPartName As MSForms.Label
Private lblPartInfo As MSForms.Label
Private lblDesc As MSForms.Label
Private txtDesc As MSForms.TextBox
Private lblSize As MSForms.Label
Private txtSize As MSForms.TextBox
Private WithEvents btnSkipPart As MSForms.CommandButton
Private WithEvents btnSkipRest As MSForms.CommandButton

' review page
Private lblReviewWarn As MSForms.Label
Private lstReview As MSForms.ListBox
Private WithEvents chkOverwrite As MSForms.CheckBox
Private chkSave As MSForms.CheckBox

' progress
Private lblStatus As MSForms.Label
Private pbBack As MSForms.Label
Private pbFill As MSForms.Label

Private pages(1 To PAGE_COUNT) As Object     ' one frame per page
Private pageTitle(1 To PAGE_COUNT) As MSForms.Label
Private history As Collection                ' pages visited, for Back
Private curPage As Long
Private curPart As Long
Private busy As Boolean
Private finished As Boolean
Private lastLogPath As String
Private fso As Object

'===============================================================================
' Form life cycle
'===============================================================================
Private Sub UserForm_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set history = New Collection
    BuildUI
    Set gUI = Me

    chkDesc.Value = Opt_AddDesc
    chkMat.Value = Opt_AddMat
    chkWeight.Value = Opt_AddWeight
    chkSize.Value = Opt_AddSize
    optKeep.Value = Not Opt_DeleteOld
    optDelete.Value = Opt_DeleteOld
    optMatAll.Value = True
    optSizeAll.Value = True
    txtSizeAll.Text = SizeAll
    chkOverwrite.Value = Opt_Overwrite
    chkSave.Value = Opt_Save
    FillExceptions

    ShowPage PG_START
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
    Dim f As Object, c As Object, n As Long

    Me.Caption = "Add Part Properties - custom properties for every part in the assembly"
    Me.BackColor = RGB(243, 243, 243)
    Me.Width = 624
    Me.Height = 486

    '--- header ------------------------------------------------------------------
    Set c = NewCtl(Me, "Forms.Label.1", 0, 0, 624, 44)
    c.BackColor = RGB(31, 56, 100)
    Set c = NewCtl(Me, "Forms.Label.1", 12, 4, 400, 22, "Add Part Properties")
    c.BackStyle = 0: c.ForeColor = RGB(255, 255, 255): c.Font.Size = 14: c.Font.Bold = True
    Set c = NewCtl(Me, "Forms.Label.1", 12, 26, 600, 14, _
                   PartCount & " part(s) in " & fso.GetFileName(TopDoc.GetPathName) & _
                   "  -  DESCRIPTION, Material, Weight (kg) and Size on the Custom tab")
    c.BackStyle = 0: c.ForeColor = RGB(200, 215, 235)

    For n = 1 To PAGE_COUNT
        Set pages(n) = NewCtl(Me, "Forms.Frame.1", 8, 50, 600, 330)
        pages(n).BackColor = RGB(243, 243, 243)
        pages(n).SpecialEffect = 0                  ' flat
        pages(n).BorderStyle = 0
        Set pageTitle(n) = NewCtl(pages(n), "Forms.Label.1", 4, 2, 590, 18)
        pageTitle(n).Font.Size = 11
        pageTitle(n).Font.Bold = True
        pageTitle(n).ForeColor = RGB(31, 56, 100)
    Next n

    '--- start -------------------------------------------------------------------
    Set f = pages(PG_START)
    pageTitle(PG_START).Caption = "1.  What to do"
    Set c = NewCtl(f, "Forms.Label.1", 4, 24, 440, 28, _
                   PartCount & " part file(s) will be processed (each file once, sub-assemblies included). " & _
                   GetSkippedCount() & " skipped.")
    c.WordWrap = True
    Set btnShowSkipped = NewCtl(f, "Forms.CommandButton.1", 470, 24, 120, 22, "Show skipped...")
    btnShowSkipped.Enabled = (GetSkippedCount() > 0)

    Set c = NewCtl(f, "Forms.Frame.1", 4, 58, 586, 118, "Properties to add")
    c.BackColor = RGB(243, 243, 243)
    Set chkDesc = NewCtl(c, "Forms.CheckBox.1", 8, 6, 130, 18, PROP_DESC)
    Hint c, 140, 8, "Text, typed for each part."
    Set chkMat = NewCtl(c, "Forms.CheckBox.1", 8, 30, 130, 18, PROP_MAT)
    Hint c, 140, 32, "Text, typed once for all parts - with exceptions if needed."
    Set chkWeight = NewCtl(c, "Forms.CheckBox.1", 8, 54, 130, 18, PROP_WEIGHT)
    Hint c, 140, 56, "Equation linked to SW-Mass, in kg. Added automatically to every part."
    Set chkSize = NewCtl(c, "Forms.CheckBox.1", 8, 78, 130, 18, PROP_SIZE)
    Hint c, 140, 80, "Text, typed - one value for all parts (default ""-"") or per part."

    Set c = NewCtl(f, "Forms.Frame.1", 4, 184, 586, 104, "Old custom properties (Custom tab)")
    c.BackColor = RGB(243, 243, 243)
    Set optKeep = NewCtl(c, "Forms.OptionButton.1", 8, 6, 560, 18, "Keep existing properties")
    Set optDelete = NewCtl(c, "Forms.OptionButton.1", 8, 28, 560, 18, _
                           "Delete all existing custom properties first, then add the ticked ones")
    Set lblDelWarn = NewCtl(c, "Forms.Label.1", 26, 50, 550, 30)
    lblDelWarn.ForeColor = RED
    lblDelWarn.WordWrap = True
    Hint f, 8, 296, "Configuration-specific properties are never changed."

    '--- material ----------------------------------------------------------------
    Set f = pages(PG_MAT)
    pageTitle(PG_MAT).Caption = "2.  Material"
    NewCtl f, "Forms.Label.1", 4, 30, 140, 16, "Material for all parts:"
    Set txtMat = NewCtl(f, "Forms.TextBox.1", 146, 27, 300, 20)
    Set optMatAll = NewCtl(f, "Forms.OptionButton.1", 4, 62, 500, 18, "Apply to all parts")
    Set optMatExc = NewCtl(f, "Forms.OptionButton.1", 4, 84, 500, 18, _
                           "Apply to all parts, with exceptions (pick parts that get a different material)")
    Hint f, 4, 112, "The text is written as typed. It does not change the SOLIDWORKS material used for the mass."

    '--- exceptions --------------------------------------------------------------
    Set f = pages(PG_EXC)
    pageTitle(PG_EXC).Caption = "2b.  Material exceptions"
    Hint f, 4, 22, "Tick parts, type their material and click Set. Unchanged parts get the material for all parts."
    ColumnHeader f, 22, 38, "Part"
    ColumnHeader f, 222, 38, "In sub-assembly"
    ColumnHeader f, 402, 38, "Material"
    Set lstExc = NewCtl(f, "Forms.ListBox.1", 4, 52, 586, 210)
    lstExc.ColumnCount = 3
    lstExc.ColumnWidths = "200 pt;180 pt;180 pt"
    lstExc.ListStyle = 1                ' fmListStyleOption: tick boxes
    lstExc.MultiSelect = 1              ' fmMultiSelectMulti
    lstExc.BackColor = RGB(255, 255, 255)
    NewCtl f, "Forms.Label.1", 4, 274, 160, 16, "Material for ticked parts:"
    Set txtExcMat = NewCtl(f, "Forms.TextBox.1", 150, 271, 180, 20)
    Set btnExcSet = NewCtl(f, "Forms.CommandButton.1", 336, 270, 60, 22, "Set")
    Set btnExcReset = NewCtl(f, "Forms.CommandButton.1", 402, 270, 110, 22, "Reset to common")
    Set btnExcNone = NewCtl(f, "Forms.CommandButton.1", 518, 270, 72, 22, "Untick all")

    '--- size --------------------------------------------------------------------
    Set f = pages(PG_SIZE)
    pageTitle(PG_SIZE).Caption = "3.  Size"
    Set optSizeAll = NewCtl(f, "Forms.OptionButton.1", 4, 30, 200, 18, "Same value for all parts:")
    Set txtSizeAll = NewCtl(f, "Forms.TextBox.1", 206, 29, 240, 20)
    Set optSizePer = NewCtl(f, "Forms.OptionButton.1", 4, 58, 500, 18, "Type per part (asked on the page of each part)")

    '--- per part ----------------------------------------------------------------
    Set f = pages(PG_PART)
    Set lblPartName = NewCtl(f, "Forms.Label.1", 4, 26, 586, 18)
    lblPartName.Font.Size = 10: lblPartName.Font.Bold = True
    Set lblPartInfo = NewCtl(f, "Forms.Label.1", 4, 46, 586, 42)
    lblPartInfo.ForeColor = GREY: lblPartInfo.WordWrap = True
    Set lblDesc = NewCtl(f, "Forms.Label.1", 4, 100, 90, 16, PROP_DESC & ":")
    Set txtDesc = NewCtl(f, "Forms.TextBox.1", 96, 97, 494, 20)
    Set lblSize = NewCtl(f, "Forms.Label.1", 4, 128, 90, 16, PROP_SIZE & ":")
    Set txtSize = NewCtl(f, "Forms.TextBox.1", 96, 125, 240, 20)
    Hint f, 4, 158, "Next (or Enter) stores the values. An empty DESCRIPTION keeps the existing one. " & _
                    "The part is highlighted in the graphics area."
    Set btnSkipPart = NewCtl(f, "Forms.CommandButton.1", 4, 186, 120, 22, "Skip this part")
    Set btnSkipRest = NewCtl(f, "Forms.CommandButton.1", 130, 186, 140, 22, "Skip remaining parts")

    '--- review ------------------------------------------------------------------
    Set f = pages(PG_REVIEW)
    pageTitle(PG_REVIEW).Caption = "Review"
    Set lblReviewWarn = NewCtl(f, "Forms.Label.1", 4, 22, 586, 16)
    lblReviewWarn.ForeColor = RED
    Set lstReview = NewCtl(f, "Forms.ListBox.1", 4, 56, 586, 230)
    lstReview.ColumnCount = 6
    lstReview.BackColor = RGB(255, 255, 255)
    Set chkOverwrite = NewCtl(f, "Forms.CheckBox.1", 4, 294, 280, 18, "Overwrite existing values")
    Set chkSave = NewCtl(f, "Forms.CheckBox.1", 290, 294, 280, 18, "Save parts after writing")

    '--- progress + buttons ------------------------------------------------------
    Set lblStatus = NewCtl(Me, "Forms.Label.1", 8, 386, 600, 14, "")
    Set pbBack = NewCtl(Me, "Forms.Label.1", 8, 402, 600, 6)
    pbBack.BackColor = RGB(222, 222, 222)
    Set pbFill = NewCtl(Me, "Forms.Label.1", 8, 402, 1, 6)
    pbFill.BackColor = RGB(0, 120, 215)
    pbFill.Visible = False

    Set btnBack = NewCtl(Me, "Forms.CommandButton.1", 316, 420, 92, 26, "< Back")
    Set btnNext = NewCtl(Me, "Forms.CommandButton.1", 414, 420, 96, 26, "Next >")
    btnNext.BackColor = RGB(0, 120, 215)
    btnNext.ForeColor = RGB(255, 255, 255)
    btnNext.Font.Bold = True
    btnNext.Default = True
    Set btnCancel = NewCtl(Me, "Forms.CommandButton.1", 516, 420, 92, 26, "Cancel")
    btnCancel.Cancel = True
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

Private Sub Hint(container As Object, ByVal l As Single, ByVal t As Single, ByVal caption As String)
    Dim c As Object
    Set c = NewCtl(container, "Forms.Label.1", l, t, 590 - l, 14, caption)
    c.ForeColor = GREY
    c.Font.Size = 8
End Sub

Private Sub ColumnHeader(container As Object, ByVal l As Single, ByVal t As Single, ByVal caption As String)
    Dim c As Object
    Set c = NewCtl(container, "Forms.Label.1", l, t, 180, 12, caption)
    c.Font.Size = 8
    c.Font.Bold = True
    c.ForeColor = RGB(90, 90, 90)
End Sub

'===============================================================================
' Called by the AddPartProperties module
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

'===============================================================================
' Navigation
'===============================================================================
Private Sub ShowPage(ByVal pg As Long)
    Dim n As Long
    curPage = pg
    For n = 1 To PAGE_COUNT
        pages(n).Visible = (n = pg)
    Next n
    lblStatus.Caption = ""
    Select Case pg
        Case PG_START: UpdateStart
        Case PG_MAT: txtMat.SetFocus
        Case PG_SIZE: UpdateSize
        Case PG_PART: LoadPart
        Case PG_REVIEW: FillReview
    End Select
    btnBack.Enabled = (history.Count > 0 Or (pg = PG_PART And curPart > 1)) And Not finished
    btnNext.Caption = IIf(pg = PG_REVIEW, "Apply", "Next >")
End Sub

Private Sub GoPage(ByVal pg As Long)
    history.Add curPage
    ShowPage pg
End Sub

' Page that follows pg, given the choices made so far.
Private Function PageAfter(ByVal pg As Long) As Long
    Select Case pg
        Case PG_START
            If Opt_AddMat Then PageAfter = PG_MAT Else PageAfter = PageAfter(PG_EXC)
        Case PG_MAT
            If optMatExc.Value Then PageAfter = PG_EXC Else PageAfter = PageAfter(PG_EXC)
        Case PG_EXC
            If Opt_AddSize Then PageAfter = PG_SIZE Else PageAfter = PageAfter(PG_SIZE)
        Case PG_SIZE
            If NeedPartPages() Then PageAfter = PG_PART Else PageAfter = PG_REVIEW
    End Select
End Function

Private Function NeedPartPages() As Boolean
    NeedPartPages = Opt_AddDesc Or (Opt_AddSize And Not SizeSameForAll)
End Function

Private Sub btnNext_Click()
    Dim i As Long
    If busy Then Exit Sub
    If finished Then
        OpenLog
        Exit Sub
    End If
    Select Case curPage
        Case PG_START
            Opt_AddDesc = chkDesc.Value
            Opt_AddMat = chkMat.Value
            Opt_AddWeight = chkWeight.Value
            Opt_AddSize = chkSize.Value
            Opt_DeleteOld = optDelete.Value
            ' choices that only exist on skipped pages fall back to their defaults
            If Not Opt_AddSize Then SizeSameForAll = True
            GoPage PageAfter(PG_START)
        Case PG_MAT
            If Trim$(txtMat.Text) = "" Then
                MsgBox "Type the material for all parts.", vbExclamation, TITLE
                txtMat.SetFocus
                Exit Sub
            End If
            MatCommon = Trim$(txtMat.Text)
            If optMatAll.Value Then
                For i = 1 To PartCount
                    pMatExc(i) = False
                Next i
            End If
            FillExceptions
            GoPage PageAfter(PG_MAT)
        Case PG_EXC
            GoPage PageAfter(PG_EXC)
        Case PG_SIZE
            SizeSameForAll = optSizeAll.Value
            SizeAll = Trim$(txtSizeAll.Text)
            If SizeAll = "" Then SizeAll = "-"
            GoPage PageAfter(PG_SIZE)
        Case PG_PART
            StorePart True
            NextPart
        Case PG_REVIEW
            Apply
    End Select
End Sub

Private Sub btnBack_Click()
    If busy Or finished Then Exit Sub
    If curPage = PG_PART And curPart > 1 Then
        StorePart True
        curPart = curPart - 1
        ShowPage PG_PART
        Exit Sub
    End If
    If curPage = PG_PART Then EndPartPages
    If history.Count = 0 Then Exit Sub
    Dim pg As Long
    pg = history(history.Count)
    history.Remove history.Count
    If pg = PG_PART Then curPart = PartCount
    ShowPage pg
End Sub

Private Sub btnCancel_Click()
    If busy Then Exit Sub
    If Not finished And curPage <> PG_START Then
        If MsgBox("Stop without changing any file?", vbQuestion + vbYesNo, TITLE) <> vbYes Then Exit Sub
    End If
    If curPage = PG_PART Then EndPartPages
    Unload Me
End Sub

'===============================================================================
' Start page
'===============================================================================
Private Sub chkDesc_Click()
    UpdateStart
End Sub
Private Sub chkMat_Click()
    UpdateStart
End Sub
Private Sub chkWeight_Click()
    UpdateStart
End Sub
Private Sub chkSize_Click()
    UpdateStart
End Sub
Private Sub optKeep_Click()
    UpdateStart
End Sub
Private Sub optDelete_Click()
    UpdateStart
End Sub

Private Sub UpdateStart()
    Dim anyTicked As Boolean, i As Long, n As Long
    anyTicked = chkDesc.Value Or chkMat.Value Or chkWeight.Value Or chkSize.Value
    If optDelete.Value Then
        For i = 1 To PartCount
            If pPropCount(i) > 0 Then n = n + 1
        Next i
        lblDelWarn.Caption = "All properties on the Custom tab of " & n & " part(s) will be deleted. " & _
                             "This cannot be undone (the deleted values are written to the log)."
    Else
        lblDelWarn.Caption = ""
    End If
    btnNext.Enabled = anyTicked Or optDelete.Value
End Sub

Private Sub btnShowSkipped_Click()
    Dim s As String
    s = GetSkippedText()
    If Len(s) > 1500 Then s = Left$(s, 1500) & "  ..." & vbCrLf & "(full list in the log)"
    MsgBox "These components are not changed:" & vbCrLf & vbCrLf & s, vbInformation, TITLE
End Sub

'===============================================================================
' Material exceptions
'===============================================================================
Private Sub FillExceptions()
    Dim i As Long
    lstExc.Clear
    For i = 1 To PartCount
        lstExc.AddItem fso.GetFileName(pPath(i))
        lstExc.List(i - 1, 1) = IIf(pWhere(i) = "", "(top level)", pWhere(i))
        lstExc.List(i - 1, 2) = IIf(pMatExc(i), pMat(i), "(all: " & MatCommon & ")")
    Next i
End Sub

Private Sub lstExc_Change()
    If lstExc.ListIndex >= 0 Then SelectPart TopDoc, lstExc.ListIndex + 1, False
End Sub

Private Sub btnExcSet_Click()
    Dim i As Long, v As String, n As Long
    v = Trim$(txtExcMat.Text)
    If v = "" Then
        MsgBox "Type the material for the ticked parts.", vbExclamation, TITLE
        Exit Sub
    End If
    For i = 1 To PartCount
        If lstExc.Selected(i - 1) Then
            pMatExc(i) = True
            pMat(i) = v
            n = n + 1
        End If
    Next i
    If n = 0 Then MsgBox "Tick one or more parts first.", vbInformation, TITLE
    RefreshExceptions
End Sub

Private Sub btnExcReset_Click()
    Dim i As Long
    For i = 1 To PartCount
        If lstExc.Selected(i - 1) Then pMatExc(i) = False
    Next i
    RefreshExceptions
End Sub

Private Sub btnExcNone_Click()
    Dim i As Long
    For i = 1 To PartCount
        lstExc.Selected(i - 1) = False
    Next i
End Sub

Private Sub RefreshExceptions()
    Dim i As Long, n As Long
    For i = 1 To PartCount
        lstExc.List(i - 1, 2) = IIf(pMatExc(i), pMat(i), "(all: " & MatCommon & ")")
        If pMatExc(i) Then n = n + 1
    Next i
    lblStatus.Caption = n & " exception(s)."
End Sub

'===============================================================================
' Size
'===============================================================================
Private Sub optSizeAll_Click()
    UpdateSize
End Sub
Private Sub optSizePer_Click()
    UpdateSize
End Sub

Private Sub UpdateSize()
    txtSizeAll.Enabled = optSizeAll.Value
    txtSizeAll.BackColor = IIf(optSizeAll.Value, RGB(255, 255, 255), RGB(235, 235, 235))
End Sub

'===============================================================================
' Per-part pages
'===============================================================================
Private Sub LoadPart()
    Dim i As Long
    If curPart < 1 Then curPart = 1
    i = curPart
    pageTitle(PG_PART).Caption = "Part " & i & " of " & PartCount
    lblPartName.Caption = fso.GetFileName(pPath(i))
    lblPartInfo.Caption = pPath(i) & vbCrLf & "Configuration: " & pConfig(i) & _
                          IIf(pWhere(i) = "", "", "    In: " & pWhere(i))

    lblDesc.Visible = Opt_AddDesc
    txtDesc.Visible = Opt_AddDesc
    txtDesc.Text = pDesc(i)
    lblSize.Visible = Opt_AddSize And Not SizeSameForAll
    txtSize.Visible = lblSize.Visible
    txtSize.Text = pSize(i)

    SelectPart TopDoc, i, True
    If txtDesc.Visible Then
        txtDesc.SetFocus
        txtDesc.SelStart = 0
        txtDesc.SelLength = Len(txtDesc.Text)
    ElseIf txtSize.Visible Then
        txtSize.SetFocus
    End If
End Sub

' Stores the text boxes of the current part. keep = False: Skip (keep existing).
Private Sub StorePart(ByVal keep As Boolean)
    Dim i As Long
    i = curPart
    If i < 1 Or i > PartCount Then Exit Sub
    If keep Then
        If Opt_AddDesc Then
            pDesc(i) = Trim$(txtDesc.Text)
            pDescSet(i) = (pDesc(i) <> "")
            If Not pDescSet(i) Then pDesc(i) = pOldDesc(i)
        End If
        If Opt_AddSize And Not SizeSameForAll Then
            pSize(i) = Trim$(txtSize.Text)
            If pSize(i) = "" Then pSize(i) = "-"
            pSizeSet(i) = True
        End If
    Else
        pDescSet(i) = False
        pDesc(i) = pOldDesc(i)
        pSizeSet(i) = False
        pSize(i) = IIf(pHasSize(i), pOldSize(i), "-")
    End If
End Sub

Private Sub NextPart()
    If curPart < PartCount Then
        curPart = curPart + 1
        ShowPage PG_PART
    Else
        EndPartPages
        GoPage PG_REVIEW
    End If
End Sub

Private Sub btnSkipPart_Click()
    StorePart False
    NextPart
End Sub

Private Sub btnSkipRest_Click()
    Dim i As Long, first As Long
    first = curPart
    For i = first To PartCount
        curPart = i
        StorePart False
    Next i
    EndPartPages
    GoPage PG_REVIEW
End Sub

Private Sub EndPartPages()
    On Error Resume Next
    TopDoc.ClearSelection2 True
    TopDoc.ViewZoomtofit2
End Sub

'===============================================================================
' Review and apply
'===============================================================================
Private Sub chkOverwrite_Click()
    If curPage = PG_REVIEW And Not busy Then FillReview
End Sub

Private Sub FillReview()
    Dim i As Long, n As Long
    Opt_Overwrite = chkOverwrite.Value
    chkOverwrite.Visible = Not Opt_DeleteOld

    lstReview.ColumnWidths = "130 pt;" & IIf(Opt_AddDesc Or Opt_DeleteOld, "120 pt;", "0 pt;") & _
                             IIf(Opt_AddMat Or Opt_DeleteOld, "80 pt;", "0 pt;") & _
                             IIf(Opt_AddWeight Or Opt_DeleteOld, "80 pt;", "0 pt;") & _
                             IIf(Opt_AddSize Or Opt_DeleteOld, "60 pt;", "0 pt;") & "140 pt"
    lstReview.Clear
    lstReview.AddItem "Part"
    lstReview.List(0, 1) = PROP_DESC
    lstReview.List(0, 2) = PROP_MAT
    lstReview.List(0, 3) = PROP_WEIGHT
    lstReview.List(0, 4) = PROP_SIZE
    lstReview.List(0, 5) = IIf(finished, "Result", "Note")
    For i = 1 To PartCount
        lstReview.AddItem fso.GetFileName(pPath(i))
        lstReview.List(i, 1) = PlannedText(i, PROP_DESC)
        lstReview.List(i, 2) = PlannedText(i, PROP_MAT)
        lstReview.List(i, 3) = PlannedText(i, PROP_WEIGHT)
        lstReview.List(i, 4) = PlannedText(i, PROP_SIZE)
        lstReview.List(i, 5) = IIf(pResult(i) <> "", pResult(i), PlannedNote(i))
        If pPropCount(i) > 0 Then n = n + 1
    Next i

    If Opt_DeleteOld Then
        lblReviewWarn.Caption = "All old custom properties of " & n & " part(s) will be deleted before writing."
    Else
        lblReviewWarn.Caption = ""
    End If
    If Not finished Then ShowProgress PartCount & " part(s) will be changed. Click Apply to write.", -1
End Sub

Private Sub Apply()
    Dim i As Long, okCount As Long, failCount As Long, msg As String

    Opt_Overwrite = chkOverwrite.Value
    Opt_Save = chkSave.Value
    msg = "Write the properties to " & PartCount & " part(s)" & IIf(Opt_Save, " and save them", "") & "?"
    If Opt_DeleteOld Then msg = msg & vbCrLf & vbCrLf & _
        "All old properties on the Custom tab of these parts will be DELETED first. This cannot be undone."
    If MsgBox(msg, IIf(Opt_DeleteOld, vbExclamation, vbQuestion) + vbOKCancel, TITLE) <> vbOK Then Exit Sub

    SetBusy True
    LogMsg ""
    LogMsg "Started " & Now
    LogMsg "Add: " & IIf(Opt_AddDesc, PROP_DESC & " ", "") & IIf(Opt_AddMat, PROP_MAT & " ", "") & _
           IIf(Opt_AddWeight, PROP_WEIGHT & " ", "") & IIf(Opt_AddSize, PROP_SIZE, "") & _
           " | old properties: " & IIf(Opt_DeleteOld, "DELETE", "keep") & _
           " | overwrite: " & Opt_Overwrite & " | save: " & Opt_Save
    For i = 1 To PartCount
        SetStatus "Writing " & fso.GetFileName(pPath(i)) & "  (" & i & " / " & PartCount & ")", (i - 1) / PartCount
        If WritePart(i) Then okCount = okCount + 1 Else failCount = failCount + 1
        lstReview.List(i, 5) = pResult(i)
    Next i

    On Error Resume Next
    TopDoc.EditRebuild3                 ' refresh the BOM
    On Error GoTo 0
    LogMsg ""
    LogMsg "Finished " & Now & ": " & okCount & " updated, " & failCount & " failed, " & GetSkippedCount() & " skipped."
    lastLogPath = SaveLog(TopDoc)

    finished = True
    SetBusy False
    FillReview
    btnNext.Caption = "Open log"
    btnCancel.Caption = "Close"
    btnBack.Enabled = False
    chkOverwrite.Enabled = False
    chkSave.Enabled = False

    ShowProgress okCount & " updated, " & failCount & " failed, " & GetSkippedCount() & " skipped.", 1
    If failCount > 0 Then lblStatus.ForeColor = RED
    If MsgBox(okCount & " part(s) updated, " & failCount & " failed, " & GetSkippedCount() & " skipped." & _
              IIf(lastLogPath <> "", vbCrLf & vbCrLf & "Log: " & lastLogPath & vbCrLf & vbCrLf & "Open the log?", ""), _
              IIf(failCount > 0, vbExclamation, vbInformation) + IIf(lastLogPath <> "", vbYesNo, vbOKOnly), TITLE) = vbYes Then
        OpenLog
    End If
End Sub

Private Sub OpenLog()
    If lastLogPath = "" Or Not fso.FileExists(lastLogPath) Then
        MsgBox "No log file was written.", vbInformation, TITLE
    Else
        Shell "notepad.exe """ & lastLogPath & """", vbNormalFocus
    End If
End Sub

Private Sub SetBusy(ByVal b As Boolean)
    busy = b
    btnBack.Enabled = Not b
    btnNext.Enabled = Not b
    btnCancel.Enabled = Not b
    Me.MousePointer = IIf(b, 11, 0)     ' fmMousePointerHourGlass / default
    Me.Repaint
End Sub
