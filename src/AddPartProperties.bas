Attribute VB_Name = "AddPartProperties"
'===============================================================================
' AddPartProperties.bas  -  SOLIDWORKS VBA macro (engine + entry point)
'
' Adds the custom properties DESCRIPTION, Material, Weight and Size to every
' part used in the active assembly (sub-assemblies included). The wizard lives
' in the frmAddPartProperties UserForm; this module does the work.
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
' See docs/AddPartProperties-spec.md.
'===============================================================================
Option Explicit

'------------------------------ property names ---------------------------------
Public Const PROP_DESC As String = "DESCRIPTION"
Public Const PROP_MAT As String = "Material"
Public Const PROP_WEIGHT As String = "Weight"
Public Const PROP_SIZE As String = "Size"

'------------------------------ options (set by the wizard) --------------------
Public Opt_AddDesc As Boolean
Public Opt_AddMat As Boolean
Public Opt_AddWeight As Boolean
Public Opt_AddSize As Boolean
Public Opt_DeleteOld As Boolean         ' delete every Custom-tab property first
Public Opt_Overwrite As Boolean         ' replace existing values (Keep mode only)
Public Opt_Save As Boolean              ' save each part after writing

Public MatCommon As String              ' material for all parts
Public SizeSameForAll As Boolean        ' True: SizeAll for every part
Public SizeAll As String

'------------------------------ parts (1-based, PartCount entries) --------------
Public PartCount As Long
Public pPath() As String                ' full path of the part file
Public pComp() As SldWorks.Component2   ' one instance in the top assembly
Public pWhere() As String               ' sub-assembly path, for display
Public pConfig() As String              ' referenced configuration
Public pOldDesc() As String             ' existing values ("" = none)
Public pOldMat() As String
Public pOldSize() As String
Public pHasDesc() As Boolean
Public pHasMat() As Boolean
Public pHasSize() As Boolean
Public pHasWeight() As Boolean
Public pPropCount() As Long             ' number of properties on the Custom tab
Public pDesc() As String                ' value typed in the per-part dialog
Public pDescSet() As Boolean            ' False = keep existing (Skip)
Public pSize() As String
Public pSizeSet() As Boolean
Public pMatExc() As Boolean             ' True = exception, uses pMat
Public pMat() As String
Public pResult() As String              ' result after writing

'------------------------------ API constants ----------------------------------
Private Const swDocASSEMBLY As Long = 2
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swComponentSuppressed As Long = 0
Private Const swComponentLightweight As Long = 1
Private Const swComponentFullyResolved As Long = 2
Private Const swComponentFullyLightweight As Long = 4
Private Const swToolboxStandardPart As Long = 1
Private Const ADD_OK As Long = 0                ' swCustomInfoAddResult_AddedOrChanged

'------------------------------ shared state -----------------------------------
Public swApp As SldWorks.SldWorks
Public TopDoc As SldWorks.ModelDoc2     ' the active assembly
Public gUI As Object                   ' the open wizard (receives progress)
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
        MsgBox problem, vbExclamation, "Add Part Properties"
        Exit Sub
    End If

    SetStatus "Collecting parts..."
    CollectParts TopDoc
    SetStatus ""
    If PartCount = 0 Then
        MsgBox "No part in this assembly can be changed." & vbCrLf & vbCrLf & _
               IIf(GetSkippedText() = "", "The assembly has no parts.", "Skipped:" & vbCrLf & Left$(GetSkippedText(), 1500)), _
               vbExclamation, "Add Part Properties"
        Exit Sub
    End If
    frmAddPartProperties.Show vbModeless
End Sub

Public Sub InitEngine()
    If swApp Is Nothing Then Set swApp = Application.SldWorks
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pdmFolders = CreateObject("Scripting.Dictionary")
    pdmFolders.CompareMode = vbTextCompare
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
' Active assembly and its parts
'===============================================================================
Public Function GetTopAssembly(ByRef problem As String) As SldWorks.ModelDoc2
    Dim m As SldWorks.ModelDoc2
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
Public Sub CollectParts(topModel As SldWorks.ModelDoc2)
    Dim seen As Object, pending As Object, vComps As Variant, k As Long, c As SldWorks.Component2
    Dim p As String, asmFolder As String, reason As String, doc As SldWorks.ModelDoc2

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

    Dim swAssy As SldWorks.AssemblyDoc
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
Private Function InstanceSkipReason(c As SldWorks.Component2) As String
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

Private Sub AddPart(c As SldWorks.Component2, ByVal p As String, doc As SldWorks.ModelDoc2)
    Dim i As Long, cpm As SldWorks.CustomPropertyManager, names As Variant
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

Public Function GetSkippedText() As String
    GetSkippedText = skippedText
End Function

Public Function GetSkippedCount() As Long
    GetSkippedCount = skippedCount
End Function

' "SubAsm-1 / SubSub-1" for a part inside sub-assemblies, "" at top level.
Private Function ParentPath(c As SldWorks.Component2) As String
    Dim parent As SldWorks.Component2, s As String
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
Private Function IsEnvelope(c As SldWorks.Component2) As Boolean
    On Error Resume Next
    IsEnvelope = c.IsEnvelope
End Function

Private Function IsReadOnlyFile(ByVal path As String) As Boolean
    On Error Resume Next
    IsReadOnlyFile = (fso.GetFile(path).Attributes And 1) <> 0
End Function

Private Function IsReadOnlyDoc(doc As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    IsReadOnlyDoc = doc.IsOpenedReadOnly Or doc.IsOpenedViewOnly
End Function

Private Function ToolboxType(doc As SldWorks.ModelDoc2) As Long
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
' Planned values (shown in the Review page)
'===============================================================================
' What will happen to property prop of part i. Returns the text for the Review
' table: the value to write, "(unchanged)", "(kept: ...)", "(deleted)" or "".
Public Function PlannedText(ByVal i As Long, ByVal prop As String) As String
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

Public Function PlannedNote(ByVal i As Long) As String
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

Public Function MaterialFor(ByVal i As Long) As String
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
Public Function WritePart(ByVal i As Long) As Boolean
    Dim doc As SldWorks.ModelDoc2, cpm As SldWorks.CustomPropertyManager
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
                cpm.Add3 delNames(k), delTypes(k), delVals(k), swCustomPropertyDeleteAndAdd
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
Private Function CanWrite(cpm As SldWorks.CustomPropertyManager, ByVal prop As String) As Boolean
    If Opt_DeleteOld Or Opt_Overwrite Then
        CanWrite = True
    ElseIf ReadProp(cpm, prop, "") Then
        LogMsg "  " & prop & ": exists, kept (overwrite is off)."
    Else
        CanWrite = True
    End If
End Function

' Writes a Text property and reads it back.
Private Function SetText(cpm As SldWorks.CustomPropertyManager, ByVal prop As String, ByVal value As String) As Boolean
    Dim rc As Long, back As String
    value = Trim$(value)
    rc = cpm.Add3(prop, swCustomInfoText, value, swCustomPropertyDeleteAndAdd)
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
Private Function SetWeight(doc As SldWorks.ModelDoc2, cpm As SldWorks.CustomPropertyManager, ByVal path As String) As Boolean
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

Private Function WriteWeight(cpm As SldWorks.CustomPropertyManager, ByVal expr As String) As Boolean
    Dim rc As Long
    rc = cpm.Add3(PROP_WEIGHT, swCustomInfoText, expr, swCustomPropertyDeleteAndAdd)
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

Private Function RealMassKg(doc As SldWorks.ModelDoc2) As Double
    Dim ext As Object, mp As Object
    On Error Resume Next
    Set ext = doc.Extension
    Set mp = ext.CreateMassProperty
    If Not mp Is Nothing Then RealMassKg = mp.Mass
End Function

Private Function ResolvedValue(cpm As SldWorks.CustomPropertyManager, ByVal prop As String) As String
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
Private Function ReadProp(cpm As SldWorks.CustomPropertyManager, ByVal prop As String, ByRef value As String) As Boolean
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

Public Sub SelectPart(topModel As SldWorks.ModelDoc2, ByVal i As Long, ByVal zoom As Boolean)
    On Error Resume Next
    topModel.ClearSelection2 True
    If pComp(i).Select4(False, Nothing, False) Then
        If zoom Then topModel.ViewZoomToSelection
    End If
    topModel.GraphicsRedraw2
End Sub

' fraction: 0..1 for the progress bar, negative = leave the bar as it is.
Public Sub SetStatus(ByVal text As String, Optional ByVal fraction As Double = -1)
    On Error Resume Next
    swApp.Frame.SetStatusBarText IIf(text = "", "", "Add Part Properties: " & text)
    If Not gUI Is Nothing Then gUI.ShowProgress text, fraction
    DoEvents
End Sub

Public Sub LogMsg(ByVal text As String)
    logText = logText & text & vbCrLf
    Debug.Print text
End Sub

' Writes the log next to the top assembly. Returns its path ("" if not written).
Public Function SaveLog(topModel As SldWorks.ModelDoc2) As String
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
