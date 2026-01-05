Attribute VB_Name = "modNavigatorArrows"
Option Explicit

'===============================================================================
' Navigator Arrows – Focus Mode (GOLD RELEASE)
' Excel 2010 → M365 | 1k–50k+ rows | State-loss & edge-case safe
'
' Author : Akshay Solanki 
'===============================================================================

'====================== CONFIG ======================
Private Const SHP_LEFT  As String = "NavArrows_Left"
Private Const SHP_RIGHT As String = "NavArrows_Right"
Private Const NAV_SHEET As String = "Navigator_List"

Private Const ARROW_WIDTH_PTS   As Double = 14
Private Const ARROW_MARGIN_PTS  As Double = 1
Private Const MIN_EDGE_COLWIDTH As Double = 3.5
Private Const WRAP_NAVIGATION   As Boolean = True
'===================================================

'================== VOLATILE STATE ==================
Private mDataSheet As Worksheet   ' RAM (can reset)
'===================================================


'===============================================================================
' PUBLIC – APPLY (BUILD + ENTER FOCUS)
'===============================================================================
Public Sub NavigatorArrows_Apply()
    Dim caller As String
    caller = GetCallerNameSafe()

    '--- Arrow dispatcher
    If caller = SHP_LEFT Or caller = SHP_RIGHT Then
        If Not RestoreState() Then
            MsgBox "Navigator not initialized. Run Apply again.", vbInformation
            Exit Sub
        End If

        If caller = SHP_LEFT Then NavigateByIndex -1
        If caller = SHP_RIGHT Then NavigateByIndex 1
        Exit Sub
    End If

    '--- Normal Apply
    Dim rngFilter As Range
    Set rngFilter = GetFilterRangeOrRegion(ActiveSheet)
    If rngFilter Is Nothing Then
        MsgBox "Place cursor inside filtered data.", vbExclamation
        Exit Sub
    End If

    Set mDataSheet = rngFilter.Parent

    PlaceOrUpdateArrows mDataSheet, rngFilter
    BuildNavigatorList rngFilter

    NavigateByIndex 0
End Sub


'===============================================================================
' PUBLIC – REMOVE (EXIT FOCUS + CLEAN UI)
'===============================================================================
Public Sub NavigatorArrows_Remove()
    On Error Resume Next

    RestoreState
    ExitFocus_Internal

    If Not mDataSheet Is Nothing Then
        mDataSheet.Shapes(SHP_LEFT).Delete
        mDataSheet.Shapes(SHP_RIGHT).Delete
    End If

    Worksheets(NAV_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub


'===============================================================================
' STATE RECOVERY (NON-VOLATILE FAILSAFE)
'===============================================================================
Private Function RestoreState() As Boolean
    If Not mDataSheet Is Nothing Then
        RestoreState = True
        Exit Function
    End If

    On Error Resume Next
    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)

    If Not navSh Is Nothing Then
        Dim sheetName As String
        sheetName = CStr(navSh.Range("F1").Value)

        If sheetName <> vbNullString Then
            Set mDataSheet = Nothing
            Set mDataSheet = Worksheets(sheetName)
            RestoreState = (Not mDataSheet Is Nothing)
        End If
    End If
    On Error GoTo 0
End Function


'===============================================================================
' NAVIGATION CORE
'===============================================================================
Private Sub NavigateByIndex(ByVal direction As Long)
    On Error GoTo SafeExit

    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)

    Dim curIdx As Variant
    curIdx = navSh.Range("E1").Value

    If Not IsNumeric(curIdx) Then Exit Sub
    If curIdx <> Fix(curIdx) Then Exit Sub

    Dim lastIdx As Long
    lastIdx = navSh.Cells(navSh.Rows.Count, 1).End(xlUp).Row - 1
    If lastIdx <= 0 Then Exit Sub

    Dim idx As Long
    idx = CLng(curIdx) + direction

    If idx < 1 Then idx = IIf(WRAP_NAVIGATION, lastIdx, 1)
    If idx > lastIdx Then idx = IIf(WRAP_NAVIGATION, 1, lastIdx)

    navSh.Range("E1").Value = idx
    ApplySingleRowFocus CLng(navSh.Cells(idx + 1, 2).Value)

SafeExit:
End Sub


'===============================================================================
' FOCUS MODE – TURBO (O(1))
'===============================================================================
Private Sub ApplySingleRowFocus(ByVal targetRow As Long)
    On Error GoTo CleanUp
    If mDataSheet Is Nothing Then Exit Sub
    If targetRow < 1 Or targetRow > mDataSheet.Rows.Count Then Exit Sub

    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)

    Dim firstRow As Long, lastRow As Long
    firstRow = CLng(navSh.Cells(2, 2).Value)
    lastRow = CLng(navSh.Cells(navSh.Rows.Count, 2).End(xlUp).Value)

    If firstRow < 1 Or lastRow < firstRow Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    mDataSheet.Rows(firstRow & ":" & lastRow).Hidden = True
    mDataSheet.Rows(targetRow).Hidden = False
    mDataSheet.Rows(targetRow).Select

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'===============================================================================
' EXIT FOCUS – RESTORE ALL
'===============================================================================
Private Sub ExitFocus_Internal()
    On Error Resume Next
    If mDataSheet Is Nothing Then Exit Sub

    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)

    Dim firstRow As Long, lastRow As Long
    firstRow = CLng(navSh.Cells(2, 2).Value)
    lastRow = CLng(navSh.Cells(navSh.Rows.Count, 2).End(xlUp).Value)

    If firstRow < 1 Or lastRow < firstRow Then Exit Sub

    Application.ScreenUpdating = False
    mDataSheet.Rows(firstRow & ":" & lastRow).Hidden = False
    Application.ScreenUpdating = True
End Sub


'===============================================================================
' BUILD NAVIGATOR LIST (PERSISTENT)
'===============================================================================
Private Sub BuildNavigatorList(ByVal rngFilter As Range)
    Dim navSh As Worksheet
    Set navSh = GetOrCreateNavigatorSheet()

    navSh.Cells.Clear
    navSh.Range("A1:C1").Value = Array("Index", "Row", "Item")

    navSh.Range("F1").Value = rngFilter.Parent.Name
    navSh.Range("E1").Value = 1

    Dim dataRng As Range
    Set dataRng = rngFilter.Offset(1).Resize(rngFilter.Rows.Count - 1)

    Dim rowsList() As Long
    rowsList = GetVisibleRowNumbers(dataRng)

    If UBound(rowsList) < LBound(rowsList) Then Exit Sub

    Dim n As Long
    n = UBound(rowsList) + 1
    If n <= 0 Then Exit Sub

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 3)

    Dim i As Long
    For i = 1 To n
        outArr(i, 1) = i
        outArr(i, 2) = rowsList(i - 1)
        outArr(i, 3) = mDataSheet.Cells(rowsList(i - 1), rngFilter.Column).Value
    Next i

    navSh.Range("A2").Resize(n, 3).Value = outArr
End Sub


'===============================================================================
' GET VISIBLE ROW NUMBERS (FILTER-ACCURATE + SAFE FALLBACK)
'===============================================================================
Private Function GetVisibleRowNumbers(ByVal rng As Range) As Long()
    Dim rows() As Long, count As Long
    Dim vis As Range, a As Range, r As Range

    On Error Resume Next
    Set vis = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not vis Is Nothing Then
        ReDim rows(0 To vis.Cells.Count - 1)
        count = 0
        For Each a In vis.Areas
            For Each r In a.Rows
                rows(count) = r.Row
                count = count + 1
            Next r
        Next a
    Else
        ReDim rows(0 To rng.Rows.Count - 1)
        count = 0
        Dim rr As Long
        For rr = rng.Row To rng.Row + rng.Rows.Count - 1
            If Not rng.Parent.Rows(rr).Hidden Then
                rows(count) = rr
                count = count + 1
            End If
        Next rr
    End If

    If count = 0 Then
        ReDim rows(0 To -1)
    Else
        ReDim Preserve rows(0 To count - 1)
    End If

    GetVisibleRowNumbers = rows
End Function


'===============================================================================
' UI – ARROWS
'===============================================================================
Private Sub PlaceOrUpdateArrows(ByVal sh As Worksheet, ByVal rngFilter As Range)
    Dim headerRow As Long: headerRow = rngFilter.Row
    Dim firstCol As Long: firstCol = rngFilter.Column
    Dim lastCol As Long: lastCol = firstCol + rngFilter.Columns.Count - 1

    If sh.Columns(firstCol).ColumnWidth < MIN_EDGE_COLWIDTH Then sh.Columns(firstCol).ColumnWidth = MIN_EDGE_COLWIDTH
    If sh.Columns(lastCol).ColumnWidth < MIN_EDGE_COLWIDTH Then sh.Columns(lastCol).ColumnWidth = MIN_EDGE_COLWIDTH

    Dim cLeft As Range: Set cLeft = sh.Cells(headerRow, firstCol)
    Dim cRight As Range: Set cRight = sh.Cells(headerRow, lastCol)

    Dim arrowH As Double
    arrowH = cLeft.Height - (2 * ARROW_MARGIN_PTS)
    If arrowH < 10 Then arrowH = 10

    DeleteShapeIfExists sh, SHP_LEFT
    DeleteShapeIfExists sh, SHP_RIGHT

    With sh.Shapes.AddShape(msoShapeLeftArrow, cLeft.Left + ARROW_MARGIN_PTS, cLeft.Top + ARROW_MARGIN_PTS, ARROW_WIDTH_PTS, arrowH)
        .Name = SHP_LEFT
        .OnAction = "NavigatorArrows_Apply"
        .Placement = xlMoveAndSize
    End With

    With sh.Shapes.AddShape(msoShapeRightArrow, cRight.Left + cRight.Width - ARROW_WIDTH_PTS - ARROW_MARGIN_PTS, cRight.Top + ARROW_MARGIN_PTS, ARROW_WIDTH_PTS, arrowH)
        .Name = SHP_RIGHT
        .OnAction = "NavigatorArrows_Apply"
        .Placement = xlMoveAndSize
    End With
End Sub


'===============================================================================
' HELPERS
'===============================================================================
Private Function GetCallerNameSafe() As String
    If TypeName(Application.Caller) = "String" Then GetCallerNameSafe = Application.Caller
End Function

Private Function GetFilterRangeOrRegion(ByVal sh As Worksheet) As Range
    If sh.AutoFilterMode Then
        If Not sh.AutoFilter Is Nothing Then
            Set GetFilterRangeOrRegion = sh.AutoFilter.Range
            Exit Function
        End If
    End If
    Set GetFilterRangeOrRegion = ActiveCell.CurrentRegion
End Function

Private Function GetOrCreateNavigatorSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = Worksheets(NAV_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = Worksheets.Add
        sh.Name = NAV_SHEET
        sh.Visible = xlSheetVeryHidden
    End If

    Set GetOrCreateNavigatorSheet = sh
End Function

Private Sub DeleteShapeIfExists(ByVal sh As Worksheet, ByVal shpName As String)
    On Error Resume Next
    sh.Shapes(shpName).Delete
    On Error GoTo 0
End Sub
