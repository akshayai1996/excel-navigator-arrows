Option Explicit

'===============================================================================
' Navigator Arrows – Focus Mode (v5.0 - Dashboard Header Edition)
' Features: Frozen Header, "Always Visible" Button, Right-Side Placement
'===============================================================================

'====================== CONFIG ======================
Public Const NAV_SHEET As String = "Navigator_List"
Private Const SHP_LEFT  As String = "NavArrows_Left"
Private Const SHP_RIGHT As String = "NavArrows_Right"
Private Const BTN_NAME  As String = "Nav_Go_Button"

Private Const ARROW_WIDTH_PTS   As Double = 14
Private Const ARROW_MARGIN_PTS  As Double = 1
Private Const MIN_EDGE_COLWIDTH As Double = 3.5
Private Const WRAP_NAVIGATION   As Boolean = True

' New Config: Where does the data start?
Private Const DATA_START_ROW    As Long = 6 
'===================================================

'================== STATE ===========================
Private mDataSheet As Worksheet
'===================================================


'===============================================================================
' PUBLIC – APPLY NAVIGATOR
'===============================================================================
Public Sub NavigatorArrows_Apply()

    Dim caller As String
    caller = GetCallerNameSafe()

    '--- Arrow click dispatcher
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

    '--- SETUP UI & WAIT ---
    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)
    navSh.Visible = xlSheetVisible
    navSh.Activate
    
    ' Select first data item so user is ready
    navSh.Cells(DATA_START_ROW, 1).Select
    
    ' Force Scroll to top
    ActiveWindow.ScrollRow = 1
    
    ' Add the Magic Button
    AddGoButton navSh

End Sub


'===============================================================================
' PUBLIC – REMOVE NAVIGATOR
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
    Application.StatusBar = False

    On Error GoTo 0
End Sub


'===============================================================================
' BUTTON CLICK HANDLER
'===============================================================================
Public Sub NavigatorJumpButton_Click()
    Dim navSh As Worksheet
    Set navSh = ActiveSheet
    
    If navSh.Name <> NAV_SHEET Then Exit Sub

    Dim r As Long
    r = ActiveCell.Row
    
    ' Safety: User must click inside the data area, not the header
    If r < DATA_START_ROW Then
        MsgBox "Please select an item from the list below (Row " & DATA_START_ROW & "+).", vbInformation
        Exit Sub
    End If
    
    Dim idxVal As Variant
    idxVal = navSh.Cells(r, 1).Value ' Col A is Index
    
    If Not IsNumeric(idxVal) Or IsEmpty(idxVal) Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If
    
    JumpToIndex CLng(idxVal)
End Sub


'===============================================================================
' NAVIGATION CORE
'===============================================================================
Private Sub JumpToIndex(ByVal idx As Long)
    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)
    If navSh Is Nothing Then Exit Sub

    navSh.Range("E1").Value = idx
    ApplySingleRowFocus CLng(navSh.Cells(idx + DATA_START_ROW - 1, 2).Value) ' Adjusted for offset
    UpdateStatusBar idx
    
    navSh.Visible = xlSheetVeryHidden
End Sub

Private Sub NavigateByIndex(ByVal direction As Long)
    Dim navSh As Worksheet
    Set navSh = Worksheets(NAV_SHEET)

    Dim curIdx As Variant
    curIdx = navSh.Range("E1").Value
    If Not IsNumeric(curIdx) Then Exit Sub

    Dim lastIdx As Long
    ' Calculate last index based on data count
    lastIdx = navSh.Cells(navSh.Rows.Count, 1).End(xlUp).Row - (DATA_START_ROW - 1)
    If lastIdx <= 0 Then Exit Sub

    Dim idx As Long
    idx = CLng(curIdx) + direction

    If idx < 1 Then idx = IIf(WRAP_NAVIGATION, lastIdx, 1)
    If idx > lastIdx Then idx = IIf(WRAP_NAVIGATION, 1, lastIdx)

    navSh.Range("E1").Value = idx
    ApplySingleRowFocus CLng(navSh.Cells(idx + DATA_START_ROW - 1, 2).Value)
    UpdateStatusBar idx

    If navSh.Visible <> xlSheetVeryHidden Then
        navSh.Visible = xlSheetVeryHidden
    End If
End Sub


'===============================================================================
' UI BUILDER (UPDATED FOR HEADER)
'===============================================================================
Private Sub BuildNavigatorList(ByVal rngFilter As Range)

    Dim navSh As Worksheet
    Set navSh = GetOrCreateNavigatorSheet()
    
    navSh.Cells.Clear
    navSh.DrawingObjects.Delete
    
    ' Clear Freeze Panes first to reset view
    navSh.Activate
    ActiveWindow.FreezePanes = False
    
    ' Store Metadata in hidden top row
    navSh.Range("E1").Value = 1
    navSh.Range("F1").Value = rngFilter.Parent.Name
    
    ' Draw Headers at Row 5 (Just above data)
    Dim hdrRow As Long
    hdrRow = DATA_START_ROW - 1
    navSh.Cells(hdrRow, 1).Resize(1, 3).Value = Array("Index", "Row", "Item")
    navSh.Cells(hdrRow, 1).Resize(1, 3).Font.Bold = True
    navSh.Cells(hdrRow, 1).Resize(1, 3).Interior.Color = RGB(240, 240, 240)

    ' Get Data
    Dim dataRng As Range
    Set dataRng = rngFilter.Offset(1).Resize(rngFilter.Rows.Count - 1)

    Dim rowsList() As Long
    rowsList = GetVisibleRowNumbers(dataRng)
    If UBound(rowsList) < LBound(rowsList) Then Exit Sub

    Dim n As Long
    n = UBound(rowsList) + 1

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 3)

    Dim i As Long
    For i = 1 To n
        arr(i, 1) = i
        arr(i, 2) = rowsList(i - 1)
        arr(i, 3) = mDataSheet.Cells(rowsList(i - 1), rngFilter.Column).Value
    Next i

    ' Paste Data starting at DATA_START_ROW
    navSh.Cells(DATA_START_ROW, 1).Resize(n, 3).Value = arr
    navSh.Columns("A:C").AutoFit
    
    '--- FREEZE PANES ---
    ' We select cell A6 (Start of data) and freeze everything above it
    navSh.Cells(DATA_START_ROW, 1).Select
    ActiveWindow.FreezePanes = True
    '--------------------
End Sub

Private Sub AddGoButton(ByVal sh As Worksheet)
    ' Places button in the Header Area (Rows 2-4), Column F-H
    ' This ensures it is "Right Side" and "Always Visible"
    
    Dim btn As Shape
    Dim targetRange As Range
    
    ' Position: Columns F to H, Rows 2 to 4
    Set targetRange = sh.Range("F2:H4") 
    
    Set btn = sh.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    
    With btn
        .Name = BTN_NAME
        .OnAction = "NavigatorJumpButton_Click"
        .Fill.ForeColor.RGB = RGB(0, 112, 192) ' Professional Blue
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow25
        
        With .TextFrame2
            .TextRange.Text = "GO / JUMP"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        .Placement = xlFreeFloating
    End With
    
    ' Instruction Text in Header
    With sh.Range("F1")
        .Value = "Select a row below, then click:"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
End Sub


'===============================================================================
' HELPERS
'===============================================================================
Private Sub ApplySingleRowFocus(ByVal targetRow As Long)
    If mDataSheet Is Nothing Then Exit Sub
    
    Dim navSh As Worksheet: Set navSh = Worksheets(NAV_SHEET)
    ' Recalculate range based on new data start
    Dim lastRowList As Long
    lastRowList = navSh.Cells(navSh.Rows.Count, 2).End(xlUp).Row
    
    If lastRowList < DATA_START_ROW Then Exit Sub

    ' Get the MIN and MAX row numbers from the list column (Col B)
    ' This is safer than assuming sorted order
    Dim firstRow As Long, lastRow As Long
    ' Simple approximation: min is usually top, max is bottom
    firstRow = CLng(navSh.Cells(DATA_START_ROW, 2).Value)
    lastRow = CLng(navSh.Cells(lastRowList, 2).Value)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    mDataSheet.Rows(firstRow & ":" & lastRow).Hidden = True
    mDataSheet.Rows(targetRow).Hidden = False

    mDataSheet.Activate
    mDataSheet.Cells(targetRow, 1).Select
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub ExitFocus_Internal()
    If mDataSheet Is Nothing Then Exit Sub
    Dim navSh As Worksheet: Set navSh = Worksheets(NAV_SHEET)
    Dim lastRowList As Long: lastRowList = navSh.Cells(navSh.Rows.Count, 2).End(xlUp).Row
    If lastRowList < DATA_START_ROW Then Exit Sub
    
    Dim firstRow As Long: firstRow = CLng(navSh.Cells(DATA_START_ROW, 2).Value)
    Dim lastRow As Long: lastRow = CLng(navSh.Cells(lastRowList, 2).Value)
    
    Application.ScreenUpdating = False
    mDataSheet.Rows(firstRow & ":" & lastRow).Hidden = False
    Application.ScreenUpdating = True
End Sub

Private Function RestoreState() As Boolean
    If Not mDataSheet Is Nothing Then
        RestoreState = True
        Exit Function
    End If
    On Error Resume Next
    Dim navSh As Worksheet: Set navSh = Worksheets(NAV_SHEET)
    If Not navSh Is Nothing Then
        Dim nm As String: nm = CStr(navSh.Range("F1").Value)
        If nm <> "" Then
            Set mDataSheet = Worksheets(nm)
            RestoreState = Not mDataSheet Is Nothing
        End If
    End If
    On Error GoTo 0
End Function

Private Sub UpdateStatusBar(ByVal idx As Long)
    Dim navSh As Worksheet: Set navSh = Worksheets(NAV_SHEET)
    Dim total As Long: total = navSh.Cells(navSh.Rows.Count, 1).End(xlUp).Row - (DATA_START_ROW - 1)
    Application.StatusBar = "Navigator: " & idx & " / " & total
End Sub

Private Function GetVisibleRowNumbers(ByVal rng As Range) As Long()
    Dim rows() As Long, cnt As Long
    Dim vis As Range, a As Range, r As Range
    On Error Resume Next
    Set vis = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not vis Is Nothing Then
        ReDim rows(0 To vis.Cells.Count - 1)
        For Each a In vis.Areas
            For Each r In a.rows
                rows(cnt) = r.Row
                cnt = cnt + 1
            Next r
        Next a
    Else
        ReDim rows(0 To -1)
    End If
    If cnt > 0 Then ReDim Preserve rows(0 To cnt - 1)
    GetVisibleRowNumbers = rows
End Function

Private Sub PlaceOrUpdateArrows(ByVal sh As Worksheet, ByVal rngFilter As Range)
    Dim headerRow As Long: headerRow = rngFilter.Row
    Dim firstCol As Long: firstCol = rngFilter.Column
    Dim lastCol As Long: lastCol = firstCol + rngFilter.Columns.Count - 1
    
    Dim cLeft As Range: Set cLeft = sh.Cells(headerRow, firstCol)
    Dim cRight As Range: Set cRight = sh.Cells(headerRow, lastCol)
    Dim h As Double: h = cLeft.Height - 2
    
    DeleteShapeIfExists sh, SHP_LEFT
    DeleteShapeIfExists sh, SHP_RIGHT
    
    With sh.Shapes.AddShape(msoShapeLeftArrow, cLeft.Left + 1, cLeft.Top + 1, ARROW_WIDTH_PTS, h)
        .Name = SHP_LEFT: .OnAction = "NavigatorArrows_Apply": .Placement = xlMoveAndSize
    End With
    With sh.Shapes.AddShape(msoShapeRightArrow, cRight.Left + cRight.Width - ARROW_WIDTH_PTS - 1, cRight.Top + 1, ARROW_WIDTH_PTS, h)
        .Name = SHP_RIGHT: .OnAction = "NavigatorArrows_Apply": .Placement = xlMoveAndSize
    End With
End Sub

Private Function GetCallerNameSafe() As String
    If TypeName(Application.Caller) = "String" Then GetCallerNameSafe = Application.Caller
End Function

Private Function GetFilterRangeOrRegion(ByVal sh As Worksheet) As Range
    If sh.AutoFilterMode And Not sh.AutoFilter Is Nothing Then
        Set GetFilterRangeOrRegion = sh.AutoFilter.Range
        Exit Function
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
    End If
    Set GetOrCreateNavigatorSheet = sh
End Function

Private Sub DeleteShapeIfExists(ByVal sh As Worksheet, ByVal shpName As String)
    On Error Resume Next
    sh.Shapes(shpName).Delete
    On Error GoTo 0
End Sub
