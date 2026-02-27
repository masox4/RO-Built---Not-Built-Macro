Attribute VB_Name = "RO_StockList_Macros"
Option Explicit

Private Const MAIN_SHEET As String = "RO StockList"
Private Const RAW_SHEET As String = "RAW RO StockList"

' Run this once to create/recreate all action buttons.
Public Sub SetupActionButtons()
    AddOrReplaceButton MAIN_SHEET, "F1", "btnClearMain", "Clear Old Data", "ClearOldData_Main"
    AddOrReplaceButton RAW_SHEET, "E1", "btnClearRaw", "Clear Old Data", "ClearOldData_Raw"
    AddOrReplaceButton MAIN_SHEET, "K1", "btnImportRO", "Import RO StockList", "ImportROStockList"
    AddOrReplaceButton MAIN_SHEET, "L1", "btnShowNotBuilt", "Show NOT Built RO Poducts", "ShowNotBuiltROProducts", RGB(0, 97, 0)
    AddOrReplaceButton MAIN_SHEET, "M1", "btnUndoShowNotBuilt", "Undo Show NOT Built", "UndoShowNotBuilt"

    MsgBox "Buttons added/updated on '" & MAIN_SHEET & "' and '" & RAW_SHEET & "'.", vbInformation
End Sub

' Backward compatible entry point name.
Public Sub SetupClearButtons()
    SetupActionButtons
End Sub

' Clears RO StockList data from row 2, columns A:AD,
' down to the first blank SKU row in column A.
Public Sub ClearOldData_Main()
    ClearDataUntilBlankRow ThisWorkbook.Worksheets(MAIN_SHEET), "A", "A", "AD", 2
    MsgBox "Old data cleared from " & MAIN_SHEET & ".", vbInformation
End Sub

' Clears RAW RO StockList data from row 2, columns A:AE,
' down to the first blank SKU row in column A.
Public Sub ClearOldData_Raw()
    ClearDataUntilBlankRow ThisWorkbook.Worksheets(RAW_SHEET), "A", "A", "AE", 2
    MsgBox "Old data cleared from " & RAW_SHEET & ".", vbInformation
End Sub

' Imports RAW RO data into RO StockList with requested formula columns.
Public Sub ImportROStockList()
    Dim wsMain As Worksheet
    Dim wsRaw As Worksheet
    Dim lastRawRow As Long
    Dim lastMainRow As Long

    Set wsMain = ThisWorkbook.Worksheets(MAIN_SHEET)
    Set wsRaw = ThisWorkbook.Worksheets(RAW_SHEET)

    ' Guard: require old data to be cleared first.
    If Trim$(CStr(wsMain.Cells(2, "A").Value)) <> "" Then
        MsgBox "Please click 'Clear Old Data' on " & MAIN_SHEET & " before importing.", vbExclamation
        Exit Sub
    End If

    lastRawRow = LastContiguousRow(wsRaw, "A", 2)
    If lastRawRow < 2 Then
        MsgBox "No data found to import on " & RAW_SHEET & " (from row 2 down).", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.CutCopyMode = False

    ' Ensure destination is clean.
    wsMain.Range("A2:AD" & wsMain.Rows.Count).ClearContents

    ' Copy values + number formats so IDs/EAN formatting is preserved.
    wsRaw.Range("A2:A" & lastRawRow).Copy
    wsMain.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats

    ' Shift source B:AA to destination C:AB (inserts Built/Not Built at column B).
    wsRaw.Range("B2:AA" & lastRawRow).Copy
    wsMain.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats

    Application.CutCopyMode = False

    lastMainRow = lastRawRow

    ' Built / Not Built formula in column B.
    wsMain.Range("B2:B" & lastMainRow).Formula = "=IF(COUNTIF('EYES Invent'!A:A, A2)>0, ""Built"", ""Not Built"")"

    ' Ignore formula in column AC.
    wsMain.Range("AC2:AC" & lastMainRow).Formula = "=IF(COUNTIF('Odd Sizes'!A:A, A2)>0, ""Ignore"", ""-"")"

    ' Missed formula in column AD.
    wsMain.Range("AD2:AD" & lastMainRow).Formula = "=IF(COUNTIF('Missing Sizes'!A:A, A2)>0, ""Missed"", ""-"")"

    Application.ScreenUpdating = True

    MsgBox "Import complete: " & (lastMainRow - 1) & " row(s) imported to " & MAIN_SHEET & ".", vbInformation
End Sub

Public Sub ShowNotBuiltROProducts()
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range

    Set wsMain = ThisWorkbook.Worksheets(MAIN_SHEET)
    lastRow = LastContiguousRow(wsMain, "A", 2)

    If lastRow < 2 Then
        MsgBox "No rows found to filter on " & MAIN_SHEET & ".", vbExclamation
        Exit Sub
    End If

    Set filterRange = wsMain.Range("A1:AD" & lastRow)

    If wsMain.AutoFilterMode Then
        On Error Resume Next
        wsMain.ShowAllData
        On Error GoTo 0
    End If

    filterRange.AutoFilter Field:=29, Criteria1:="-"
    filterRange.AutoFilter Field:=30, Criteria1:="-"
    filterRange.AutoFilter Field:=2, Criteria1:="<>Built"
End Sub

Public Sub UndoShowNotBuilt()
    Dim wsMain As Worksheet

    Set wsMain = ThisWorkbook.Worksheets(MAIN_SHEET)

    On Error Resume Next
    If wsMain.FilterMode Then wsMain.ShowAllData
    On Error GoTo 0

    If wsMain.AutoFilterMode Then wsMain.AutoFilterMode = False
End Sub

Private Sub ClearDataUntilBlankRow(ByVal ws As Worksheet, ByVal keyCol As String, ByVal fromCol As String, ByVal toCol As String, ByVal startRow As Long)
    Dim lastRow As Long

    lastRow = LastContiguousRow(ws, keyCol, startRow)

    If lastRow < startRow Then Exit Sub

    ws.Range(fromCol & startRow & ":" & toCol & lastRow).ClearContents
End Sub

Private Function LastContiguousRow(ByVal ws As Worksheet, ByVal keyCol As String, ByVal startRow As Long) As Long
    Dim r As Long

    r = startRow
    Do While Trim$(CStr(ws.Cells(r, keyCol).Value)) <> ""
        r = r + 1
    Loop

    LastContiguousRow = r - 1
End Function

Private Sub AddOrReplaceButton(ByVal sheetName As String, ByVal anchorCell As String, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, Optional ByVal fillColor As Variant)
    Dim ws As Worksheet
    Dim target As Range
    Dim btn As Shape
    Dim w As Double
    Dim h As Double
    Dim chosenColor As Long
    Dim horizontalPadding As Double

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set target = ws.Range(anchorCell)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    h = 18
    horizontalPadding = 6

    chosenColor = RGB(0, 51, 102)
    If Not IsMissing(fillColor) Then chosenColor = CLng(fillColor)

    ' Create with a temporary width, then resize to text width + equal left/right padding.
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.Left + 1, target.Top + 1, 60, h)

    With btn
        .Name = shapeName
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        .TextFrame2.MarginLeft = horizontalPadding
        .TextFrame2.MarginRight = horizontalPadding
        .TextFrame2.MarginTop = 2
        .TextFrame2.MarginBottom = 2
        .TextFrame2.WordWrap = msoFalse

        w = .TextFrame2.TextRange.BoundWidth + (.TextFrame2.MarginLeft + .TextFrame2.MarginRight)
        If w < 46 Then w = 46
        .Width = w
        .Height = h

        .Fill.ForeColor.RGB = chosenColor
        .Line.Visible = msoFalse
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    End With
End Sub
