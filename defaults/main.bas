Private TARGET_SHEET_NAMES As Variant
Private Sub InitializeTargetSheetNames()
    TARGET_SHEET_NAMES = Array("CS女子給", "BS女子給", "HS女子給", "JS女子給", "GS女子給")
End Sub

Sub Execute()
  Dim sheetName As Variant
  Dim sheet As Worksheet
  Dim rebatePrimaryCategoryStartColumn As Long 
  Dim variablePayPrimaryCategoryStartColumn As Long
  Dim tmpSheet As Worksheet
  Dim currentSheetLastColumn As Long
  Dim currentTmpSheetLastColumn As Long

  Call InitializeTargetSheetNames
  For Each sheetName In TARGET_SHEET_NAMES
    On Error Resume Next
    Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If sheet Is Nothing Then
        Debug.Print sheetName & " does not exist."
        Set sheet = Nothing '
    Else
      rebatePrimaryCategoryStartColumn = FindColumnByHeader(sheet, PRIMARY_CATEGORY_ROW_NUM, REBATE_PRIMARY_CATEGORY)
      variablePayPrimaryCategoryStartColumn = FindColumnByHeader(sheet, PRIMARY_CATEGORY_ROW_NUM, VARIABLE_PAY_PRIMARY_CATEGORY)

      Set tmpSheet = PrepareTmpSheet()

      ' バックまでの列（支給額等~基本給の列）を一時的なシートにコピー
      Call CopyColumnsToOtherSheet(sheet, tmpSheet, 1, rebatePrimaryCategoryStartColumn - 1, 1)

      ' バックを期待する順序に並び替えて一時的なシートにコピー
      Call SortAndCopyColumns(sheet, tmpSheet)
      ' 期待する並び順以外のものが含まれていた場合は不明な項目としてコピー
      Call CopyIrrelevantColumns(sheet, tmpSheet, CLng(rebatePrimaryCategoryStartColumn), CLng(variablePayPrimaryCategoryStartColumn - 1))

      ' 最後に変動給のところを一時的なシートにコピー
      currentSheetLastColumn = sheet.Cells(1, sheet.Columns.Count).End(xlToLeft).Column
      currentTmpSheetLastColumn = tmpSheet.Cells(1, tmpSheet.Columns.Count).End(xlToLeft).Column
      Call CopyColumnsToOtherSheet(sheet, tmpSheet, variablePayPrimaryCategoryStartColumn, currentSheetLastColumn, currentTmpSheetLastColumn + 1)

      ' 元のシートに一時的なシートの内容をコピー
      Call OverwriteOriginalSheet(sheet, tmpSheet)

      ' 一時的なシートを削除
      Call Cleanup

      Set sheet = Nothing
    End If
  Next sheetName

  Set sheet = ThisWorkbook.Sheets("GS女子給")
End Sub

