Function FindColumnByHeader(sheet As Worksheet, headerRow As Long, headerText As String) As Long
    Dim lastColumn As Long
    Dim i As Long
    
    If sheet Is Nothing Then
        FindColumnByHeader = -1
        Exit Function
    End If
    
    ' 指定された行で使用されている最後の列を取得
    lastColumn = sheet.Cells(headerRow, sheet.Columns.Count).End(xlToLeft).Column
    
    ' 指定されたヘッダーテキストを探す
    For i = 1 To lastColumn
        If sheet.Cells(headerRow, i).Value = headerText Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i
    
    ' ヘッダーが見つからなかった場合
    FindColumnByHeader = -1
End Function


Function PrepareTmpSheet() As Worksheet
    Dim tmpSheet As Worksheet
    Dim sheetExists As Boolean
    Dim i As Integer

    sheetExists = False

    ' シートが既に存在するか確認
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = TMP_SHEET_NAME Then
            sheetExists = True
            Exit For
        End If
    Next i

    ' シートが存在する場合は削除
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(TMP_SHEET_NAME).Delete
        Application.DisplayAlerts = True
    End If

    ' 新しいシートを追加して名前を設定
    Set tmpSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tmpSheet.Name = TMP_SHEET_NAME

    ' 作成したシートを返す
    Set PrepareTmpSheet = tmpSheet
End Function

' 指定された範囲の列をコピーする関数
Sub CopyColumnsToOtherSheet(sheet As Worksheet, destinationSheet As Worksheet, ByVal fromColumn As Long, ByVal toColumn As Long, ByVal destinationColumn As Long)
    Dim lastRow As Long

    ' sheetの使用されている最終行を取得
    lastRow = sheet.Cells(sheet.Rows.Count, fromColumn).End(xlUp).Row

    ' 指定された列数分を一時的なシートにコピー
    sheet.Range(sheet.Cells(1, fromColumn), sheet.Cells(lastRow, toColumn)).Copy Destination:=destinationSheet.Cells(1, destinationColumn)
End Sub

Sub SortAndCopyColumns(sheet As Worksheet, tmpSheet As Worksheet)
    Dim categories As Collection
    Dim secondaryHeaders As Variant
    Dim currentTmpSheetLastColumn As Long
    Dim destinationColumn As Long
    Dim primaryCategory As Variant
    Dim secondaryCategory As Variant
    Dim currentColumn As Long
    Dim primaryCategoryLoopIdx As Long

    Set categories = DesiredOrderColumnStructure()
    secondaryHeaders = GetCurrentSecondaryHeaders(sheet)
    currentTmpSheetLastColumn = tmpSheet.Cells(1, tmpSheet.Columns.Count).End(xlToLeft).Column
    destinationColumn = currentTmpSheetLastColumn + 1
    
    For Each primaryCategory In categories
        primaryCategoryLoopIdx = 0
        For Each secondaryCategory In primaryCategory.Children
            currentColumn = FindInArray(secondaryHeaders, secondaryCategory.Name)
            
            ' 並び替え対象の既存シートから列を一時的なシートへコピー
            If currentColumn > 0 Then
                Call CopyColumnsToOtherSheet(sheet, tmpSheet, currentColumn, currentColumn, destinationColumn)
                ' 既存の大項目名は削除（空白に）
                tmpSheet.Cells(PRIMARY_CATEGORY_ROW_NUM, destinationColumn).Value = ""
            Else
                ' 見つからなかった場合、小項目の名前を1行目に設定
                tmpSheet.Cells(SECONDARY_CATEGORY_ROW_NUM, destinationColumn).Value = secondaryCategory.Name
            End If
            
            ' childrenの先頭のみ、大項目の新しい名前を2行目に設定
            If primaryCategoryLoopIdx = 0 Then
                tmpSheet.Cells(PRIMARY_CATEGORY_ROW_NUM, destinationColumn).Value = primaryCategory.Name
            End If

            destinationColumn = destinationColumn + 1
            primaryCategoryLoopIdx = primaryCategoryLoopIdx + 1
        Next secondaryCategory
    Next primaryCategory
End Sub

Sub CopyIrrelevantColumns(sheet As Worksheet, tmpSheet As Worksheet, targetRangeStart As Long, targetRangeEnd As Long)
    Dim i As Long
    Dim secondaryHeaders As Variant
    Dim targetSecondaryHeaders As Variant
    Dim irrelevantHeaders As New Collection
    Dim irrelevantHeader As Variant 
    Dim currentTmpSheetLastColumn As Long
    Dim destinationColumn As Long
    Dim currentColumn As Long
    Dim desiredSecondaryColumns As Variant
    Dim irrelevantHeaderLoopIdx As Long
    
    secondaryHeaders = GetCurrentSecondaryHeaders(sheet)
    targetSecondaryHeaders = GetCurrentSecondaryHeaders(sheet, targetRangeStart, targetRangeEnd)
    desiredSecondaryColumns = DesiredOrderSecondaryColumns()
    currentTmpSheetLastColumn = tmpSheet.Cells(1, tmpSheet.Columns.Count).End(xlToLeft).Column
    destinationColumn = currentTmpSheetLastColumn + 1
    irrelevantHeaderLoopIdx = 0

    ' 不要なヘッダーを特定
    For i = LBound(targetSecondaryHeaders) To UBound(targetSecondaryHeaders)
        If Not IsInArray(targetSecondaryHeaders(i), desiredSecondaryColumns) Then
            irrelevantHeaders.Add targetSecondaryHeaders(i)
        End If
    Next i
    
    ' 不要な列をコピー
    For Each irrelevantHeader In irrelevantHeaders
        currentColumn = FindInArray(secondaryHeaders, CStr(irrelevantHeader))
        If currentColumn > 0 Then
            ' 並び替え対象の既存のシートから列を一時的なシートへコピー
            Call CopyColumnsToOtherSheet(sheet, tmpSheet, currentColumn, currentColumn, destinationColumn)
            
            ' 既存の大項目名は削除（空白に）
            tmpSheet.Cells(PRIMARY_CATEGORY_ROW_NUM, destinationColumn).Value = ""
            
            ' 最初の不要な列にだけ大項目名「不明」を設定
            If irrelevantHeaderLoopIdx = 0 Then
                tmpSheet.Cells(PRIMARY_CATEGORY_ROW_NUM, destinationColumn).Value = "不明"
            End If

            destinationColumn = destinationColumn + 1
        End If

        irrelevantHeaderLoopIdx = irrelevantHeaderLoopIdx + 1
    Next irrelevantHeader
End Sub

Sub OverwriteOriginalSheet(originalSheet As Worksheet, tmpSheet As Worksheet)
    Dim lastColumn As Long
    Dim tempLastColumn As Long

    ' 一時的なシートと元のシートの最終列を取得
    lastColumn = originalSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    tempLastColumn = tmpSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    ' オリジナルのシートの列数が一時的なシートの列数より多い場合、余分な列を削除
    If lastColumn > tempLastColumn Then
        originalSheet.Columns((tempLastColumn + 1) & ":" & lastColumn).Delete
    End If

    ' 一時的なシートから元のシートに内容をコピー
    tmpSheet.Range(tmpSheet.Cells(1, 1), tmpSheet.Cells(tmpSheet.Rows.Count, tempLastColumn)).Copy
    originalSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    
    Application.CutCopyMode = False
End Sub

Sub Cleanup()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Sheets(TMP_SHEET_NAME)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

' 配列での検索を行う補助関数
Private Function FindInArray(arr As Variant, search As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = search Then
            FindInArray = i
            Exit Function
        End If
    Next i
    FindInArray = -1 ' 見つからなかった場合
End Function

' 配列内に指定された値が存在するかどうかをチェックする補助関数
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Private Function GetCurrentSecondaryHeaders(sheet As Worksheet, Optional ByVal startColumn As Long = 1, Optional ByVal endColumn As Variant) As Variant
    Dim headerRange As Range
    Dim headers As Variant
    Dim lastColumn As Long

    ' endColumnが指定されていない場合、シートの最終列を使用
    If IsMissing(endColumn) Then
        lastColumn = sheet.Cells(SECONDARY_CATEGORY_ROW_NUM, sheet.Columns.Count).End(xlToLeft).Column
    Else
        lastColumn = endColumn
    End If

    ' 指定した範囲のヘッダー行から列範囲を取得
    Set headerRange = sheet.Range(sheet.Cells(SECONDARY_CATEGORY_ROW_NUM, startColumn), sheet.Cells(SECONDARY_CATEGORY_ROW_NUM, lastColumn))

    ' 範囲の値を配列として取得
    headers = headerRange.Value

    ' 配列の1行目（この場合は唯一の行）を返す
    ' VBAでは、Range.Valueが1行のみの場合でも2次元配列を返すため、明示的に1行目を指定
    GetCurrentSecondaryHeaders = Application.Index(headers, 1, 0)
End Function
