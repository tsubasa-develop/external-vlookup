Sub VlookupFromOtherWorkbook()
    Dim currentWorkbook As Workbook
    Dim otherWorkbook As Workbook
    Dim otherSheet As Worksheet
    Dim resultSheet As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim lookupValue As String
    Dim lookupRange As Range
    Dim result As Variant
    Dim searchColumn As Range

    Set currentWorkbook = ThisWorkbook
    Set currentSheet = currentWorkbook.Sheets("search")

    ' C3セルからファイル名を取得
    filePath = currentWorkbook.Path & "/" & currentSheet.Range("C3").Value

    ' C4セルからシート名を取得
    sheetName = currentSheet.Range("C4").Value

    ' C5セルからVLOOKUPの検索キーを取得
    lookupValue = currentSheet.Range("C5").Value
    lookupIndex = currentSheet.Range("C6").Value

    ' 他のブックを開く
    On Error GoTo ErrorHandler ' エラーハンドリングの設定
    Set otherWorkbook = Workbooks.Open(filePath)
    On Error GoTo 0 ' エラーハンドリングを解除

    ' シートを指定
    Set otherSheet = otherWorkbook.Sheets(sheetName)

    ' 検索範囲（A列が検索キー、B列が対応する値と仮定）
    Set searchColumn = otherSheet.Range("A:Z")

    ' 検索キーを使ってVLOOKUPのような検索を実行
    result = Application.WorksheetFunction.VLookup(lookupValue, searchColumn, lookupIndex, False)

    ' 結果を出力
    currentSheet.Cells(3, 5).Value = result

    ' 他のブックを閉じる（保存しない）
    otherWorkbook.Close SaveChanges:=False

    Exit Sub

ErrorHandler:
    MsgBox "ファイルを開けませんでした。ファイル名とシート名を確認してください。"
End Sub
