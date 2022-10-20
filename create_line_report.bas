Attribute VB_Name = "Module2"
Sub main()
    create_sheet
    write_report
End Sub

Sub create_sheet()
    'import Yahoo Finance CSV
    'C:\Users\mm_in\Downloads\quotes.csv
    'D:\Users\author\Download\quotes.csv
    ActiveWorkbook.Queries.Add Name:="yahoof", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""D:\Users\author\Download\quotes.csv""),[Delimiter="","", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""Symbol"", type text}, {""Current Price"", ty" & _
        "pe number}, {""Date"", type date}, {""Time"", type text}, {""Change"", type number}, {""Open"", type number}, {""High"", type number}, {""Low"", type number}, {""Volume"", Int64.Type}, {""Trade Date"", type text}, {""Purchase Price"", type text}, {""Quantity"", type text}, {""Commission"", type text}, {""High Limit"", type text}, {""Low Limit"", type text}, {""Comme" & _
        "nt"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=yahoof;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [yahoof]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "yahoof"
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.Name = "YahooFinance"
    
    'import investing SOX
    ActiveWorkbook.Queries.Add Name:="SOX30", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Web.Page(Web.Contents(""https://jp.investing.com/indices/phlx-semiconductor-components""))," & Chr(13) & "" & Chr(10) & "    Data0 = Source{0}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data0,{{"""", type text}, {""名前"", type text}, {""現在値"", type number}, {""高値"", type number}, {""安値"", type number}, {""前日比"", type number}, {""前日比%"", Percentage.Type}, {""出来高" & _
        """, type text}, {""時間"", type text}, {""2"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""SOX30"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [SOX30]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "sox"
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.Name = "SOX30"
    
    'import US2
    ActiveWorkbook.Queries.Add Name:="US2Y", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Web.Page(Web.Contents(""https://jp.investing.com/rates-bonds/u.s.-2-year-bond-yield""))," & Chr(13) & "" & Chr(10) & "    Data6 = Source{6}[Data]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Data6,{{"""", type text}, {""名前"", type text}, {""価格"", type number}, {""前日比"", type number}, {""変動%"", Percentage.Type}, {""2"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""US2Y"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [US2Y]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "us2y"
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.Name = "US2Y"
End Sub

Sub write_report()
    Dim i As Integer
    Dim last_row
    Dim items() As Variant
    Dim report As String
    ReDim stock_pairs(50, 2)
    Dim stock_idx As String
    Dim stock_name As String
    Dim path As String
    Dim report_for_td As String
    Dim tv_stock_ary() As Variant
    tv_stock_ary = Array("TVC-MOVE", "DJ-REIT", "INDEX-BDI")
    start_driver
    'D:\Users\author\OneDrive\Stock Workspace\Workspace\StockPair.csv
    path = "D:\Users\author\OneDrive\Stock Workspace\Workspace\StockPair.csv"
    stock_pairs = loadCSV(path)
    Worksheets("YahooFinance").Activate
    Columns("J:P").ClearContents
    
    Dim close_price As Double
    Dim last_price As Double
    For Each idx In tv_stock_ary
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        Call fetch_tradingview_data(close_price, last_price, idx)
        delta_price = close_price - last_price
        Cells(last_row + 1, 1).Value = idx
        Cells(last_row + 1, 2).Value = close_price
        Cells(last_row + 1, 5).Value = delta_price
    Next idx
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    'make change% column
    Range(Cells(2, 10), Cells(last_row, 10)).Formula2 = "=E2 / (B2 - E2) * 100"
    
    'make sentence
    For i = 2 To get_last_row()
        If Cells(i, 1) Like "*=X" Or Cells(i, 1) = "^TNX" Then
            items = Array(Cells(i, 1) & ":", Round(Cells(i, 2), 3), getFormattedItem(Cells(i, 5), 3), getFormattedItem(Cells(i, 10), 3) & "%, ")
        Else
            items = Array(Cells(i, 1) & ":", Round(Cells(i, 2), 2), getFormattedItem(Cells(i, 5), 2), getFormattedItem(Cells(i, 10), 2) & "%, ")
        End If
        Cells(i, 11) = Join(items)
    Next i
    
    'add US2Y
    us2y_ary = Array("2年債金利: ", Worksheets("US2Y").Cells(2, 3), "% ", getFormattedItem(Worksheets("US2Y").Cells(2, 4), 3), " (", getFormattedItem(Worksheets("US2Y").Cells(2, 5) * 100, 3), "%)")
    Worksheets("YahooFinance").Activate
    'us10y_row = Range("A1:K200").Find("^TNX", LookAt:=xlPart).Row
    'Rows(us10y_row + 1).Insert
    last_row = get_last_row()
    Cells(last_row + 1, 11) = Join(us2y_ary)
    Cells(last_row + 1, 1) = "US2Y"
    
    Dim stock_sort_ary As Variant
    stock_sort_ary = Array(Array("^TNX", "US2Y", 1), Array("^TNX", "TVC-MOVE", 0))
    For i = LBound(stock_sort_ary) To UBound(stock_sort_ary)
        before_row = Range("A1:K200").Find(stock_sort_ary(i)(0), LookAt:=xlPart).Row
        after_row = Range("A1:K200").Find(stock_sort_ary(i)(1), LookAt:=xlPart).Row
        Rows(after_row).Cut
        Rows(before_row + stock_sort_ary(i)(2)).Insert
    Next
    
    'replace index with stock name
    For i = LBound(stock_pairs, 1) To UBound(stock_pairs, 1)
        stock_idx = stock_pairs(i, 0)
        stock_name = stock_pairs(i, 1)
        Range(Cells(2, 11), Cells(get_last_row(), 11)).Replace stock_idx, stock_name, LookAt:=xlPart
    Next i
    
    'connect sentence
    For i = 2 To get_last_row()
        report = report & Cells(i, 11)
    Next i
    Cells(get_last_row() + 1, 11).Value = report
    
    'count rising index
    'Worksheets("SOX30").Activate
    Dim rise_idx As Variant
    With WorksheetFunction
        rise_idx = .CountIf(Worksheets("SOX30").Range("F2:F31"), ">0")
    End With
    Worksheets("YahooFinance").Cells(get_last_row() + 1, 11).Value = "SOXの上昇銘柄数: " & rise_idx
    
    last_row = get_last_row()
    Dim pred_data As Variant
    pred_data = fetch_cme_data()
    For x = LBound(pred_data, 1) + 1 To 4
        For y = LBound(pred_data, 2) To UBound(pred_data, 2)
            Worksheets("YahooFinance").Cells(x + last_row, y + 10).Value = pred_data(x, y)
            If x >= 3 And y = 1 Then
                Worksheets("YahooFinance").Cells(x + last_row, y + 10).NumberFormatLocal = "yyyy/mm/dd"
            End If
        Next
    Next
    
    close_driver
End Sub

Function getFormattedItem(item, round_num) As Variant
    Dim formatted As String
    formatted = Round(item, round_num)
    If formatted > 0 Then
        formatted = "+" & formatted
    End If
    getFormattedItem = formatted
End Function

Function get_last_row() As Long
    get_last_row = Range("A1:K1").EntireColumn.Find("*", , , , 1, 2).Row
    'get_last_row = Cells(Rows.Count, 1).End(xlUp).Row
End Function

'https://ateitexe.com/excel-vba-csv-to-multi-dimensional-array/
Function loadCSV(ByVal path As String) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  
  file = path
  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
  ReDim ary(max_n - 1, 1) As Variant     '取得した行数で2次元配列の再定義
  
  'CSVファイルを配列へ
  Open file For Input As #1 'CSVファイルを開く
  Do Until EOF(1) '最終行までループ
    Line Input #1, buf '読み込んだデータを1行ずつみていく
    tmp = Split(buf, ",") 'カンマで分割
    For i = 0 To UBound(tmp) '項目数ぶんループ
      ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
    Next i
    n = n + 1 '配列の次の行へ
  Loop
  Close #1 'CSVファイルを閉じる
  
  loadCSV = ary
End Function


