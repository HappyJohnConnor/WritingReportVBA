Attribute VB_Name = "Module1"
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
    'D:\Users\author\Document\Stock Workspace\Workspace\StockPair.csv
    path = "D:\Users\author\Document\Stock Workspace\Workspace\StockPair.csv"
    stock_pairs = loadCSV(path)
    Worksheets("YahooFinance").Activate
    Columns("J:P").ClearContents
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    'make change% column
    Range(Cells(2, 10), Cells(last_row, 10)).Formula2 = "=E2 / (B2 - E2) * 100"
    
    'make sentence
    For i = 2 To last_row
        items = Array(Cells(i, 1) & ":", Round(Cells(i, 2), 2), getFormattedItem(Cells(i, 5)), getFormattedItem(Cells(i, 10)) & "%, ")
        Cells(i, 11) = Join(items)
    Next i
    Cells(last_row + 1, 11) = get_reprot_from_TradingView()
    
    'replace index with stock name
    For i = LBound(stock_pairs, 1) To UBound(stock_pairs, 1)
        stock_idx = stock_pairs(i, 0)
        stock_name = stock_pairs(i, 1)
        Range(Cells(2, 11), Cells(last_row + 1, 11)).Replace stock_idx, stock_name, LookAt:=xlPart
    Next i
    
    'connect sentence
    For i = 2 To last_row + 1
        report = report & Cells(i, 11)
    Next i
    Cells(last_row + 3, 11).Value = report
    
    'count rising index
    Worksheets("SOX30").Activate
    Dim rise_idx As Variant
    With WorksheetFunction
        rise_idx = .CountIf(Range("F2:F31"), ">0")
    End With
    Worksheets("YahooFinance").Cells(last_row + 4, 11).Value = "SOXの上昇銘柄数: " & rise_idx
    
    'write
    Worksheets("US2Y").Activate
    item = Array("2年債金利: ", Cells(2, 3), "% ", getFormattedItem2(Cells(2, 4)), " (", getFormattedItem2(Cells(2, 5) * 100), "%)")
    Worksheets("YahooFinance").Cells(last_row + 5, 11).Value = Join(item)
    
End Sub

Function getFormattedItem(item) As Variant
    Dim formatted As String
    formatted = Round(item, 2)
    If formatted > 0 Then
        formatted = "+" & formatted
    End If
    getFormattedItem = formatted
End Function

Function getFormattedItem2(item) As Variant
    Dim formatted As String
    If item > 0 Then
        formatted = "+" & item
    End If
    getFormattedItem2 = formatted
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
