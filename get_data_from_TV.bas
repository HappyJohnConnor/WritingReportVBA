Attribute VB_Name = "Module5"

Dim driver As New Selenium.ChromeDriver
Function get_reprot_from_TradingView() As String
    Dim close_price As Double
    Dim last_price As Double
    Dim idx_array() As Variant
    Dim words_array() As Variant
    Dim report As String
    idx_array = Array("TVC-MOVE", "DJ-REIT", "INDEX-BDI")
    'driver.AddArgument "headless"
    driver.Start
    
    For Each idx In idx_array
        Call getStockPriceFromTV(close_price, last_price, idx)
        delta_price = close_price - last_price
        change_per = delta_price / last_price * 100
        words_array = Array(idx, ":", Round(close_price, 2), getFormattedItem(delta_price), getFormattedItem(change_per) & "%, ")
        report = report & Join(words_array)
    Next idx
    driver.Close
    Debug.Print report
    get_reprot_from_TradingView = report
    
End Function

Function getStockPriceFromTV(ByRef close_price As Double, ByRef last_price As Double, idx As Variant)
    Dim element As WebElement
    url_string = "https://jp.tradingview.com/symbols/" & idx
    driver.Get url_string
    Application.Wait [Now() + "00:00:00.30"]
    With driver
        Set element = .FindElementByClass("tv-category-header__price-line")
        close_price = CDbl(element.FindElementByClass("tv-symbol-price-quote__value").Text)
        last_price = CDbl(element.FindElementByClass("js-header-fundamentals").FindElementByClass("js-symbol-prev-close").Text)
    End With
End Function

Function getFormattedItem(item) As Variant
    Dim formatted As String
    formatted = Round(item, 2)
    If formatted > 0 Then
        formatted = "+" & formatted
    End If
    getFormattedItem = formatted
End Function
