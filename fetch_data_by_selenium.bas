Attribute VB_Name = "Module3"
'Functions which use Webdriver directly

Dim driver As New Selenium.ChromeDriver

Function start_driver()
    driver.Start
End Function


Function close_driver()
 driver.Close
End Function

Function fetch_cme_data() As Variant
    Dim tbody_element As WebElement
    Dim btn_element As WebElement
    Dim trs As WebElements
    Dim tds As WebElements
    Dim ths As WebElements
    Dim myBy As New By
    Dim pred_data(2, 1) As Variant
    Dim pred_talbe As Variant
    url_string = "https://www.cmegroup.com/ja/trading/interest-rates/countdown-to-fomc.html"
    'driver.Start
    driver.Get url_string
    'Application.Wait [Now() + "00:00:2"]
    
    With driver
        .SwitchToFrame .FindElementByTag("iframe")
        'Application.Wait [Now() + "00:00:5"]
        Dim i As Long
        i = 0
        Do Until i > 5
            If .IsElementPresent(myBy.linktext("Probabilities")) Then
                Exit Do
            End If
            .Wait 1000
            i = i + 1
        Loop
        Set btn_element = .FindElementByLinkText("Probabilities")
        hoge = .ExecuteScript("arguments[0].click();", btn_element)
        '.FindElementByLinkText("Probabilities").Click
        
        tbl_xpath = "//*[@id=""MainContent_pnlContainer""]/div[3]/div/div/table[2]"
        pred_table = .FindElementByXPath(tbl_xpath).AsTable.Data
        
    End With
    fetch_cme_data = pred_table
End Function

Function fetch_tradingview_data(ByRef close_price As Double, ByRef last_price As Double, idx As Variant)
    Dim element As WebElement
    Dim myBy As New By
    url_string = "https://jp.tradingview.com/symbols/" & idx
    driver.Get url_string
    Application.Wait [Now() + "00:00:00.30"]
    With driver
        Call wait_unti_appear("tv-category-header__price-line")
        Set element = .FindElementByClass("tv-category-header__price-line")
        Call wait_unti_appear("tv-symbol-price-quote__value")
        close_price = CDbl(element.FindElementByClass("tv-symbol-price-quote__value").Text)
        last_price = CDbl(element.FindElementByClass("js-header-fundamentals").FindElementByClass("js-symbol-prev-close").Text)
    End With
End Function

Function wait_unti_appear(class_name As String)
    Dim myBy As New By
    With driver
        Dim i As Long
        i = 0
        Do Until i > 5
            If .IsElementPresent(myBy.Class(class_name)) Then
                Exit Do
            End If
            .Wait 1000
            i = i + 1
        Loop
    End With
End Function
