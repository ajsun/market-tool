Function YahooFinance(ticker As String, item As String)
    Dim strURL As String, strCSV As Double, itemFound As Integer, tag As String
 
    itemFound = 0
    If item = "ask" Then
        tag = "a"
        itemFound = 1
    ElseIf item = "bid" Then
        tag = "b"
        itemFound = 1
    ElseIf item = "52weeklow" Then
        tag = "j"
        itemFound = 1
    ElseIf item = "marketcap" Then
        tag = "j1"
        itemFound = 2
    ElseIf item = "open" Then
        tag = "o"
        itemFound = 1
    ElseIf item = "averagedailyvolume" Then
        tag = "a2"
        itemFound = 1
    ElseIf item = "askrealtime" Then
        tag = "b2"
        itemFound = 1
    ElseIf item = "bidsize" Then
        tag = "b6"
        itemFound = 1
    ElseIf item = "commision" Then
        tag = "c3"
        itemFound = 1
    ElseIf item = "dividendshare" Then
        tag = "d"
        itemFound = 1
    ElseIf item = "EPS" Then
        tag = "e"
        itemFound = 1
    ElseIf item = "epsestimatenextyear" Then
        tag = "e8"
        itemFound = 1
    ElseIf item = "52weekhigh" Then
        tag = "k"
        itemFound = 1
    ElseIf item = "holdsingain" Then
        tag = "g4"
        itemFound = 1
    ElseIf item = "marketcaprealtime" Then
        tag = "j3"
        itemFound = 1
    ElseIf item = "lasttradesize" Then
        tag = "k3"
        itemFound = 1
    ElseIf item = "lasttradewithtime" Then
        tag = "l"
        itemFound = 1
    ElseIf item = "name" Then
        tag = "n"
        itemFound = 2
    ElseIf item = "previousclose" Then
        tag = "p"
        itemFound = 1
    ElseIf item = "pricesales" Then
        tag = "p5"
        itemFound = 1
    ElseIf item = "peratio" Then
        tag = "r"
        itemFound = 1
    ElseIf item = "pegratio" Then
        tag = "r5"
        itemFound = 1
    ElseIf item = "symbol" Then
        tag = "s"
        itemFound = 1
    ElseIf item = "lasttradetime" Then
        tag = "t1"
        itemFound = 1
    ElseIf item = "1yeartargetprice" Then
        tag = "t8"
        itemFound = 1
    ElseIf item = "asksize" Then
        tag = "a5"
        itemFound = 1
    ElseIf item = "bidrealtime" Then
        tag = "b3"
        itemFound = 1
    ElseIf item = "ebitda" Then
        tag = "j4"
        itemFound = 2
    ElseIf item = "lasttraderealtimewithtime" Then
        tag = "k1"
        itemFound = 1
    ElseIf item = "pricebook" Then
        tag = "p6"
        itemFound = 1
    ElseIf item = "dividendpaydate" Then
        tag = "r1"
        itemFound = 1
    ElseIf item = "priceepsestimatecurrentyear" Then
        tag = "r6"
        itemFound = 1
    ElseIf item = "sharesowned" Then
        tag = "s1"
        itemFound = 1
    ElseIf item = "volume" Then
        tag = "v"
        itemFound = 1
    ElseIf item = "52weekrange" Then
        tag = "w"
        itemFound = 1
    ElseIf item = "stockexchange" Then
        tag = "x"
        itemFound = 1
    ElseIf item = "changefrom52weeklow" Then
        tag = "j5"
        itemFound = 1
    End If
 
    If itemFound = 1 Then
        strURL = "http://download.finance.yahoo.com/d/quotes.csv?s=" & ticker & "&f=" & tag
        Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
        XMLHTTP.Open "GET", strURL, False
        XMLHTTP.send
        YahooFinance = Val(XMLHTTP.responseText)
        Set XMLHTTP = Nothing
    ElseIf itemFound = 2 Then
        strURL = "http://download.finance.yahoo.com/d/quotes.csv?s=" & ticker & "&f=" & tag
        Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
        XMLHTTP.Open "GET", strURL, False
        XMLHTTP.send
        YahooFinance = XMLHTTP.responseText
        Set XMLHTTP = Nothing
    Else
        YahooFinance = "Item Not Found"
    End If
 
End Function
