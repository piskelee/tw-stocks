<%@ codepage=65001%>
<!--#include file="lib.inc"-->


<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" charset="UTF-8">
    <link href="style.css" rel="stylesheet" type="text/css" />
    <link rel="preconnect" href="https://fonts.gstatic.com">
    <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
    <title>股票清單</title>
</head>
<body>
<div class="wrap">
    <a href="index.asp">股票清單</a>
    <a href="index2.asp">股票線型</a>

    <a href="index.html">股票清單</a>
    <a href="index2.html">股票線型</a>

<H1>股票清單</H1>
<%
    sql = " SELECT T.STOCK_ID,T.STOCK_NAME,T.STOCK_DESC,T.NOW_PRICES,T.KD,T.RSI,T.MACD,T.BAND,T.OBV,T.LAST_UPDATE_TIME FROM ETF_LIST T order by t.STOCK_ID "
    set rs=Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    Response.Write "<table>"
    Response.Write "<thead>"
    Response.Write "<tr>"

    Response.Write "<th>"
    Response.Write "ITEM"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "ID"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "NAME"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "DESC"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "PRICES"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "KD"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "BAND"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "UPDATE_TIME"
    Response.Write "</th>"

    Response.Write "</tr>"
    Response.Write "</thead>"

    Response.Write "<tbody>"

    Dim itemNumber
    itemNumber = 1

    While Not rs.EOF
        Response.Write "<tr>"

        Response.Write "<td>"
        Response.Write itemNumber
        Response.Write "</td>"

        STOCK_ID = rs("STOCK_ID")
        STOCK_ID = Split(STOCK_ID, ".")(0) '
        Response.Write "<td>" & STOCK_ID & "</td>"

        Response.Write "<td>"
        Response.Write rs("STOCK_NAME")
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write rs("STOCK_DESC")
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write rs("NOW_PRICES")
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write rs("KD")
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write rs("BAND")
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write rs("LAST_UPDATE_TIME")
        Response.Write "</td>"

        Response.Write "</tr>"
        rs.MoveNext
        itemNumber = itemNumber + 1
    Wend
    Response.Write "</tbody>"
    Response.Write "</table>"
%>
</div>
</br>
</br>
</body>
</html>


<%
    ' 關閉記錄集和連接
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
%>