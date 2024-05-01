<%@ codepage=65001%>
<!--#include file="lib.inc"-->


<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta charset="UTF-8">
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
    <link href="style.css" rel="stylesheet" type="text/css" />
    <link rel="preconnect" href="https://fonts.gstatic.com">
    <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
    <title>股票清單</title>
</head>
<body>
<div class="wrap">
    <a href="index.asp">股票清單</a>
    <a href="index3.asp">股票線型</a>

    <a href="index.html">股票清單</a>
    <a href="index2.html">股票線型</a>

    <h1>股票清單</h1>
    <%
        sql = " SELECT t.STOCK_ID,t.STOCK_NAME from ETF_LIST t order by t.STOCK_ID "
        set rs=Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn

        Response.Write "<div class='table-responsive'>"
        Response.Write "<table class='table'>"
        Response.Write "<thead>"
        Response.Write "<tr>"

        Response.Write "<th>"
        Response.Write "項目"
        Response.Write "</th>"

        Response.Write "</tr>"
        Response.Write "</thead>"
        Response.Write "<tbody>"

        Dim itemNumber
        itemNumber = 1

        While Not rs.EOF
            STOCK_ID = rs("STOCK_ID")
            STOCK_ID = Split(STOCK_ID, ".")(0) '

            Response.Write "<tr>"
            Response.Write "<td>"
            Response.Write "<img src=""" & STOCK_ID & "_kd.png"" class=""img-fluid"" width=""360"" height=""250"">"
            Response.Write "<img src=""" & STOCK_ID & "_band.png"" class=""img-fluid"" width=""360"" height=""250"">"
            Response.Write "</td>"
            Response.Write "</tr>"
            rs.MoveNext
            itemNumber = itemNumber + 1
        Wend
        Response.Write "</tbody>"
        Response.Write "</table>"
        Response.Write "</div>"
        ' 關閉記錄集和連接
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
    %>
</div>
<br>
<br>
</body>
</html>