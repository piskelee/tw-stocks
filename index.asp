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
    <h1>股票清單</h1>
    <%
        sql = " SELECT T.STOCK_ID,T.STOCK_NAME,T.STOCK_DESC,T.NOW_PRICES,T.KD,T.RSI,T.MACD,T.BAND,T.OBV,T.LAST_UPDATE_TIME FROM ETF_LIST T order by t.STOCK_ID "
        set rs=Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn

        Response.Write "<div class='table-responsive'>"
        Response.Write "<table class='table'>"
        Response.Write "<thead>"
        Response.Write "<tr>"

        Response.Write "<th>"
        Response.Write "ID"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "名稱"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "價格"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "KD"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "BAND"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "MACD"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "OBV"
        Response.Write "</th>"

        Response.Write "<th>"
        Response.Write "更新時間"
        Response.Write "</th>"

        Response.Write "</tr>"
        Response.Write "</thead>"
        Response.Write "<tbody>"

        Dim itemNumber
        itemNumber = 1

        While Not rs.EOF
            STOCK_ID = rs("STOCK_ID")
            STOCK_ID = Split(STOCK_ID, ".")(0) '
            Response.Write "<td>" & STOCK_ID & "</td>"

            Response.Write "<td>"
            Response.Write "<a href='" & STOCK_ID & ".html'>" & rs("STOCK_NAME") & "</a>"
            Response.Write "</td>"

            Response.Write "<td>"
            Response.Write rs("NOW_PRICES")
            Response.Write "</td>"

            Response.Write "<td>"
            If rs("KD") = "Y" Then
                Response.Write "<span style='color:red; font-weight:bold;'>" & rs("KD") & "</span>"
            Else
                Response.Write rs("KD")
            End If
            Response.Write "</td>"

            Response.Write "<td>"
            If rs("BAND") = "Y" Then
                Response.Write "<span style='color:red; font-weight:bold;'>" & rs("BAND") & "</span>"
            Else
                Response.Write rs("BAND")
            End If
            Response.Write "</td>"

	        Response.Write "<td>"
            If rs("MACD") = "Y" Then
                Response.Write "<span style='color:red; font-weight:bold;'>" & rs("MACD") & "</span>"
            Else
                Response.Write rs("MACD")
            End If
            Response.Write "</td>"
    
            Response.Write "<td>"
            If rs("OBV") = "Y" Then
                Response.Write "<span style='color:red; font-weight:bold;'>" & rs("OBV") & "</span>"
            Else
                Response.Write rs("OBV")
            End If
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