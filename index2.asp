<%@ codepage=65001%>
<!--#include file="lib.inc"-->



<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" charset="UTF-8">
    <link href="style.css" rel="stylesheet" type="text/css" />
    <link rel="preconnect" href="https://fonts.gstatic.com">
    <link href="https://fonts.googleapis.com/css2?family=Oswald&display=swap" rel="stylesheet">
    <title>股票線型</title>
    <style>
        .popup-container {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
        }
        .popup-image {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            max-width: 90%;
            max-height: 90%;
        }
    </style>
</head>
<body>
<div class="wrap">
    <a href="index.asp">股票清單</a>
    <a href="index2.asp">股票線型</a>

<H1>股票線型</H1>
<%
    sql = " SELECT t.STOCK_ID,t.STOCK_NAME from ETF_LIST t order by t.STOCK_ID "
    set rs=Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    Response.Write "<table>"
    Response.Write "<thead>"
    Response.Write "<tr>"

    Response.Write "<th>"
    Response.Write "項次"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "ID"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "NAME"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "KD"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write "BAND"
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write ""
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write ""
    Response.Write "</th>"

    Response.Write "<th>"
    Response.Write ""
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
        Response.Write "<a href=""javascript:void(0)"" onclick=""showPopup('" & STOCK_ID & "_kd.png')"">"
        Response.Write "<img src=""" & STOCK_ID & "_kd.png"" width=""360"" height=""250"">"
        Response.Write "</a>"
        Response.Write "</td>"

        Response.Write "<td>"
        Response.Write "<a href=""javascript:void(0)"" onclick=""showPopup('" & STOCK_ID & "_band.png')"">"
        Response.Write "<img src=""" & STOCK_ID & "_band.png"" width=""360"" height=""250"">"
        Response.Write "</a>"
        Response.Write "</td>"

        Response.Write "</tr>"
        rs.MoveNext
        itemNumber = itemNumber + 1
    Wend
    Response.Write "</tbody>"
    Response.Write "</table>"
%>

<div class="popup-container" id="popupContainer" onclick="hidePopup()">
    <img class="popup-image" id="popupImage" src="" alt="Popup Image">
</div>

<script>
    function showPopup(imagePath) {
        const popupContainer = document.getElementById("popupContainer");
        const popupImage = document.getElementById("popupImage");
        popupImage.src = imagePath;
        popupContainer.style.display = "block";
    }

    function hidePopup() {
        const popupContainer = document.getElementById("popupContainer");
        popupContainer.style.display = "none";
    }
</script>

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