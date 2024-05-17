<% 
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    Dim connStr
    connStr = "Provider=SQLOLEDB;Data Source=DESKTOP-EULMLFK;Initial Catalog=MY_ETF;User ID=sa;Password=t2ecws;Charset=UTF-8;"
    conn.Open connStr
%>