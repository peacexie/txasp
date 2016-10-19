<!--#include file="inc_perm.asp"-->
<%
rTim01 = Timer()

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Tab Data</title>
<link href="/inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Pragma" content="no-cache">
</head>
<body>
<%
SET rs=Server.CreateObject("Adodb.Recordset")
'SELECT * FROM
'(SELECT * FROM CUSTOMER ORDER BY NAME)
'WHERE ROWNUM <= 3
'ORDER BY ROWNUM ASC

'All:479480 CST_CUSTOMER :CUSTOMER_ID, NAME FROM
'All:1343302 PRD_PUBLISH_CONFIG:ID, PRICE2, CREATE_DATE 
RecA = 20 '20 '140123
RecB = 10


sql = "       SELECT * FROM SYS_STAFF   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 "
sql = sql& "  ORDER BY LOGIN DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("LOGIN")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"
  
sql = "       SELECT * FROM SYS_STAFF   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 AND length(MOBILEPHONE)>8 "
sql = sql& "  ORDER BY MOBILEPHONE DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("MOBILEPHONE")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"


sql = "       SELECT * FROM CST_CUSTOMER   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 "
sql = sql& "  ORDER BY LOGIN DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("LOGIN")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"
  
sql = "       SELECT * FROM CST_CUSTOMER   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 "
sql = sql& "  ORDER BY MOBILE DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("MOBILE")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"


rTim01 = Timer() - rTim01
Response.Write "<br><br> LOGIN DESC : "&RecA&" - "&rTim01&"<br><br>"


SET rs=Nothing
%>

</body>
</html>
