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

Function rs_DoSql(xconn,xSql) 
   Dim cnFPubcnDo 
   Set cnDo = Server.CreateObject("Adodb.Connection")
   cnDo.Open xconn 
   cnDo.Execute(xSql)
   cnDo.Close()
   Set cnDo = Nothing
End Function

If Request("yAct")<>"" AND Request("ID")<>"" Then
  Call rs_DoSql(conn,"DELETE FROM SYS_STAFF WHERE LOGIN='"&Request("ID")&"'")
  Response.Write "Del OK!"
End If

SET rs=Server.CreateObject("Adodb.Recordset")


Response.Write "<br>1.LOGIN LIKE 'demo%'"
sql = "       SELECT * FROM CST_CUSTOMER   " 'where LOGIN>'a'  
sql = sql& "  WHERE ROWNUM <= 10 AND LOGIN LIKE 'demo%' "
sql = sql& "  ORDER BY LOGIN DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("STAFF_ID")&"] "&" ["&rs("LOGIN")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"


Response.Write "<br>MOBILEPHONE LIKE '17%'"  
sql = "       SELECT * FROM CST_CUSTOMER   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 AND MOBILE LIKE '17%' "
sql = sql& "  ORDER BY MOBILE DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("STAFF_ID")&"] "&" ["&rs("MOBILE")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"



Response.Write "<br>1.LOGIN LIKE 'demo%'"
sql = "       SELECT * FROM SYS_STAFF   " 'where LOGIN>'a'  
sql = sql& "  WHERE ROWNUM <= 10 AND LOGIN LIKE 'demo%' "
sql = sql& "  ORDER BY LOGIN DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("STAFF_ID")&"] "&" ["&rs("LOGIN")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"


Response.Write "<br>MOBILEPHONE LIKE '17%'"  
sql = "       SELECT * FROM SYS_STAFF   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 10 AND MOBILEPHONE LIKE '17%' "
sql = sql& "  ORDER BY MOBILEPHONE DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" ["&rs("STAFF_ID")&"] "&" ["&rs("MOBILEPHONE")&"] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"



Response.Write "<br>1.LOGIN LIKE 'aa%'"
sql = "       SELECT * FROM SYS_STAFF   " 'where LOGIN>'a' 
sql = sql& "  WHERE ROWNUM <= 100 AND LOGIN LIKE 'aaa%' "
sql = sql& "  ORDER BY LOGIN DESC "
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br> ------ "&" [<a href='?yAct=Del&ID="&rs("LOGIN")&"'>"&rs("LOGIN")&"</a>] "&rs("CUSTOMER_ID")&" - "&rs("NAME")
  rs.MoveNext()
Loop
rs.Close
Response.Write "<br>"


rTim01 = Timer() - rTim01
Response.Write "<br> LOGIN DESC : "&RecA&" - "&rTim01&""


SET rs=Nothing
%>

</body>
</html>
