<!--#include file="inc_perm.asp"-->
<%
rTim01 = Timer()

'1.在ORACLE中实现SELECT TOP N 
'则按NAME的字母顺抽出前三个顾客的SQL语句如下所示：
'SELECT * FROM
'(SELECT * FROM CUSTOMER ORDER BY NAME)
'WHERE ROWNUM <= 3
'ORDER BY ROWNUM ASC

'2.在TOP N纪录中抽出第M（M <= N）条记录
'得到以NAME的字母顺排序的第二个顾客的信息的SQL语句应该这样写：
'SELECT ID, NAME FROM
'( SELECT ROWNUM RECNO, ID, NAME FROM
'    (SELECT * FROM CUSTOMER ORDER BY NAME)
'  WHERE ROWNUM <= 3
'  ORDER BY ROWNUM ASC 
')
'WHERE RECNO = 2

'3.抽出按某种方式排序的记录集中的第N条记录
' SELECT ID, NAME FROM
' ( SELECT ROWNUM RECNO, ID, NAME FROM
'     (SELECT * FROM CUSTOMER ORDER BY NAME)
'   WHERE ROWNUM <= 2
'   ORDER BY ROWNUM ASC
' )
' WHERE RECNO = 2

'4.抽出按某种方式排序的记录集中的第M条记录开始的X条记录
'则抽取NAME的字母顺的第2条记录开始的3条记录的SQL语句为：
' SELECT ID, NAME FROM
' (  SELECT ROWNUM RECNO, ID, NAME FROM
'      (SELECT * FROM CUSTOMER ORDER BY NAME)
'    WHERE ROWNUM <= (2 + 3 - 1)
'    ORDER BY ROWNUM ASC 
' )
' WHERE RECNO BETWEEN 2 AND (2 + 3 - 1)


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


'All:2,023,420 PRD_PUBLISH_CONFIG:ID, PRICE2, CREATE_DATE 
RecA = 12*1024 '20 '1343302
RecB = 10

sql = ""
sql = "       SELECT RECNO, ID, PRICE2, CREATE_DATE FROM "
sql = sql& "  ( "
sql = sql& "    SELECT ROWNUM RECNO, ID, PRICE2, CREATE_DATE FROM " 
sql = sql& "      ( SELECT ID, PRICE2, CREATE_DATE FROM PRD_PUBLISH_CONFIG "
'sql = sql& "        WHERE PRICE2>200 AND TO_CHAR (CREATE_DATE, 'YYYY-MM-DD')>'2007-06-30' "
sql = sql& "        ORDER BY ID DESC "
sql = sql& "      ) "
sql = sql& "    WHERE ROWNUM <= ("&RecA&" + "&RecB&" - 1) "  ' " '初始的 最优先SQL语句
sql = sql& "    ORDER BY ROWNUM ASC "
sql = sql& "  ) "
sql = sql& "  WHERE RECNO BETWEEN "&RecA&" AND ("&RecA&" + "&RecB&" - 1) "

' OK 
' sql = "SELECT * FROM (SELECT * FROM PRD_PUBLISH_CONFIG ORDER BY ID DESC) WHERE ROWNUM <= 3 ORDER BY ROWNUM ASC "


Response.Write sql

  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open sql,conn,3,1
  rs.MoveFirst()
  'rs.MoveNext() 
  Do While NOT rs.Eof
    Response.Write "<br> ------ "&" "&rs("RECNO")&" "&rs("ID")&" - "&rs("PRICE2")&" - "&rs("CREATE_DATE") ' 
  rs.MoveNext() 
  Loop
  rs.Close
  SET rs=Nothing

rTim01 = Timer() - rTim01
Response.Write "<br><br> 1343302 : "&RecA&" - "&rTim01&"<br><br>"



%>

</body>
</html>
