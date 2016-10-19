
<!--#include file="inc_perm.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test Products ... </title>
</head>

<body>
<pre style="color:#F00">
？？？ NowDT = 当天 2009-11-xx 无资料，证明数据库资料不是最近资料，这样个调试带来麻烦?!
TypID  = 150 
NowDT = "2009-05-15"
typeid = 142 
nowdate = "2009-05-15"
</pre>
<%

TypID  = 142 
NowDT = "2009-05-15"

'PRD_PRODUCT a
sqlca =" a.PRODUCT_ID as PID "
sqlca =sqlca&" ,a.NAME as PName "
sqlca =sqlca&" ,a.PRODUCT_TYPE_ID as PType "
sqlca =sqlca&" ,a.DAYSCONSUMING as PDays "
sqlca =sqlca&" ,a.DESCRIPTION as PRem "
sqlca =sqlca&" ,a.NOTES as PNotes "
sqlca =sqlca&" ,a.STATE as PState "
sqlca =sqlca&" ,a.PRODUCT_NO as PNO "

'PRD_COMBIN c
sqlcc =" c.PRODUCT_COM_ID as CID "
sqlcc =sqlcc&" ,c.NAME as CName "
sqlcc =sqlcc&" ,c.PRODUCTCODE as CCode "
sqlcc =sqlcc&" ,c.TOTAL_QUANTITY as QTotal "
sqlcc =sqlcc&" ,c.GROUP_QUANTITY as QGroup "
sqlcc =sqlcc&" ,c.ADULT_QUANTITY as QAdult "
sqlcc =sqlcc&" ,c.CHILD_QUANTITY as QChild "
sqlcc =sqlcc&" ,c.BABY_QUANTITY as QBaby "
sqlcc =sqlcc&" ,c.ADULT_PRICE as PAdult "
sqlcc =sqlcc&" ,c.CHILD_PRICE as PChild "
sqlcc =sqlcc&" ,c.BABY_PRICE as PBaby "
sqlcc =sqlcc&" ,c.MONEYTYPE as CMType "
sqlcc =sqlcc&" ,c.START_DATE as CStart "
sqlcc =sqlcc&" ,c.END_DATE as CEnd "


 
sql = "select "&sqlcc&" "
sql =sql&" from PRD_COMBIN c "
sql =sql&" where TO_CHAR (c.EFF_DATE, 'YYYY-MM-DD')<='"&NowDT&"' "
sql =sql&"        and TO_CHAR(c.EXP_DATE,'YYYY-MM-DD')>='"&NowDT&"' "
sql =sql&"        and TO_CHAR(c.End_Date,'YYYY-MM-DD')>='"&NowDT&"' "
sql =sql&"        and c.state=1 and c.customertypeid=1 "
sql =sql&"        and c.PRODUCT_ID=24705 " 
Response.Write "<br><b>d.PRD_COMBIN.深圳欢乐谷、玛雅海滩一天</b><br>"&vbcrlf&"<!--"&sql&"-->"
Response.Write "- [CID - ---CName--- - ---CCode---] - [QTotal - QGroup - QAdult - QChild] - [PAdult - PChild - PBaby] - [CMType - CStart - CEnd]"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " [ "&rs("CID")
  Response.Write " - "&rs("CName")
  Response.Write " - "&rs("CCode")
  Response.Write " ] "
  
  Response.Write " [ "&rs("QTotal")
  Response.Write " - "&rs("QGroup")
  Response.Write " - "&rs("QAdult")
  Response.Write " - "&rs("QChild")
  Response.Write " ] "
  
  Response.Write " [ "&rs("PAdult")
  Response.Write " - "&rs("PChild")
  Response.Write " - "&rs("PBaby")
  Response.Write " ] "
  
  Response.Write " [ "&rs("CMType")
  Response.Write " - "&rs("CStart")
  Response.Write " - "&rs("CEnd")
  Response.Write " ] "
  
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 "&Now()&"<br><br>"



sql = "select "&sqlca&" ,b.PAdult "
sql =sql&" from prd_product a "
sql =sql&" inner join ( select max(product_id) as ProID,min(adult_price) as PAdult,customertypeid as CCusTP "
sql =sql&"              from PRD_COMBIN c "
sql =sql&"              where TO_CHAR(c.EFF_DATE, 'YYYY-MM-DD')<='"&NowDT&"' and TO_CHAR(c.EXP_DATE,'YYYY-MM-DD')>='"&NowDT&"' "
sql =sql&"              and TO_CHAR(c.End_Date,'YYYY-MM-DD')>='"&NowDT&"' and c.state=1 group by c.product_id,c.customertypeid "
sql =sql&" ) b on(a.product_id=b.ProID) "
sql =sql&" where b.CCusTP=1 and a.product_type_id=150 and a.owner_entity_id=1 " 
'sql =sql&" where b.CCusTP=1 and a.product_type_id IN ( "
'sql =sql&"        select product_type_id from domain_product_type "
'sql =sql&"        where parent_product_type ="&TypID&" )"
'sql =sql&" and a.owner_entity_id=1 " 
sql =sql&" ORDER BY a.product_type_id " 
Response.Write "<br><b>0.prd_product.深圳(珠三角)150</b><br>"&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " - "&rs("PID")
  Response.Write " - "&rs("PName")
  Response.Write " - "&rs("PType")
  Response.Write " - "&rs("PAdult")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 "&Now()&"<br><br>"



sql = "select product_type_id,parent_product_type,name from domain_product_type "
sql = sql& " WHERE "
sql = sql& " product_type_id=142 or parent_product_type=142 "  
sql = sql& " ORDER BY parent_product_type DESC,product_type_id DESC " 
Response.Write "<br><b>0.product_type_id=142</b><br>"&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"&rs("product_type_id")&" - "&rs("parent_product_type")&" - "&rs("name")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br><b>以上资料</b>End product_type_id=142 "&Now()&"<br><br>"



sql2 = "select product_type_id,parent_product_type,name from domain_product_type "
sql2 = sql2& " WHERE "
'sql2 = sql2& " product_type_id=142 or parent_product_type=142 "  
'sql2 = sql2& " product_type_id=153 or product_type_id=149 "  
sql2 = sql2& " product_type_id IN(149,153) "  
sql2 = sql2& " ORDER BY parent_product_type DESC,product_type_id DESC " 
Response.Write "<br><b>0.product_type_id=142</b><br>"&vbcrlf&"<!--"&sql&"-->"
SET rs2=Server.CreateObject("Adodb.Recordset") 
rs2.Open sql2,conn,1,1
Do While NOT rs2.Eof  
  Response.Write "<br>"&rs2("product_type_id")&" - "&rs2("parent_product_type")&" - "&rs2("name")
  
sql = "select "&sqlca&" ,b.PAdult "
sql =sql&" from prd_product a "
sql =sql&" inner join ( select max(product_id) as ProID,min(adult_price) as PAdult,customertypeid as CCusTP "
sql =sql&"              from PRD_COMBIN c "
sql =sql&"              where TO_CHAR(c.EFF_DATE, 'YYYY-MM-DD')<='"&NowDT&"' and TO_CHAR(c.EXP_DATE,'YYYY-MM-DD')>='"&NowDT&"' "
sql =sql&"              and TO_CHAR(c.End_Date,'YYYY-MM-DD')>='"&NowDT&"' and c.state=1 group by c.product_id,c.customertypeid "
sql =sql&" ) b on(a.product_id=b.ProID) "
sql =sql&" where b.CCusTP=1 and a.product_type_id="&rs2("product_type_id")&" and a.owner_entity_id=1 "  
sql =sql&" ORDER BY a.product_type_id " 
'Response.Write "<br><b>0.prd_product.深圳(珠三角)150</b><br>"&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " ---------- "&rs("PID")
  Response.Write " - "&rs("PName")
  Response.Write " - "&rs("PType")
  Response.Write " - "&rs("PAdult")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 IN(149,153) "&Now()&"<br><br>"
  
rs2.MoveNext()
Loop
rs2.Close
SET rs2=Nothing
Response.Write "<br><br><b>以上资料</b>End product_type_id=142 "&Now()&"<br><br>"



sql = "select a.* from prd_product a " 
sql =sql&" where ROWNUM <= 5 and a.product_type_id="&TypID&" " 
Response.Write "<br>1."&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " - "&rs("Product_ID")
  Response.Write " - "&rs("product_type_id")
  Response.Write " - "&rs("NAME")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 "&Now()&"<br><br>"


sql = "select a.* " 'Product_ID,product_type_id,NAME
sql =sql&" from prd_product a " 
sql =sql&" inner join PRD_COMBIN b on (a.PRODUCT_ID=b.PRODUCT_ID) " '
sql =sql&" where ROWNUM <= 5 and a.product_type_id="&TypID&" "

Response.Write "<br>2."&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " - "&rs("Product_ID")
  Response.Write " - "&rs("product_type_id")
  Response.Write " - "&rs("NAME")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 "&Now()&"<br><br>"



'typeid=195　表示产品类别是ＭＴＣ专线 
'产品列表SQL 
typeid = 142 
nowdate = "2009-05-28" '2007-4-15(s),2008-4-15(v),2008-9-15(v),2008-12-15(v),2009-03-15(v2),2009-05-15(v2),2009-05-28(v2)
sql =    " select a.product_id as Product_ID,a.product_type_id as product_type_id,a.daysconsuming as daycount,a.name as PrdName,a.TypeFlag as TypeFlag,b.prices as Prices "
sql =sql&" from prd_product a "
sql =sql&" inner join ( select max(product_id) as product_id,min(adult_price) as prices,customertypeid as customertypeid "
sql =sql&"              from PRD_COMBIN c "
sql =sql&"              where TO_CHAR(c.EFF_DATE,'YYYY-MM-DD')<='"&nowdate&"' "
sql =sql&"                 and TO_CHAR(c.EXP_DATE,'YYYY-MM-DD')>='"&nowdate&"' "
sql =sql&"                 and TO_CHAR(c.End_Date,'YYYY-MM-DD')>='"&nowdate&"' and c.state=1 "
sql =sql&"              group by c.product_id,c.customertypeid"
sql =sql&"      ) b on(a.product_id=b.product_id)  "
sql =sql&" where b.customertypeid=1 " 
sql =sql&" and a.product_type_id in ( "
sql =sql&"        select product_type_id from domain_product_type "
sql =sql&"        where parent_product_type ="&typeid&" " ' or product_type_id="&typeid&"
sql =sql&"      ) "
sql =sql&" and a.owner_entity_id=1 " 
Response.Write "<br><b>1.org.data.list</b>"&vbcrlf&"<!--"&sql&"-->"
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"
  Response.Write " - "&rs("Product_ID")
  Response.Write " - "&rs("product_type_id")
  Response.Write " - "&rs("PrdName")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 资料测试 "&Now()&"<br><br>"

'Response.End()
'' left    join 




%>


</body>
</html>
