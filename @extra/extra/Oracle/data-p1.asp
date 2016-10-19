
<!--#include file="inc_perm.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test Products ... </title>
</head>

<body>

<%


''//////////// Oracle 属性测试
sql = "select PRICE_ID,BABY_PRICE,UPDATE_DATE from PRD_PRODUCTPRICE "
sql = sql& " WHERE "
sql = sql& " ROWNUM <= 5 AND "                  ''// 不支持Top N
sql = sql& " TO_CHAR (UPDATE_DATE, 'YYYY-MM-DD')>'2008-04-20' "
'sql = sql& " UPDATE_DATE>#2008-04-20# "        ''//驱动程序不支持所需的属性
'sql = sql& " UPDATE_DATE>'2008-04-20' "        '' 无资料，资料不正确
' // name='产品类型树'                          '' 不支持 [ColName]
sql = sql& " ORDER BY UPDATE_DATE DESC " 
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"&rs("BABY_PRICE")&" - "&rs("UPDATE_DATE")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br>End Oracle 属性测试 "&Now()&"<br><br>"



''//////////// 测试 domain_product_type
sql = "select product_type_id,parent_product_type,name from domain_product_type "
sql = sql& " WHERE "
sql = sql& " product_type_id=142 or parent_product_type=142 "  
' or product_type_id=123 or parent_product_type=123 
sql = sql& " ORDER BY parent_product_type DESC,product_type_id DESC " 
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof  
  Response.Write "<br>"&rs("product_type_id")&" - "&rs("parent_product_type")&" - "&rs("name")
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing
Response.Write "<br><br><b>以上资料</b>End domain_product_type 测试 "&Now()&"<br><br>"



''//////////// domain_product_type Tree 测试
sql = "select name,parent_product_type,product_type_id from domain_product_type "
sql = sql& " where name='产品类型树' or parent_product_type=0 "
sql = sql& " order by product_type_id "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1
Do While NOT rs.Eof
  Response.Write "<br>"&rs("product_type_id")&" - "&rs("name")
  pTypID = rs("product_type_id")
  sql2 = "select name,parent_product_type,product_type_id from domain_product_type "
  sql2 = sql2& " where parent_product_type="&pTypID&"  "
  sql2 = sql2& " order by product_type_id "
  SET rs2=Server.CreateObject("Adodb.Recordset") 
  rs2.Open sql2,conn,1,1
  Do While NOT rs2.Eof
    Response.Write "<br> ------ "&rs2("product_type_id")&" - "&rs2("parent_product_type")&" - "&rs2("name")
  rs2.MoveNext()
  Loop
  rs2.Close
  SET rs2=Nothing
  
rs.MoveNext()
Loop
rs.Close
SET rs=Nothing




''//////////// Oracle 正式数据测试

typeid  = 150 
nowdate = "2009-12-31"

sql = "select a.product_id as Product_ID,a.product_type_id,a.daysconsuming as daycount,a.name as PrdName,a.TypeFlag as TypeFlag,b.prices as Prices "
sql =sql&" from prd_product a "
sql =sql&" inner join ( select max(product_id) as product_id,min(adult_price) as prices,customertypeid as customertypeid  from prd_combin c "
sql =sql&"              where TO_CHAR (c.EFF_DATE, 'YYYY-MM-DD')<='"+nowdate+"'  and TO_CHAR(c.EXP_DATE,'YYYY-MM-DD')>='"+nowdate+"' and TO_CHAR(c.End_Date,'YYYY-MM-DD')>='"+nowdate+"' "
sql =sql&"              and c.state=1 group by c.product_id,c.customertypeid"
sql =sql&"            ) b on(a.product_id=b.product_id) "
sql =sql&" where  b.customertypeid=1 and a.product_type_id in( select product_type_id from domain_product_type  where parent_product_type ="&typeid&" or product_type_id="&typeid&")  and a.owner_entity_id=1" 

sql = "select * from prd_product where product_type_id in( select product_type_id from domain_product_type where parent_product_type ="&typeid&" or product_type_id="&typeid&")"








Response.Write "<br><br> "&Now()&"<br><br>"





%>

</body>
</html>
