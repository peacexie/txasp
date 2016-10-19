<!--#include file="inc_perm.asp"-->
<%
sTab = ""
sTab = sTab& "CST_CUSTOMER (客户资料表) "
sTab = sTab& "DOMAIN_STATE (状态表) "
sTab = sTab& "SALE_CONSUMER (订单消费者) " 
sTab = sTab& "SALE_GROUP_ORDER (团队子订单表) " 
sTab = sTab& "SALE_GROUP_ORDER_HIS (团队订单变更表) " 
'sTab = sTab& "SALE_ORDER (订单表) "
sTab = sTab& "SALE_ORDER (订单表) "
sTab = sTab& "SALE_REC_DETAIL (订单收费明细表) " 
sTab = sTab& "SALE_SUFFIX_REQUIREMENT (附加需求表) "
sTab = sTab& "PRD_COM (组合资源_需) "
'sTab = sTab& "PRD_COMBIN (发布组合产品) " 
sTab = sTab& "PRD_COMBIN (发布组合产品) "
sTab = sTab& "PRD_COMBINDETAIL (发布产品的产品项) " 
sTab = sTab& "PRD_COMBINDETAILHIS (发布产品的产品项历史) " 
sTab = sTab& "PRD_COMBINHIS (发布产品的产品项历史) " 
sTab = sTab& "PRD_PRODUCT (产品信息) " 
sTab = sTab& "PRD_PRODUCTDETAIL (产品项) " 
sTab = sTab& "PRD_PRODUCTELEMENT (产品元素及成本) "
sTab = sTab& "PRD_PRODUCTPRICE (产品价格) " 
sTab = sTab& "PRD_PUBLISHELEMENT (产品价格) " 
sTab = sTab& "SYS_STAFF (操作员) "
sTab = sTab& "DOMAIN_PRODUCT_TYPE (产品类型树) "

aTab = Split(sTab,")")
RecAll = 0

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Tab List</title>
<meta http-equiv="Pragma" content="no-cache">
</head>
<body>
<table width="580"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#f4f4f4">
  <%

For i = 0 To uBound(aTab)-1
 aRow = Split(aTab(i),"(")
 TBName = Replace(aRow(0)," ","")
 TBNRem = aRow(1)
 TBRec  = RequestSafe(rs_Count(conn,TBName),"N",0)
 RecAll = RecAll + TBRec
%>
  <tr bgcolor="#FFFFFF">
    <td align="right">&nbsp;<%=i+1%></td>
    <td><%=TBName%></td>
    <td align="left"><%=TBNRem%></td>
    <td align="right"><%=TBRec%></td>
    <td align="center"><a href="ora_stru.asp?xTab=<%=TBName%>">Stru</a></td>
    <td align="center"><a href="ora_data.asp?xTab=<%=TBName%>" target="_blank">Data</a></td>
  </tr>
  <%
Next
%>
  <tr bgcolor="#FFFFFF">
    <td align="right"><%=i%></td>
    <td align="center">&nbsp;Tabels</td>
    <td align="right">&nbsp;</td>
    <td align="right"><%=RecAll%></td>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
<%  
i = 0
RecAll = 0
sql = "select TABLE_NAME from all_tables where owner = 'anytravel'"
sql = "select * from user_tables"
'sql = "select * from dba_tables"
Set rs = Server.Createobject("Adodb.Recordset")
rs.Open sql,conn,1,1
Do While Not rs.EOF
 TBName = rs("TABLE_NAME")
 TBNRem = "" 'aRow(1)
 TBRec  = RequestSafe(rs_Count(conn,TBName),"N",0)
 RecAll = RecAll + TBRec
%>
  <tr bgcolor="#FFFFFF">
    <td align="right">&nbsp;<%=i+1%></td>
    <td><%=TBName%></td>
    <td align="left"><%=TBNRem%></td>
    <td align="right"><%=TBRec%></td>
    <td align="center"><a href="ora_stru.asp?xTab=<%=TBName%>">Stru</a></td>
    <td align="center"><a href="ora_data.asp?xTab=<%=TBName%>" target="_blank">Data</a></td>
  </tr>
  <%
  i = i + 1
  rs.MoveNext()
Loop 'Next
rs.Close()
SET rs = Nothing 
%>
  <tr bgcolor="#FFFFFF">
    <td align="right"><%=i%></td>
    <td align="center">&nbsp;Tabels</td>
    <td align="right">&nbsp;</td>
    <td align="right"><%=RecAll%></td>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
