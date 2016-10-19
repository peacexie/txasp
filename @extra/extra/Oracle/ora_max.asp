<!--#include file="inc_perm.asp"-->
<%
sTab = ""

sTab = sTab& "PRD_SALE_REL (2834747) " 
sTab = sTab& "PRD_PUBLISH_CONFIG (1343302) " 
sTab = sTab& "CST_CUSTOMER_ENTITY (479520) " 
sTab = sTab& "CST_CUSTOMER (479480) " 
sTab = sTab& "SALE_CONSUMER (347440) " 
sTab = sTab& "SYS_ORGAN_APP_SCOPE_ORGAN (274332) " 
sTab = sTab& "SYS_STAFF_APP_ORGAN (198611) " 
sTab = sTab& "PRD_PUBLISHELEMENT (180955) " 
sTab = sTab& "PRD_CONF (143195) " 
sTab = sTab& "FIN_INVOICES_SALE (118666) " 
sTab = sTab& "COUNT_LIST (114118) " 
sTab = sTab& "CST_INTENT_HIS (99003) " 
sTab = sTab& "PRI_PRODUCT_PRICE_TYPE (77073) " 
sTab = sTab& "PRD_COMBINHIS (76835) " 
sTab = sTab& "CST_INTENT (76715) " 
sTab = sTab& "SYS_LOGINUSER (76624) " 
sTab = sTab& "SALE_GROUP_ORDER (60590) " 
sTab = sTab& "SALE_GROUP_ORDER_HIS (60250) " 
'sTab = sTab& "TESTWSF (55608) " 
sTab = sTab& "SALE_REC_DETAIL (53940) " 
sTab = sTab& "PRD_GROUP_DETAIL (53024) " 
sTab = sTab& "SALE_RECEIVABLE (51143) " 
sTab = sTab& "PRD_PUNHCOST (46592) "

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
</table>
</body>
</html>
