<!--#include file="inc_perm.asp"-->
<%

Server.ScriptTimeout = 3600  'Second 999999
send = Request("send")
Page = RequestS("Page","N",1)
xTab = "SYS_MAX" 'RequestS("xTab","C",96)
xWhr = Request("xWhr") :Response.Write " --- "&xWhr

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


sqlOrd = " ORDER BY TABNAME,MAX_VALUE DESC "


  sqlCols = "*" :fid = 0
  sql00 = " SELECT "&sqlCols&" FROM "&xTab&" WHERE ROWNUM =1 "
  Response.Write sql00
  SET rs=Server.CreateObject("Adodb.Recordset")
  rs.Open sql00,conn,1,1
  vdt1 = ""
  vdt2 = ""
  vdt3 = ""
  FOR i = 0 TO rs.Fields.Count-1
	f1 = rs.Fields(i).Name
	ft = rs.Fields(i).Type
	fs = rs.Fields(i).DefinedSize
	rsf0 = rs.Fields(0).Name
	vdt1 = vdt1 & "<td>"&f1&"</td>"
	vdt2 = vdt2 & "<td>"&ft&"</td>"
	vdt3 = vdt3 & "<td>"&fs&"</td>"
  NEXT
  rs.Close
  SET rs=Nothing
'End If
  
  


%>
<table border="0" cellpadding="1" cellspacing="1">
  <tr>
    <td>F.Type</td>
    <%= vdt2 %></tr>
  <tr>
    <td>F.Size</td>
    <%= vdt3 %></tr>
  <tr bgcolor="#CCCCCC">
    <td>Fields</td>
    <%= vdt1 %></tr>
  <%

   Set cn = Server.CreateObject("ADODB.Connection")
   cn.ConnectionTimeout = 480 '
   cn.CommandTimeout = 481 
   cn.Open conn
   
  sql2 = " SELECT * FROM "&xTab&" WHERE "
  sql2 = sql2& " (TABNAME LIKE 'SYS_%') OR (TABNAME LIKE 'cst_%') OR (TABNAME LIKE 'sale_%') OR (TABNAME LIKE 'web_%')" 
  sql2 = sql2& sqlOrd
  
  'Response.Write " : "&sql2
	  
	  SET rs2=Server.CreateObject("Adodb.Recordset")  
      rs2.Open sql2 ,cn,1,1
	  rs2.PageSize = 720 ':Response.Write sql2
  if not rs2.eof then
    if int(Page)>rs2.PageCount or int(Page)<1 Then
      Page = 1
    end if
%>
  <form id="fm01y" name="fm01y" action="?" method="post">
    <%
  rs2.AbsolutePage = Page
  j = 0 
  for irec = 1 to rs2.PageSize

	    j = j + 1
%>
    <tr>
      <td nowrap><input 
			name="yID" type="checkbox" id="yID" value="<%=rs2.Fields(0)%>">
        <input name="KeyID" type="hidden" id="KeyID" value="<%=rs2.Fields(0).Name%>">
        <%=(Page-1)*rs2.PageSize+irec%></td>
      <%
		for k = 0 to i-1
		  tVal = rs2.Fields(k)&""
		  tVal = LeftB(tVal,48)
		  tVal = Replace(tVal,"<","&lt")
		  tVal = Replace(tVal,">","&gt")
		  tNam = rs2.Fields("TABNAME")
		  If tNam="cst_customer" Or tNam="SYS_STAFF" Or tNam="sale_order" Or tNam="sale_group_order" Then 'SALE_REC_DETAIL
	  	    tVal = "<span style='color:#00F'>"&tVal&"</span>"
		  End If
		  
		    Response.Write "<td nowrap>"&tVal&"</td>"
		next
%>
    </tr>
    <%
      rs2.movenext
  if rs2.eof then exit for
  next
%>
    <tr>
      <td bgcolor="#CCCCCC" width="20"><input name="yAll" type="checkbox" id="yAll" onClick="ySel()">
        <span id="yFlag" style="visibility:hidden ">N</span></td>
      <td colspan="255" bgcolor="#CCCCCC"><input name="xTab" type="hidden" value="<%=xTab%>">
        <input name="xWhr" type="hidden" value="<%=xWhr%>">
        <input name="Page" type="hidden" value="<%=Page%>">
        <input name="send" type="hidden" id="send" value="yDel">
        <input type="submit" name="Submit" value="删除选择"></td>
    </tr>
  </form>
</table>
<%

end if
Response.Write RS_PageThis(rs2,Page)

%>
<script>
function chPage(xPage)
{
    document.fm01p.Page.value=xPage;
	document.fm01p.submit();
}

</script>
<script>
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.fm01y.yID.length;i++)
   {document.fm01y.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.fm01y.yID.length;i++)
   {document.fm01y.yID.item(i).checked=false;}
   }
}  
</script>
<table border="0" cellpadding="1" cellspacing="1">
  <tr>
    <form action="?" method="post" name="fm01p">
      <td nowrap bgcolor="#CCCCCC">Where:
        <input name="xWhr" type="text" id="xWhr" value="<%=xWhr%>" size="24" maxlength="96">
        <input type="submit" name="Submit" value="Submit">
        <input name="xTab" type="hidden" value="<%=xTab%>">
        <input name="send" type="hidden" id="send" value="SDel">
        <input name="Page" type="hidden" value="<%=Page%>"></td>
    </form>
  </tr>
</table>
<%
rs2.Close
SET rs2=Nothing  
   cn.Close()
   Set cn = Nothing

%>
</body>
</html>
