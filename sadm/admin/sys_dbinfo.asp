<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<%

yAct = Request("yAct")
KW = RequestS("KW","C",32)
TP = RequestS("TP",3,24)
Page = RequestS("Page","N",1)
TabID = RequestS("TabID","C",32)


If TabID="InfoPics" Then
  TabCol = Split("KeyCode;InfSubj;ImgName;LogATime",";")
  TabMsg = Split("KeyCode;主题;ImgName;LogATime",";")
  TabKey = "KeyID"
  TabGap = "'"
  TabOrd = "LogATime DESC"
Else
  TabID="InfoNews"
  TabCol = Split("KeyMod;KeyCode;InfSubj;LogATime",";")
  TabMsg = Split("KeyMod;KeyCode;主题;LogATime",";")
  TabKey = "KeyID"
  TabGap = "'"
  TabOrd = "LogATime DESC"
End If


if yAct="Del" then
  For iy = 1 To Request.Form("yID").Count
    iID = Request.Form("yID").item(iy)
	Call rs_DoSql(conn,"DELETE FROM ["&TabID&"] WHERE "&TabKey&"="&TabGap&""&iID&""&TabGap&"")
  Next
  Msg = "信息 删除完成！" 
else
  Msg = "警告 请小心管理！如误操作，自行负责！" 
end if


sqlk = ""
if KW<>"" AND TP<>"" then
  sqlk = " WHERE ( "&TP&" LIKE '%"&KW&"%' ) " 
end if


   sql = "SELECT * FROM "&TabID&sqlK&" ORDER BY "&TabOrd&""
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td colspan="<%=uBound(TabCol)+2%>" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>[<%=TabID%>]数据管理</strong></td>
          <td width="40%" align="left" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <td align="center" nowrap><a href="?TabID=InfoNews">InfoNews</a> - <a href="?TabID=InfoPics">InfoPics</a></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <form name="form1" method="post" action="?">
            <td colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;&nbsp; &nbsp;
              <input name="TabID" type="hidden" id="TabID" value="<%=TabID%>">
              <select name="TP" id="TP">
                <option value="" selected>[Null]</option>
                <option value="<%=TabKey%>"><%=TabKey%></option>
				<%
	For j=0 To uBound(TabCol)
	ColVal = rs(TabCol(j))
	%>
                <option value="<%=TabCol(j)%>"><%=TabCol(j)%></option>
                <%Next%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
            <input type="submit" name="Submit" value="搜索"></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td colspan="<%=uBound(TabCol)+2%>" align="center" bgcolor="f8f8f8"><%=RS_Page(rs,Page,"?send=pag&KW="&KW,1)%></td>
  </tr>
  <tr align="center">
    <td nowrap>ID</td>
    <%For j=0 To uBound(TabMsg)%>
    <td nowrap><%=TabMsg(j)%></td>
    <%Next%>
  </tr>
  <form id="fm01y" name="fm01y" action="?" method="post">
    <%

  if not rs.eof then
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f8f8f8"
		end if
KeyID = rs(TabKey)
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap> <%=KeyID%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <%
	For j=0 To uBound(TabCol)
	ColVal = rs(TabCol(j))
	%>
      <td nowrap><%=ColVal%></td>
      <%Next%>
    </tr>
    <%
  rs.movenext
  if rs.eof then exit for
  next
  %>
    <tr align="center" bgcolor="E0E0E0">
      <td colspan="<%=uBound(TabCol)+2%>" align="right" nowrap><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="30%" align="right"><input name="KW" type="hidden" id="KW" value="<%=KW%>">
              <input name="TP" type="hidden" id="TP" value="<%=TP%>">
              <input name="Page" type="hidden" id="Page" value="<%=Page%>">
              <input name="TabID" type="hidden" id="TabID" value="<%=TabID%>">
              <span id="yFlag" style="visibility:hidden ">N</span>
              <input name="yAll" type="checkbox" id="yAll" onClick="ySel()">
              全选</td>
            <td width="10%" align="left" nowrap>&nbsp;</td>
            <td colspan="<%=uBound(TabCol)%>" align="left" nowrap><select name="yAct" id="yAct">
                <option value="Del" selected >删除</option>
              </select>
              <input type="submit" name="Submit2" value="执行"></td>
          </tr>
        </table></td>
    </tr>
    <%
  rs.Close
  end if
  set rs = nothing
%>
  </form>
</table>
<script type="text/javascript">
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
</body>
</html>
