<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
	If KW&"" <> "" Then
	  sqlK = sqlK & " AND (InfSubj LIKE '%"&KW&"%') " 
	End If

cID = 0
sID = ""
If yAct="del_sel" then
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"DELETE FROM DocsLogs WHERE KeyID='"&iID&"'")
	End If
  Next
	
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="d_Clear" then
    sql2 = "DELETE FROM DocsLogs WHERE KeyID NOT IN(SELECT KeyID FROM DocsNews) "
    Call rs_DoSql(conn,sql2)
    sql2 = "DELETE FROM DocsLogs WHERE LogAUser NOT IN(SELECT UsrID FROM AdmUser"&Adm_aUser&") "
    Call rs_DoSql(conn,sql2)
    Msg = " 清理成功!"
End If
    sql = " SELECT DocsLogs.* "
	sql =sql& " ,DocsNews.InfSubj,DocsNews.LogATime AS PubTime,DocsNews.LogAUser AS PubUser FROM DocsLogs,DocsNews "
	sql =sql& " WHERE DocsLogs.KeyID=DocsNews.KeyID "&sqlK
	sql =sql& " ORDER BY DocsLogs.LogATime DESC" 'SetTop,
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[公文查阅记录]管理</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="TPU" type="hidden" id="TPU" value="<%=IDPerm%>">
              关键字:
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TUP="&ModID&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="20%" height="27" align="center">主题</td>
    <td width="30%" height="27" align="center" nowrap>发布人</td>
    <td width="15%" align="center" nowrap>发布时间</td>
    <td width="15%" align="center" nowrap>查看人</td>
    <td width="15%" align="center" nowrap>查看次</td>
    <td width="15%" height="27" align="center" nowrap>查看时间</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
  </tr>
  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
InfSubj = Show_Text(rs("InfSubj"))
PubTime = rs("PubTime")
PubUser = rs("PubUser")

KeyN = rs("KeyN") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="doc_view.asp?ID=<%=KeyID%>&Act=Adm" target="_blank"><%=InfSubj%></a></td>
      <td align="center" nowrap><%=PubUser%></td>
      <td align="center" nowrap><%=PubTime%></td>
      <td align="center" nowrap><%=LogAUser%></td>
      <td align="center" nowrap><%=KeyN%></td>
      <td align="center" nowrap><%=LogATime%></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>全选
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TPU" type="hidden" id="TPU" value="<%=IDPerm%>"></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <option value="d_Clear">清理.垃圾</option>
            </select></td>
      <td align="left" nowrap>&nbsp;</td>
      <td align="left" nowrap><input type="submit" name="Submit" value="执行"></td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="9">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="9" align="right"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>
</body>
</html>
