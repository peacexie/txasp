<!--#include file="dinc/_config.asp"-->
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
</head>
<body style="padding:8px" style="background-color:#FFF">

<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
ID = RequestS("ID",3,48)
  If KW&"" <> "" Then
    sqlK = sqlK & " AND (LogAUser LIKE '%"&KW&"%') " 
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
End If
  sql = " SELECT * FROM DocsLogs "
  sql =sql& " WHERE KeyID='"&ID&"' "&sqlK
  sql =sql& " ORDER BY LogATime DESC" 'SetTop,
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="CCCCFF">
  <tr align="center" bgcolor="CCCCFF">
    <td height="27" colspan="4" align="center" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>公文查阅记录</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
关键字:
<input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="4" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&ID="&ID&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="right" nowrap bgcolor="#DDDDFF">NO</td>
    <td height="27" align="center" bgcolor="#DDDDFF">查看人</td>
    <td height="27" align="center" bgcolor="#DDDDFF">查看次</td>
    <td height="27" align="center" nowrap bgcolor="#DDDDFF">查看时间</td>
  </tr>

  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F0F0F0"
		End If
KeyID = rs("KeyID")
KeyN = rs("KeyN") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td align="center"><%=LogAUser%></td>
      <td align="center"><%=KeyN%></td>
      <td align="center" nowrap><%=LogATime%></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap bgcolor="#DDDDFF">
      <span id="yFlag" style="visibility:hidden ">N</span>        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td width="20%" bgcolor="#DDDDFF">全选
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
      <td colspan="2" align="right" nowrap bgcolor="#DDDDFF"><select name="yAct" id="yAct" >
          <option value="del_sel">删除</option>
            </select>        <input type="submit" name="Submit" value="执行"></td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="6">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>

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
