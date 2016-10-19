<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

send = Request("send")
Page = RequestS("Page","N",1) 
ModID = "ParMSys" 'RequestS("ModID","C",24) 
If send = "ins" then

ElseIf send = "edt" Then
ParCode = RequestS("ParCode",3,24)
sql = " UPDATE [AdmPara] SET " 
sql = sql& " ParName = '" & RequestS("ParName",3,48) &"'" 
sql = sql& ",ParRem = '" & RequestS("ParRem",3,8000) &"' " 
sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
sql = sql& " WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 修改成功!"
ElseIf send = "del" Then
  
End If

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParRem "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE ParFlag='" & ModID &"' AND ParCode LIKE 'rX%' "
sql = sql & " ORDER BY ParCode "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
rs.PageSize = 240 
%>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="e0e0e0">
  <tr align="center" bgcolor="f8f8f8">
    <td colspan="3"><table width="100%"  border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td width="30%" align="center"><strong>系统菜单参数 设置</strong></td>
          <td align="right"><font color="#FF0000"><%=msg%></font>
          </td>
        </tr>

      </table></td>
  </tr>
  <tr bgcolor="e0e0e0">
    <td height="23" align="center" bgcolor="e0e0e0">&nbsp;编码 / 名称</td>
    <td bgcolor="e0e0e0">&nbsp;参数值 </td>
    <td nowrap bgcolor="e0e0e0">操作</td>
  </tr>
  <tr>
    <td colspan="3" align="center" bgcolor="999999"></td>
  </tr>
  <%
		  if NOT rs.EOF then

if int(Page)>rs.PageCount or int(Page)<0 Then 
   Page = 1 
End If
rs.AbsolutePage = Page 
     for i=1 to rs.PageSize 
if i mod 2 = 1 then
   col = "F8F8F8"
else
   col = "FFFFF8"
end if
ParName = Show_Form(rs("ParName"))
ParCode = Show_Form(rs("ParCode"))
ParRem = Show_Form(rs("ParRem"))
'ParRem = Replace(ParRem,chr(34),"")
If ModID="ParFil" Then
  wrap = "wrap='ON'"
ElseIf Right(ParCode,4)=".asp" Then 
  wrap = "wrap='ON'"
Else
  wrap = "wrap='OFF'"
End If
%>
  <form name="fmedit<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="center" nowrap bgcolor="<%=col%>">&nbsp;参数编码<br>
        <input name='ParCode' type='text' value="<%=ParCode%>" size='10' maxlength='12' readonly>
        <br>
&nbsp;参数名称<br>
        <input name='ParName' type='text' value="<%=ParName%>" size='10' maxlength='24'>
      </td>
      <td bgcolor="<%=col%>"><textarea name="ParRem" cols="80" rows="6" <%=wrap%> 
			  onBlur="chkF_Len(document.fmedit<%=i%>.ParRem,64000,'内容太长!')"><%=ParRem%></textarea></td>
      <td nowrap bgcolor="<%=col%>"><input type="submit" name="Submit" value="保存">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <br>
        <input type="button" name="Button" value="查看" onClick="prview('<%=ParCode%>')">
        <input name="send" type="hidden" id="send" value="edt">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
    </tr>
  </form>
  <% 
rs.MoveNext 
if rs.eof then exit for 
     next 
   end if ' //////////////////////////////////////////
%>
</table>
<%
rs.Close()
SET rs=Nothing
%>
<script type="text/javascript">
function prview(ParCode)
{
  window.open('../system/para_view.asp?ParCode='+ParCode);
}</script>
</body>
</html>
