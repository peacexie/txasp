<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

TypID     = RequestS("TypID",3,24)
TypNam2   = RequestS("TypNam2",3,255) 
ImgName   = RequestS("ImgName",3,255) 
TypResume = RequestS("TypResume",3,255) 
send      = Request("send") 

If send = "edt" Then
    sql = " UPDATE [WebTyps] SET TypNam2='"&TypNam2&"'"
	'If ImgName<>"" Then
	sql = sql & " ,ImgName='"&ImgName&"'"
	'End If
	sql = sql & " ,TypResume='"&TypResume&"' "
	sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & "WHERE TypID='"&TypID&"' AND TypMod='"&ModID&"'  "
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
Else

End If

If Left(ModID,3)="BBS" Then
  flgName = "版主(会员ID用,分开)"
ElseIf ModID="InfN624" Then
  flgName = "首页标记"
Else
  flgName = "扩展标记"
End If

    sql = " SELECT * FROM [WebTyps] "
	sql =sql& " WHERE TypMod='"&ModID&"'  "
	sql =sql& " ORDER BY TypLayer,TypID " 'LEN(TypLayer)
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1

%>
<FIELDSET>
<LEGEND>
<table width="100%" border="1" cellpadding="1" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="20%" align="center" nowrap><strong>类别设定</strong></td>
    <td width="20%" align="center" nowrap><a href="type_list.asp?send=upd">刷新</a>|<a href="../type/type_list.asp?ModID=<%=ModID%>">返回</a></td>
    <td nowrap><font color="#FF0000"><%=msg%>&nbsp;</font></td>
    <td width="15%" align="right" nowrap>&nbsp;</td>
  </tr>
</table>
</LEGEND>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="f0f0f0">
  <tr align="center" bgcolor="E0E0E0">
    <td align="center">NO</td>
    <td align="center">类别名称</td>
    <td align="center">层级</td>
    <td align="center"><%=flgName%></td>
    <td align="center">类别图片</td>
    <td align="center">类别描述</td>
    <td align="center">修改</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="7" align="right"></td>
  </tr>
  <%
	  i = 0
	  Do While Not rs.eof 
	    i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#F8F8F8"
		end if
		TypID = rs("TypID")
		TypLayer = rs("TypLayer")
		TypName = Trim(rs("TypName")&"")
		TypNam2 = Trim(rs("TypNam2")&"")
		ImgName = Trim(rs("ImgName")&"")
		TypResume = Trim(rs("TypResume")&"")
		nLayer = Len(TypLayer)-Len(Replace(TypLayer,";",""))
		nLayStr = ""
		If nLayer > 0 Then
		  For j = 1 To nLayer-1
		    nLayStr = nLayStr & "│ "
		  Next
		End If
		nLayStr = nLayStr & "├─"
	  %>
  <form name="fm<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>" OnMouseOver='this.bgColor="#FFFFCC"' onMouseOut='this.bgColor="<%=col%>"'>
      <td align="right"><%=i%></td>
      <td nowrap><%=nLayStr%><%=TypName%>
        <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
        <input name="send" type="hidden" id="send" value="edt"></td>
      <td align="center"><%=nLayer%></td>
      <td align="center">
      <%If ModID="InfN624" Then%>
      <select name="TypNam2" id="TypNam2">
      <%=Get_SOpt("Home;Inn","首页;隐藏",TypNam2,"")%>
      </select>
	  <%Else%>
      <input name="TypNam2" type="text" id="TypNam2"  value="<%=TypNam2%>" size="15" maxlength="120">
      <%End If%>
      </td>
      <td align="left"><input name="ImgName" type="text" id="ImgName" value="<%=ImgName%>" size="30" maxlength="120">
      <input type="button" name="Submit" value="选择" onClick="getRetObject(<%=i%>);window.open('../file/file_view.asp?yPath=myfile/type/')"> 
      </td>
      <td align="center"><input name="TypResume" type="text" id="TypResume" value="<%=TypResume%>" size="8" maxlength="120"></td>
      <td align="center"><input type="submit" name="Submit" value="修改"></td>
    </tr>
  </form>
  <%
	    rs.movenext
	  loop
	  
	  rs.close()
	  set rs = nothing
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="7" align="right"></td>
  </tr>
</table>
</FIELDSET>
<script type="text/javascript">
var yFile;
function getRetObject(i)
{
	yFile = eval("document.fm"+i+".ImgName");
}
</script>
</body>
</html>
