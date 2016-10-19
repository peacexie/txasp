<!--#include file="config.asp"-->
<!--#include file="sys.funcs.asp"-->
<!--#include file="../../sadm/func2/cch_Class.asp"-->

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

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
TP = "ParConfig"
Set rs=Server.Createobject("Adodb.Recordset")

If yAct="del" Then
 ID = RequestS("ID","C",48)
   sql = "DELETE FROM AdmPara WHERE ParCode='"&ID&"'"
   Call rs_DoSql(conn,sql)
   Call Add_Log(conn,Session("UsrID"),"删除:站点配置!!!"&KeyID&"","[sadm/info_config]",Msg)
   Msg = cID&" 条记录删除成功!"
ElseIf yAct="edt" Then
  sql="UPDATE [AdmPara] SET ParName='"&RequestS("ParName",C,24)&"',[ParText]='"&RequestS("ParText",C,1200)&"'"
  sql=sql&" WHERE ParCode='"&RequestS("ParCode",C,48)&"' "
  'Response.Write sql
  Call rs_DoSql(conn,sql) 
ElseIf yAct="upd" Then

  DBCode = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_Code'") 
  DBPath = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_Path'") 
  If DBCode<>Config_Code And DBPath<>Config_Path Then
   Msg1 = "初始化ID 和 系统路径 不能同时修改刷新！\n当前初始化ID: "&Config_Code&"\n当前初始化ID: "&Config_Path&"\n请还原！！！"
   Response.Write js_Alert(Msg1,"Redir","?") 
   Response.End()
  Else

  s = SysConfig() '取出最新数据 '跟新配置日期 '关闭(否则出错)
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParDate='"&Date()&"' WHERE ParCode='dateConfig'") 
  Set rs = Nothing 
  
  If DBCode<>Config_Code Then '移动数据库
   objDB=Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_"&DBCode&".Peace!DB")
   orgDB=Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_"&Config_Code&".Peace!DB")
   Call fil_move(orgDB,objDB)
   Msg1 = "\n改变了初始化ID,所有的加密算发参数都改变了，请马上修改管理员密码！"
  End If

  s="<"&"%" &s&vbcrlf& "%"&">" '更新文件
  Call File_Add2("../../upfile/sys/config/Config.asp",s,"UTF-8")
  echo(s)
  
  If DBPath=Config_Path Then  '更新Session, 提示
   Session("UsrP"&DBCode) = Session(UsrPStr)
   Response.Write js_Alert("刷新配置完毕！"&Msg1,"Redir","?") 
  Else
   Msg1 = "\n改变了系统路径,如果出错，请手动把文件移动到新目录！\n"&DBPath&""
   Response.Write js_Alert(Msg1,"","?") 'Alert
   Response.End()
  End If

 End If

End If

    sql = " SELECT * FROM [AdmPara] WHERE ParFlag='"&TP&"'"
	sql =sql& " ORDER BY ParName,ParCode" 
   
   rs.Open Sql,conn,1,1
   rs.PageSize = 240
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="6" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
      <tr align="center" bgcolor="#FFFFFF">
        <td width="20%" align="center" bgcolor="#FFFFFF"><strong>站点配置</strong></td>
        <td align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;</td>
        <!--#include file="inc_menu.asp"-->
      </tr>
    </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="24" align="center" nowrap>NO</td>
    <td align="center">编码</td>
    <td align="center" nowrap bgcolor="E0E0E0">名称</td>
    <td align="center" bgcolor="E0E0E0">站点配置-参数值</td>
    <td align="center" nowrap bgcolor="E0E0E0">修改</td>
    <td align="center" nowrap bgcolor="E0E0E0">删除</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="6" align="right" nowrap></td>
  </tr>
  
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
ParCode = rs("ParCode")
ParName = rs("ParName")
ParText = Show_Form(Left(rs("ParText")&"",120))
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")

	  %>
    <form name="flist<%=i%>" method="post" action="?">
	<tr bgcolor="<%=col%>">
      <td align="right" nowrap><input type="hidden" name="hiddenField"><%=i%></td>
      <td><%=ParCode%>
      <input name="ParCode" type="hidden" id="ParCode" value="<%=ParCode%>"></td>
      <td align="left" nowrap><input name="ParName" type="text" id="ParName" size="15" maxlength="120" value="<%=ParName%>">
<input name="yAct" type="hidden" id="yAct" value="edt"></td>
      <td align="left" bgcolor="<%=col%>"><input name="ParText" type="text" id="ParText" size="48" maxlength="120" value="<%=ParText%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">      </td>
      <td align="center" nowrap>
      <input type="submit" name="Submit" value="修改" <%=fDis%>></td>
      <td align="center" nowrap>
      <%If inStr("Code,Path,Name",Mid(ParCode,8))>0 Then%>
      <span class="col999">删除</span>  
      <%Else%>
      <a onClick="Del_YN('?ID=<%=ParCode%>&yAct=del&TP=<%=TP%>','确认删除?小心操作哦！\n也许使你系统崩溃？！！')" href="#" >删除</a>
      <%End If%>
      </td>
      </tr>
	 </form>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>

<%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="6" align="left" bgcolor="#F0F0F0"><span style="width:640px; text-align:left; padding:8px; margin:0px auto"><span class="col00F">注意</span>：<font color="#FF0000"><%=msg%></font> <br />
-=&gt;  此部分 [站点配置] 开发时已经设置好，请不要轻易修改！本信息刷新到 '/upfile/sys/config/Config.asp 文件中；<br />
-=&gt;初始化ID，系统路径：用于重新配置网站，不要轻易改动；<br />
&nbsp;&nbsp;&nbsp;&nbsp; 参考随机初始化ID：(<span class="col00F"><%=Get_GUID(Config_Vers,"PEACE"&Rnd_ID("0",1)&"ASP"&Rnd_ID("0",1)&"")%></span>) ；<br />
-=&gt;URL地址：域名改变时，设置这个参数，如从测试域名转到正式域名；<br />
-=&gt;语言版本：与前台一起控制，格式：0,1,2,3...;繁体,中文,英文,日文...；<br />
-=&gt;内容存文件：File-存文件；DB-存数据库；Html-静态； <br />
-=&gt;分公司／部门：与其他程序一起使用；<br />
-=&gt;附件大小(KB)：主要是指信息发布时的缩略图和编辑器中的图片，文件大小控制，建议200以内，单位KB；<br />
-=&gt;图片同步：如果多个语言版本，图片，产品等模块共用图片；<br>
-=&gt;认证码配置参考:
    2233444477(2:2:4:2)数字,字母,运算; <br>
    &nbsp;&nbsp;&nbsp;&nbsp; 5566(2:2)汉字; hhiijj(2:2:2)小号数字,字母; x(1)外部调用; 234567hijx(全部)      ；</span></td>
  </tr>
</table>


</body>
</html>
