<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

Page = RequestS("Page","N",1)

send = Request("send")
Page = RequestS("Page","N",1) 
ModID = RequestS("ModID","C",36)
If ModID="ParConfig" Then
  cDis = " disabled='disabled' "
Else
  cDis = ""
End If 
ModName = rs_Val("","SELECT [SysName] FROM AdmSyst WHERE [SysID]='"&ModID&"'")
If send = "ins" then
ParCode = RequestS("ParCode",3,36)
sql = " INSERT INTO [AdmPara] (" 
sql = sql& "  ParName" 
sql = sql& ", ParCode" 
sql = sql& ", ParFlag" 
sql = sql& ", ParText" 
sql = sql& ",LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & RequestS("ParName",3,48) &"'" 
sql = sql& ", '" & ParCode &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & RequestS("ParText",3,255) &"'" 
sql = sql& " ,'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"'"
sql = sql& ")"
if rs_Exist(conn,"SELECT ParCode FROM [AdmPara] WHERE ParCode='"&ParCode&"'") = "EOF" then
  Call rs_DoSql(conn,sql)
  msg = "增加成功!"
else
  msg = "增加失败!"&ParCode&" 已经存在!"
end if
ElseIf send = "edt" Then
ParCode = RequestS("ParCode",3,24)
sql = " UPDATE [AdmPara] SET " 
sql = sql& " ParName = '" & RequestS("ParName",3,48) &"'" 
sql = sql& ",ParText = '" & RequestS("ParText",3,255) &"' " 
sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
sql = sql& " WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 修改成功!"
ElseIf send = "del" Then
  ParCode = RequestS("ParCode",3,24)
  sql = "DELETE FROM [AdmPara] WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  Call Add_Log(conn,Session("UsrID"),"删除:参数!!!"&ParCode&"","para_text",Msg)
  msg = ""&ParCode&" 删除成功!"
End If

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParText "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE ParFlag='" & ModID &"' "
'If inStr("(ParConfig,ParWMark,)",ModID)>0 Then
  sql = sql & " ORDER BY ParName,ParCode "
'Else
  'sql = sql & " ORDER BY ParCode "
'End If
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
rs.PageSize = 240 

If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="e0e0e0">
  <tr align="center" bgcolor="f8f8f8">
    <td colspan="7"><table width="100%"  border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td width="20%" align="center" nowrap><strong><%=ModName%></strong><br>
            <font color="#FF0000"><%=msg%></font>&nbsp;</td>
          <td align="right" nowrap><!--#include file="para_menu.asp"--></td>
        </tr>
      </table></td>
  </tr>
  <!--RS_Page-->
  <tr bgcolor="e0e0e0">
    <td height="23" align="center">NO</td>
    <td bgcolor="e0e0e0">&nbsp;参数编码</td>
    <td bgcolor="e0e0e0">&nbsp;参数名称</td>
    <td colspan="2" bgcolor="e0e0e0">&nbsp;参数值</td>
    <td align="center">保存</td>
    <td align="center">删除</td>
  </tr>
  <tr>
    <td colspan="7" align="center" bgcolor="999999"></td>
  </tr>
  <%
		  if NOT rs.EOF then

rs.AbsolutePage = Page 
     for i=1 to rs.PageSize 
if i mod 2 = 1 then
   col = "F8F8F8"
else
   col = "FFFFF8"
end if
ParName = Show_Form(rs("ParName"))
ParCode = Show_Form(rs("ParCode"))
ParText = Show_Form(rs("ParText"))
'If ParCode="Config_Mode" AND ParText="isExpert" Or ModID="ParText" Then
 'fExpert="YES"
'End If
' fisExpert = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_Mode'")
%>
  <form name="fmedit<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><%=i%>&nbsp;</td>
      <td><input name='ParCode' type='text' value="<%=ParCode%>" size='12' maxlength='18' readonly></td>
      <td><input name='ParName' type='text' value="<%=ParName%>" size='18' maxlength='24'></td>
      <td colspan="2"><input name='ParText' type='text' value="<%=ParText%>" size='24' maxlength='120' 
			onBlur="chkF_Blank(document.fmedit<%=i%>.ParText,'不能为空!')"></td>
      <td align="center" bgcolor="<%=col%>"><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="send" type="hidden" id="send" value="edt">
        <input <%=cDis%> type="submit" name="Submit" value="保存"></td>
      <td align="center" bgcolor="<%=col%>"><input <%=cDis%> type="button" name="Button" value="删除" 
			onClick="Del_YN('para_text.asp?send=del&ModID=<%=ModID%>&Page=<%=Page%>&ParCode=<%=Show_jsStr(ParCode)%>','<%=Show_jsStr(ParName)%>\n确认删除 ?')"></td>
    </tr>
  </form>
  <% 
rs.MoveNext 
if rs.eof then exit for 
     next 
   end if ' //////////////////////////////////////////
%>
  <tr>
    <td colspan="7" align="center" bgcolor="999999"></td>
  </tr>
  <form name="fm01" method="post" action="?">
    <tr bgcolor="f0f0f0">
      <td height="23" align="center">NO</td>
      <td ><input type='text' name='ParCode' size='12' maxlength='18'></td>
      <td><input type='text' name='ParName' size='18' maxlength='24'></td>
      <td><input name='ParText' type='text' size='24' maxlength='120' 
			    ></td>
      <td colspan="3"><input type="button" name="Button" value="增加" onClick="chkData()">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>"></td>
    </tr>
  </form>
</table>
<%
rs.Close()
SET rs=Nothing
%>
<div style="line-height:8px;">&nbsp;</div>
<%If ModID="ParConfig" Then%>
<center>
  <div style="width:640px; text-align:left; padding:8px; margin:3px auto; background-color:#F0F0F0; border:1px solid #CCC" align="center"><span class="col00F">注意</span>：此部分 [站点配置] 开发时已经设置好，请不要轻易修改，修改请到[系统与设置&gt;&gt;配置]中专门修改！</div>
</center>
<%ElseIf ModID="ParText" Or ModID="ParForm" Then%>
<table width="640" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><strong>随机系统 设置说明</strong></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"> 0. 本<strong>系统参数</strong>一般不需要修改，开发人员请<span class="colF0F">郑重修改</span>！<br>
      1. 请保证member目录下，文件名mu_read_123456.asp,mu_app_123456.asp<br>
      &nbsp;&nbsp;&nbsp;&nbsp;中123456分别与tAppRandom,tAppRand12参数一致; 修改后刷新即可.<br>
      2. 参考随机字符：<%=Rnd_ID("KEY",12)%> - <%=lCase(Rnd_ID("KEY",12))%> - <%=Rnd_ID("0",12)%><br>
      &nbsp;&nbsp;&nbsp;&nbsp;随机IDStr：<%=Get_IDRnd()%> (不重复:31通用随机码)<br>
      &nbsp;&nbsp;&nbsp;&nbsp;随机SN：  <span class="col00F"><%=Get_GUID("","PEACE"&Rnd_ID("0",1)&"ASP"&Rnd_ID("0",1)&"")%></span><br>
      3. 表单参数：<br>
      &nbsp;&nbsp;&nbsp;&nbsp;410~449参数值中,第1,第3位为字母m,d,h,n,s或数字,第2位为+或*字符,后面为3~18位随机字符<br>
      &nbsp;&nbsp;&nbsp;&nbsp;460~499参数值中,第1,第2位为字母m,d,h,n,s或数字,后面为3~18位随机字符</td>
  </tr>
</table>
<%
Else 'If ModID="ParText" OR ModID="ParLink" OR ModID="ParEmail" Then
%>
<table width="640" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><span class="col00F">注意</span>： 本参数表一般不需要修改，开发人员请郑重修改！<span class="colF0F">正常运行的系统，请不要随便修改！否则后果自负！！！</span></td>
  </tr>
</table>
<%
End If
%>
<script type="text/javascript">
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.ParCode.value.length==0) 
   {   
     alert(" 参数编码 不能为空！"); 
     document.fm01.ParCode.focus();
     eflag = 1; break;
   }
 if (document.fm01.ParName.value.length==0) 
   {   
     alert(" 参数名称 不能为空！"); 
     document.fm01.ParName.focus();
     eflag = 1; break;
   }
 if (document.fm01.ParText.value.length==0) 
   {   
     alert(" 参数值 不能为空！"); 
     document.fm01.ParText.focus();
     eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
<div style="line-height:8px;">&nbsp;</div>
</body>
</html>
