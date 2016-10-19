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

send = Request("send")
Page = RequestS("Page","N",1) 
If send = "ins" then
ParCode = RequestS("ParCode",3,24)
sql = " INSERT INTO [AdmPara] (" 
sql = sql& "  ParName" 
sql = sql& ", ParCode" 
sql = sql& ", ParFlag" 
sql = sql& ", ParDate" 
sql = sql& ",LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & RequestS("ParName",3,48) &"'" 
sql = sql& ", '" & ParCode &"'" 
sql = sql& ", 'ParDate'" 
sql = sql& ", '" & RequestS("ParDate","D","1900-12-31") &"'" 
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
sql = sql& ",ParDate = '" & RequestS("ParDate","D","1900-12-31") &"' " 
sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
sql = sql& " WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 修改成功!"
ElseIf send = "del" Then
  ParCode = RequestS("ParCode",3,24)
  sql = "DELETE FROM [AdmPara] WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 删除成功!"
End If

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParDate "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE ParFlag='ParDate' "
sql = sql & " ORDER BY ParCode "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
rs.PageSize = 240 
%>
        <br>
        <table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="e0e0e0">
          <tr align="center" bgcolor="f8f8f8">
            <td colspan="7"><table width="100%"  border="0" cellspacing="1" cellpadding="1">
                <tr>
                  <td width="20%" align="center" nowrap><strong>日期参数</strong><br>
                  <font color="#FF0000"><%=msg%></font>                  </td>
                  <td align="right" nowrap><!--#include file="para_menu.asp"--></td>
                </tr>
              </table></td>
          </tr>
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
ParDate = Show_Form(rs("ParDate"))

%>
          <form name="fmedit<%=i%>" method="post" action="para_date.asp">
		  <tr bgcolor="<%=col%>">
            <td align="right"><%=i%>&nbsp;</td>
            <td><input name='ParCode' type='text' value="<%=ParCode%>" size='12' maxlength='12' readonly></td>
            <td><input name='ParName' type='text' value="<%=ParName%>" size='18' maxlength='24'></td>
            <td colspan="2"><input name='ParDate' type='text' value="<%=ParDate%>" size='12' maxlength='10' 
			onBlur="chkF_Date(document.fmedit<%=i%>.ParDate,'日期错误!')"></td>
            <td align="center" bgcolor="<%=col%>"><input name="send" type="hidden" id="send" value="edt">
            <input type="submit" name="Submit" value="保存"></td>
            <td align="center" bgcolor="<%=col%>"><input type="button" name="Button" value="删除" 
			onClick="Del_YN('para_date.asp?send=del&ParCode=<%=Show_jsStr(ParCode)%>','<%=Show_jsStr(ParName)%>\n确认删除 ?')"></td>
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
          <form name="fm01" method="post" action="para_date.asp">
            <tr bgcolor="f0f0f0">
              <td height="23" align="center">NO</td>
              <td ><input type='text' name='ParCode' size='12' maxlength='12'></td>
              <td><input type='text' name='ParName' size='18' maxlength='24'></td>
              <td><input name='ParDate' type='text' value="<%=Date()%>" size='12' maxlength='10' 
			    ></td>
              <td colspan="3">                <input type="button" name="Button" value="增加" onClick="chkData()">
              <input name="send" type="hidden" id="send" value="ins"> </td>
            </tr>
          </form>
        </table>
        <%
rs.Close()
SET rs=Nothing
%>
<div style="line-height:8px;">&nbsp;</div>
<table width="640" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
          <tr>
            <td align="center" valign="top" bgcolor="#FFFFFF"><span class="col00F">注意</span>： 本参数表一般不需要修改，开发人员请郑重修改！<span class="colF0F">正常运行的系统，请不要随便修改！否则后果自负！！！</span></td>
  </tr>
</table>
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
 tmv = chkF_Date(document.fm01.ParDate,"");
 if (tmv=='ER') 
   {   
     alert(" 日期错误!"); 
     document.fm01.ParDate.focus();
     eflag = 1; break;
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
        
</body>
</html>
