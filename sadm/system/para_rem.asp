<!--#include file="config.asp"-->
<!--#include file="../../upfile/sys/para/rkeyid.asp"-->
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
'Call rs_DoSql(conn,"UPDATE AdmPara SET ParRem=replace(ParRem,'xxxyyyzzz','aaabbbccc')")

sCode = Replace(Replace(sEditParCode," ",""),",",";")
sFlag = Replace(Replace(sEditParFlag," ",""),",",";")
sMod  = Replace(Replace(sEditParMod," ",""),",",";")

send = Request("send")
Page = RequestS("Page","N",1) 
ModID = RequestS("ModID","C",24) 
KW = RequestS("KW","C",24) 
If ModID="" Then
 ModID = "ParRem"
End If
ModName = rs_Val("","SELECT [SysName] FROM AdmSyst WHERE [SysID]='"&ModID&"'")
ModName = Replace(ModName,".rem","")
If send = "ins" then
ParCode = RequestS("ParCode",3,24)
sql = " INSERT INTO [AdmPara] (" 
sql = sql& "  ParName" 
sql = sql& ", ParCode" 
sql = sql& ", ParFlag" 
sql = sql& ", ParRem" 
sql = sql& ",LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & RequestS("ParName",3,48) &"'" 
sql = sql& ", '" & ParCode &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & RequestS("ParRem",3,480123) &"'" 
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
'ParRem = Replace(ParRem,"[/textarea]","</textarea>")
sql = " UPDATE [AdmPara] SET " 
sql = sql& " ParName = '" & RequestS("ParName",3,48) &"'" 
sql = sql& ",ParRem = '" & RequestS("ParRem",3,480123) &"' " 
sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
sql = sql& " WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 修改成功!"
ElseIf send = "del" Then
  ParCode = RequestS("ParCode",3,24)
  sql = "DELETE FROM [AdmPara] WHERE ParCode='"&ParCode&"' "
  Call rs_DoSql(conn,sql)
  Call fil_del("../../upfile/sys/para/"&ParCode)
  msg = ""&ParCode&" 删除成功!"
End If

sql = "SELECT "
sql = sql & " ParName "
sql = sql & ",ParCode "
sql = sql & ",ParFlag "
sql = sql & ",ParRem "
sql = sql & " FROM [AdmPara] "
sql = sql & " WHERE ParFlag='" & ModID &"' "
If KW<>"" Then
sql = sql & " AND ParCode LIKE '%" & KW &"%' "
End If
sql = sql & " ORDER BY ParCode "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
rs.PageSize = 240 
%>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="e0e0e0">
  <tr align="center" bgcolor="f8f8f8">
    <td colspan="3"><table width="100%"  border="0" cellspacing="1" cellpadding="1">
        <tr>
        <form name="fmSch" method="post" action="?">
          <td width="30%" align="center"><input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
<input type="submit" name="Submit2" value="搜索">
<input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
</form>
          <td rowspan="2" align="right"><!--#include file="para_menu.asp"-->
          </td>
        </tr>
        <tr>
          <td align="center"><strong><%=ModName%> 设置</strong> <font color="#FF0000"><%=msg%></font></td>
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
   col = "FFFFFF"
else
   col = "F4F4F4"
end if
ParName = Show_Form(rs("ParName"))
ParCode = Show_Form(rs("ParCode"))
ParRem = Show_Form(rs("ParRem")) 'Show_Form()
'ParRem = Replace(ParRem,"</textarea>","[/textarea]")
If ModID="ParFil" Then
  'ParRem = Server.HTMLEncode(ParRem)
  '// ...请注意，不要提交任何违反国家规定的内容！...
  wrap = "wrap='ON'"
ElseIf Right(ParCode,4)=".asp" Then 
  wrap = "wrap='OFF'"

Else
  wrap = "wrap='OFF'"
End If
flag1 = Get_SOpt(sCode,sFlag,ParCode,"Val")
flag2 = Get_SOpt(sCode,sFlag,ModID,"Val")
%>
  <form name="fmedit<%=i%>" method="post" action="para_rem.asp">
    <tr bgcolor="<%=col%>">
      <td align="center" nowrap bgcolor="<%=col%>">&nbsp;参数编码<br>
        <input name='ParCode' type='text' value="<%=ParCode%>" size='10' maxlength='12' readonly>
        <br>
&nbsp;参数名称<br>
        <input name='ParName' type='text' value="<%=ParName%>" size='10' maxlength='24'>
      </td>
      <td bgcolor="<%=col%>">
        <textarea name="ParRem" cols="80" rows="5" <%=wrap%> ><%=ParRem%></textarea>
      <div> 大窗口模式编辑：&nbsp;
        1.<a href="../../smod/adupd/para_set1.asp?ID=<%=ParCode%>&Flag=Rem&nLen=18&fRet=Close" target="_blank">文本(Text代码)编辑</a>
      &nbsp;
      <%If inStr(flag1&flag2,"Html")>0 Then%>
        2.<a href="../../smod/adupd/para_set1.asp?ID=<%=ParCode%>&Flag=Editor&nLen=18&fRet=List" target="_blank">可视(FCK插件)编辑</a>
	  <%End If%>
      </div>
      </td>
      <td nowrap bgcolor="<%=col%>">
      <input type="submit" name="Submit" value="保存" >
        <br>
        <input type="button" name="Button" value="删除" 
			onClick="Del_YN('para_rem.asp?send=del&ParCode=<%=Show_jsStr(ParCode)%>&ModID=<%=ModID%>&Page=<%=Page%>&KW=<%=KW%>','<%=Show_jsStr(ParName)%>\n确认删除 ?')">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
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
  <tr>
    <td colspan="3" align="center" bgcolor="999999"></td>
  </tr>
  <form name="fm01" method="post" action="para_rem.asp">
    <tr bgcolor="f0f0f0">
      <td height="23" align="center" nowrap bgcolor="f0f0f0">&nbsp;参数编码<br>
        <input type='text' name='ParCode' size='10' maxlength='12'>
        <br>
        参数名称<br>
        <input type='text' name='ParName' size='10' maxlength='24'>
      </td>
      <td bgcolor="f0f0f0"><textarea name="ParRem" cols="80" rows="5"></textarea>
      </td>
      <td nowrap bgcolor="f0f0f0"><input type="button" name="Button" value="增加" onClick="chkData()">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
<input name="Page" type="hidden" id="Page" value="<%=Page%>"></td>
    </tr>
  </form>
</table>
<%
rs.Close()
SET rs=Nothing
%>
<script type="text/javascript">

function prview(ParCode)
{
  window.open('para_view.asp?ParCode='+ParCode);
}
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
 if (document.fm01.ParRem.value.length>=4000) 
   {   
     alert(" 参数值 内容太长！"); 
     document.fm01.ParRem.focus();
     eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
