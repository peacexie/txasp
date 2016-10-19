<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

send = Request("send")
Page = RequestS("Page","N",1) 
KeyID = RequestS("KeyID",3,48)

MD = RequestS("MD","C",24) 
MN = RequestS("MN","C",24)
If MD="" Then
 MD = "PayType"
 MN = "付款方式"
End If
If MD="PayType" Then
 pShow = "block"
Else
 pShow = "none"
End If

If send = "ins" then
sql = " INSERT INTO [OrdPara] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfNum" 
sql = sql& ", InfTop"
sql = sql& ", InfState"
sql = sql& ",LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & Get_AutoID(24) &"'"
sql = sql& ", '" & MD &"'" 
sql = sql& ", '" & RequestS("InfSubj",3,48) &"'" 
sql = sql& ", '" & RequestS("InfCont",3,512) &"'" 
sql = sql& ", " & RequestS("InfNum","N",0) &"" 
sql = sql& ", " & RequestS("InfTop","N",0) &"" 
sql = sql& ", '" & RequestS("InfState","C",12) &"'" 
sql = sql& " ,'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"'"
sql = sql& ")"
  Call rs_DoSql(conn,sql)
  msg = "增加成功!"
ElseIf send = "edt" Then
sql = " UPDATE [OrdPara] SET " 
sql = sql& " InfSubj = '" & RequestS("InfSubj",3,48) &"'" 
sql = sql& ",InfCont = '" & RequestS("InfCont",3,512) &"' " 
sql = sql& ",InfNum = " & RequestS("InfNum","N",0) &" " 
sql = sql& ",InfTop = " & RequestS("InfTop","N",0) &" " 
sql = sql& ",InfState = '" & RequestS("InfState",3,512) &"' " 
sql = sql& ",InfPic = '" & RequestS("InfPic",3,255) &"' "
sql = sql& ",LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
sql = sql& " WHERE KeyID='"&KeyID&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 修改成功!"
ElseIf send = "del" Then
  sql = "DELETE FROM [OrdPara] WHERE KeyID='"&KeyID&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ParCode&" 删除成功!"
End If

sql = "SELECT "
sql = sql & " * FROM [OrdPara] "
sql = sql & " WHERE KeyMod='" & MD &"' "
sql = sql & " ORDER BY InfTop "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
rs.PageSize = 240 
%>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="e0e0e0">
  <tr align="center" bgcolor="f8f8f8">
    <td colspan="6"><table width="100%"  border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td width="30%" align="center"><strong><%=MN%> 设置</strong><br>           
		   <font color="#FF0000"><%=msg%>&nbsp;</font></td>
          <td align="right">
          <a href="?MD=PayType&MN=付款方式">付款方式</a> - <a href="?MD=SendType&MN=配送方式">配送方式</a> - <a href="?MD=SendTime&MN=时间要求">时间要求</a>&nbsp;&nbsp; </td>
        </tr>

      </table></td>
  </tr>
  <tr bgcolor="e0e0e0">
    <td height="23" align="center" bgcolor="e0e0e0">名称</td>
    <td bgcolor="e0e0e0">&nbsp;内容 </td>
    <td nowrap bgcolor="e0e0e0">附加金额</td>
    <td nowrap bgcolor="e0e0e0">顺序</td>
    <td nowrap bgcolor="e0e0e0">状态</td>
    <td nowrap bgcolor="e0e0e0">操作</td>
  </tr>
  <tr>
    <td colspan="6" align="center" bgcolor="999999"></td>
  </tr>
  <%
		  if NOT rs.EOF then
rs.AbsolutePage = 1 
     for i=1 to rs.PageSize 
if i mod 2 = 1 then
   col = "F8F8F8"
else
   col = "FFFFF8"
end if

KeyID = rs("KeyID")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = Show_Form(rs("InfCont"))

InfNum = rs("InfNum")
InfTop = rs("InfTop")
InfState = rs("InfState")
InfPic = rs("InfPic")

%>
  <form name="fmedit<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="center" nowrap bgcolor="<%=col%>"><input name='InfSubj' type='text' id="InfSubj" value="<%=InfSubj%>" size='10' maxlength='12'>
      <img src="<%=InfPic%>" width="100" style="display:<%=pShow%>;"></img>
      </td>
      <td bgcolor="<%=col%>"><textarea name="InfCont" cols="56" rows="3" id="InfCont" 
			  onBlur="chkF_Len(document.fmedit<%=i%>.InfCont,256,'内容太长!')" wrap='OFF'><%=InfCont%></textarea>
      <%If MD="PayType" Then%>
      <br>
      <input name='InfPic' type='text' id="InfPic" value="<%=InfPic%>" size='50' maxlength='120'>
      <input name=view2 type=button id="Button1" value="选择" onClick="getRetObject(<%=i%>);window.open('../file/file_view.asp?yPath=myfile/bank/')">
<%End If%>
      </td>
      <td nowrap bgcolor="<%=col%>"><input name='InfNum' type='text' id="InfNum" value="<%=InfNum%>" size='4' maxlength='6'></td>
      <td nowrap bgcolor="<%=col%>"><input name='InfTop' type='text' id="InfTop" value="<%=InfTop%>" size='2' maxlength='4'></td>
      <td nowrap bgcolor="<%=col%>"><input name='InfState' type='text' id="InfState" value="<%=InfState%>" size='2' maxlength='4'></td>
      <td nowrap bgcolor="<%=col%>"><input type="submit" name="Submit" value="保存">
        <input name="KeyID" type="hidden" id="KeyID" value="<%=KeyID%>">
        <br>
        <input type="button" name="Button" value="删除" 
			onClick="Del_YN('?send=del&KeyID=<%=KeyID%>&MD=<%=MD%>&MN=<%=MN%>&Page=<%=Page%>','<%=Show_jsStr(InfSubj)%>')">
        <input name="send" type="hidden" id="send" value="edt">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="MD" type="hidden" id="MD" value="<%=MD%>">
        <input name="MN" type="hidden" id="MN" value="<%=MN%>"></td>
    </tr>
  </form>
  <% 
rs.MoveNext 
if rs.eof then exit for 
     next 
   end if ' //////////////////////////////////////////
%>
  <tr>
    <td colspan="6" align="center" bgcolor="999999"></td>
  </tr>
  <form name="fm01" method="post" action="?">
    <tr bgcolor="f0f0f0">
      <td height="23" align="center" nowrap bgcolor="f0f0f0"><input name='InfSubj' type='text' id="InfSubj" size='10' maxlength='12'></td>
      <td bgcolor="f0f0f0"><textarea name="InfCont" cols="56" rows="3" id="InfCont"></textarea>
      </td>
      <td nowrap bgcolor="f0f0f0"><input name='InfNum' type='text' id="InfNum" value="0" size='4' maxlength='6'></td>
      <td nowrap bgcolor="f0f0f0"><input name='InfTop' type='text' id="InfTop" value="666" size='2' maxlength='4'></td>
      <td nowrap bgcolor="f0f0f0"><input name='InfState' type='text' id="InfState" value="-" size='2' maxlength='4'></td>
      <td nowrap bgcolor="f0f0f0"><input type="button" name="Button" value="增加" onClick="chkData()">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="Page2" type="hidden" id="Page2" value="<%=Page%>">
        <input name="MD" type="hidden" id="MD" value="<%=MD%>">
        <input name="MN" type="hidden" id="MN" value="<%=MN%>"></td>
    </tr>
  </form>
</table>
<%
rs.Close()
SET rs=Nothing
%>

<script type="text/javascript">

var yFile;
function getRetObject(i)
{
	yFile = eval("document.fmedit"+i+".InfPic");
}

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 名称 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length==0) 
   {   
     alert(" 内容 不能为空！"); 
     document.fm01.InfCont.focus();
     eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
