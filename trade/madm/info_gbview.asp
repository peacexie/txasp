<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<%

'Call Chk_URL()
Call Chk_Perm1("","") 
KeyID = RequestS("KeyID",3,48)
Act = Request("Act") 

If Act="Upd" Then

sql = " UPDATE [TradeGbook] SET " 
sql = sql& " InfType = '" & RequestS("InfType",3,255) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",C,255) &"'" 
sql = sql& ",InfCont = '" & RequestS("InfCont",C,500) &"'" 
sql = sql& ",InfReply = '" & RequestS("InfReply",C,500) &"'" 
sql = sql& ",SetShow = '" & RequestS("SetShow",3,2) &"'" 
sql = sql& ",LnkName = '" & RequestS("LnkName",3,48) &"'" 
sql = sql& ",LnkMobile = '" & RequestS("LnkMobile",3,48) &"'" 
sql = sql& ",LnkTel = '" & RequestS("LnkTel",3,48) &"'" 
sql = sql& ",LnkEmail = '" & RequestS("LnkEmail",3,255) &"'" 
sql = sql& ",LnkQQ = '" & RequestS("LnkQQ",3,48) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("MemID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='" & KeyID &"' "
Call rs_DoSql(conn,sql)

End If

  Set rs=Server.Createobject("Adodb.Recordset")
  rs.Open "SELECT * FROM [TradeGbook] WHERE KeyID='"&KeyID&"'",conn,1,1 
  If NOT rs.eof then 

KeyID = rs("KeyID")
KeyMod = rs("KeyMod") 
InfSubj = Show_Text(rs("InfSubj"))
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfType&"'")
xxxCont =  rs("InfCont")
xxxReply =  rs("InfReply") 
LnkName = Show_Text(rs("LnkName"))
LnkTel = rs("LnkTel")
LnkQQ = rs("LnkQQ")
LnkEmail = Show_Text(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")

  End If
  rs.Close()
  Set rs = Nothing
  
MD=KeyMod
ID=xxxReply
  
If MD="TraG124" Then
  MDName = "在线留言"
ElseIf MD="TraR124" Then
  MDName = "信息评论"
ElseIf MD="TraA124" Then
  MDName = "申请职位"
ElseIf MD="TraO124" Then
  MDName = "订购产品"
Else
  Response.End()
End If
  
%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<%If Act="Edit" Then%>
<div style="line-height:12px">&nbsp;</div>
<table width="700"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr>
      <td align="center" bgcolor="#FFFFFF">修改</td>
      <td bgcolor="#FFFFFF"><%=MDName%></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">主题</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" size="36" maxlength="60" value="<%=InfSubj%>">
        &nbsp;类别
        <select name="InfType" id="InfType" style="width:210px;">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='BisU124'",InfType)%>
        </select></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">内容<br />
        (250字内)<br />
        当前<span id="CntDiv1">0</span>字</td>
      <td bgcolor="#FFFFFF"><textarea name="InfCont" rows="4" style="width:450px" onBlur="CntCont(this,1)"><%=Show_Form(xxxCont)%></textarea></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">姓名</td>
      <td bgcolor="#FFFFFF"><input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="36" maxlength="24">
        &nbsp;邮件
        <input name="LnkEmail" type="text" id="LnkEmail" value="<%=LnkEmail%>" size="36" maxlength="120" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">电话</td>
      <td bgcolor="#FFFFFF"><input name="LnkTel" type="text" id="LnkTel" value="<%=LnkTel%>" size="36" maxlength="15">
        &nbsp;ＱＱ
        <input name="LnkQQ" type="text" id="LnkQQ" value="<%=LnkQQ%>" size="36" maxlength="15" /></td>
    </tr>
    <%If MD="TraG124" Then%>
    <tr>
      <td align="center" bgcolor="#FFFFFF">回复<br />
        (250字内)<br />
        当前<span id="CntDiv2">0</span>字</td>
      <td bgcolor="#FFFFFF"><textarea name="InfReply" rows="4" style="width:450px" onBlur="CntCont(this,2)"><%=Show_Form(xxxReply)%></textarea></td>
    </tr>
    <%End If%>
    <tr>
      <td width="12%" align="center" nowrap bgcolor="#FFFFFF">显示</td>
      <td nowrap bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="280" nowrap="nowrap"><select name="SetShow" id="SetShow" style="width:120px;">
                <option value="Y"  <%if SetShow="Y" then Response.Write("selected")%>>已审核</option>
                <option value="N"  <%if SetShow="N" then Response.Write("selected")%>>未审核</option>
              </select></td>
            <td align="right" nowrap="nowrap"><input type="button" name="Button" value="保存" onClick="chkData()" />
              &nbsp;&nbsp;
              <input type="reset" name="Submit2" value="重填" /></td>
            <td width="150"><input name="KeyID" type="hidden" id="KeyID" value="<%=KeyID%>" />
              <input name="Act" type="hidden" id="Act" value="Upd" /></td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<%End If%>
<%

If xxxReply="" Then xxxReply="<span class=cRed>(暂无回复)</span>"

If Len(LnkQQ)>=5 And isNumeric(LnkQQ) Then
  LnkQQ = "<span class='da01_qq'>&nbsp;</span><a href=tencent://message/?uin="&LnkQQ&"&Site="&MemSubj&"&Menu=yes>"&LnkQQ&"</a>"
Else
  If LnkQQ="" Then LnkQQ="00000"
  LnkQQ = "<span class=fntCCC>"&LnkQQ&"</span>"
End If

%>
<div style="line-height:12px">&nbsp;</div>
<table width="700" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="15%" align="center" valign="top" bgcolor="#FFFFFF">主题</td>
    <td width="40%" align="left" valign="top" bgcolor="#FFFFFF"><font color="#0000FF"><%=InfSubj%></font></td>
    <td width="15%" align="center" valign="top" bgcolor="#FFFFFF">时间</td>
    <td align="center" valign="top" bgcolor="#FFFFFF"><%=LogATime%></td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF">姓名</td>
    <td align="left" valign="top" bgcolor="#FFFFFF"><%=LnkName%></td>
    <td align="center" valign="top" bgcolor="#FFFFFF">电话</td>
    <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkTel%></td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF">类别</td>
    <td align="left" valign="top" bgcolor="#FFFFFF"><%=TypName%></td>
    <td align="center" valign="top" bgcolor="#FFFFFF">ＱＱ</td>
    <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkQQ%></td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF">IP</td>
    <td align="left" valign="top" bgcolor="#FFFFFF"><%=LogAddIP%></td>
    <td align="center" valign="top" bgcolor="#FFFFFF">邮件</td>
    <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkEmail%></td>
  </tr>
  <tr>
    <td colspan="4" align="left" valign="top" bgcolor="#FFFFFF" style="padding:5px;"><fieldset style="paddingt: 5px; border:#F0F0F0 1px solid">
        <legend class="fnt666">内容</legend>
        <table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr>
            <td class="SysCont" style="padding:3px;"><%=Show_Text(xxxCont)%></td>
          </tr>
        </table>
        <%If MD="TraR124" OR MD="TraO124" OR MD="TraA124" Then%>
        <div style="float:right; padding:5px 8px;"> 原文：<a href="../iview.asp?KeyID=<%=ID%>" class="cRed">查看</a> </div>
        <%End If%>
      </fieldset></td>
  </tr>
  <%If MD="TraG124" Then%>
  <tr>
    <td colspan="4" align="left" valign="top" bgcolor="#FFFFFF" style="padding:5px;"><fieldset style="paddingt: 5px; border:#F0F0F0 1px solid">
        <legend class="fnt666">回复</legend>
        <table width="100%" border="0" cellpadding="2" cellspacing="1">
          <tr>
            <td class="SysCont" style="padding:8px;"><%=Show_Text(xxxReply)%></td>
          </tr>
        </table>
      </fieldset></td>
  </tr>
  <%End If%>
</table>
<div style="line-height:12px">&nbsp;</div>
<script type="text/javascript">

var dfm = document.getElementById('fm01');

function CntCont(e,i)
{
	var n = e.value.length;
	getElmID('CntDiv'+i).innerHTML = n;
	if(n>250) { alert('内容超过250个字了'); }
}

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (dfm.InfSubj.value.length==0) 
   {   
     alert(" <%=vGbo_jsSubj%>"); 
     dfm.InfSubj.focus();
     eflag = 1; break;
   }
 if (dfm.InfCont.value.length==0) 
   {   
     alert(" <%=vGbo_jsCont%>"); 
	 dfm.InfCont.focus();
     eflag = 1; break;
   }
 if (dfm.InfCont.value.length>=250) 
   {   
     alert(" <%=vGbo_jsCMax%> 250!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ dfm.submit(); }
}</script>
</body>
</html>
