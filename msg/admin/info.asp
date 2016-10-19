<!--#include file="_config.asp"-->
<%
Dim res,act,msg
act = Request("act")
If act="" Then act="Balance"

If act="Charge" Then
  actName = "接口充值"
ElseIf act="doCharge" Then 
  Call chkReload()
  actName = "接口充值" 
  CrdID = RequestS("CrdID",3,72)
  CrdPW = RequestS("CrdPW",3,72)
  CrdMoney = RequestS("CrdMoney","N",0)
  CrdCount = RequestS("CrdCount","N",0) 
  CrdNote = RequestS("CrdNote",3,96)
  res = smsCharge(CrdID,CrdPW,CrdMoney,CrdCount,CrdNote)
  Call logCharge("Charge@System",CrdMoney,CrdCount,"充值"&res(2)&":("&CrdID&"/"&CrdPW&"):"&CrdNote)
  msg = res(2)&" "&Now()
ElseIf act="Balance" Then
  res = smsBalance()
  reB = rs_Val(conn,"SELECT SUM(MemBalance)AS MemCharge FROM [SmsMember]")
  reC = res(1) - reB
  actName = "接口余额"
  msg = res(2)&" "&Now()

Else
  If act="" Then
    msg = "无操作！ "
  Else
    msg = act
  End If
End If

Session("ChkCode") = Rnd_ID("",8)
'Response.Write "<br>"&smpFmtTime("2011-12-31 15:30:20")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>接口 Info</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css" />
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.rGap {
	text-align:right;
	padding-right:90px;
}
</style>
</head>
<body>
<div class="line15">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong><%=actName%></strong></td>
          <td align="right">&nbsp;</td>
          <td width="10%" align="left" nowrap="nowrap"><a href="?act=Balance">接口余额</a> | <a href="?act=Charge">接口充值</a> | <a href="xeucp.asp">接口激活</a><br />
          <a href="?act=Pass">修改密码</a> | <a href="?act=Receive">接收短信</a> | <a href="?act=Report">获取回执</a></td>
          <td width="10%" align="right" nowrap="nowrap">| <a href="debug.asp">接口调试</a><br />            
          | <a href="../out/smsapi.asp">外部调试</a></td>
        </tr>
      </table></td>
  </tr>
  <%If act="Balance" Then%>
  <tr>
    <td width="30%" align="right">系统余额(条)</td>
    <td class="rGap"><%=res(1)%> (条)</td>
  </tr>
  <tr>
    <td align="right">客户余额(条)</td>
    <td class="rGap"><%=reB%> (条)</td>
  </tr>
  <tr>
    <td align="right">可 充 值(条)</td>
    <td nowrap="nowrap" class="rGap"><%=res(1)%> - <%=reB%> = <%=reC%> (条)</td>
  </tr>
  <%End If%>
  <%If act="Charge" Then%>
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td width="30%" align="right">充值卡帐号</td>
      <td align="left"><input name="CrdID" type="text" id="CrdID" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td align="right">充值卡密码</td>
      <td align="left"><input name="CrdPW" type="text" id="CrdPW" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td align="right">短信条数(条)</td>
      <td align="left"><input name="CrdCount" type="text" id="CrdCount" size="24" maxlength="5" />
      (10万条以下)</td>
    </tr>
    <tr>
      <td align="right">折合金额(元)</td>
      <td align="left"><input name="CrdMoney" type="text" id="CrdMoney" size="24" maxlength="4" />
        (1万元以下)</td>
    </tr>
    <tr>
      <td align="right">说明(48字内)</td>
      <td align="left"><input name="CrdNote" type="text" id="CrdNote" size="24" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="right"><input name="act" type="hidden" id="act" value="doCharge" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Button" type="button" class="btn05" onclick="chkData()" value="接口充值" <%=chkDTabs("Charge")%> />
        <input type="reset" name="Button3" value="重填资料" /></td>
    </tr>

  </form>
  <%End If%>
  
  <%If act="Pass" Then%>
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td align="right">旧密码</td>
      <td align="left"><input name="SetPass" type="text" id="SetPass" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td width="30%" align="right">新密码</td>
      <td align="left"><input name="SetPNew" type="text" id="SetPNew" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td align="right">扩展(Ext)</td>
      <td align="left"><input name="SetExt" type="text" id="SetExt" size="24" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="right"><input name="act" type="hidden" id="act" value="doPass" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Submit" type="submit" class="btn05" xxxonclick="chkData()" value="修改密码" <%=chkDTabs("Pass")%> />
        <input type="reset" name="Button3" value="重填资料" /></td>
    </tr>
  </form>
  <%End If%>
  <%
  If act="doPass" Then
  SetPass = RequestS("SetPass",3,96)
  SetPNew = RequestS("SetPNew",3,96)
  SetExt = RequestS("SetExt",3,96)
  res = smsPass(SetPass,SetPNew,SetExt)
  %>
    <tr>
      <td align="right"> 修改密码 结果状态</td>
      <td align="left"><%=res(0)%>(<%=res(1)%>)</td>
    </tr>
    <tr>
      <td width="30%" align="right">修改密码 结果描述</td>
      <td align="left"><%=res(2)%></td>
    </tr>
  <%End If%>
  
  <%If act="Report" Then%>
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td align="right">最大ID(MaxID)</td>
      <td align="left"><input name="SetMax" type="text" id="SetMax" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td width="30%" align="right">返回条数(Count)</td>
      <td align="left"><input name="SetCount" type="text" id="SetCount" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td align="right">扩展(Ext)</td>
      <td align="left"><input name="SetExt" type="text" id="SetExt" size="24" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="right"><input name="act" type="hidden" id="act" value="doReport" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Submit" type="submit" class="btn05" xxxonclick="chkData()" value="获取回执" <%=chkDTabs("Report")%> />
      <input type="reset" name="Button3" value="重填资料" /></td>
    </tr>
  </form>
  <%End If%>
  <%
  If act="doReport" Then
  SetMax = RequestS("SetMax",3,96)
  SetCount = RequestS("SetCount",3,96)
  SetExt = RequestS("SetExt",3,96)
  res = smsReport(SetMax,SetCount,SetExt)
  %>
    <tr>
      <td align="right">获取回执 结果状态</td>
      <td align="left"><%=res(0)%>(<%=res(1)%>)</td>
    </tr>
    <tr>
      <td width="30%" align="right">获取回执 结果描述</td>
      <td align="left"><%=res(2)%></td>
    </tr>
  <%End If%>
  
  <%If act="Receive" Then%>
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td align="right">接收子端口(SubID)</td>
      <td align="left"><input name="SetSub" type="text" id="SetSub" size="24" maxlength="36" /></td>
    </tr>
    <tr>
      <td align="right">扩展(Ext)</td>
      <td align="left"><input name="SetExt" type="text" id="SetExt" size="24" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="right"><input name="act" type="hidden" id="act" value="doReceive" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Submit" type="submit" class="btn05" xxxonclick="chkData()" value="接收短信" <%=chkDTabs("Receive")%> />
      <input type="reset" name="Button3" value="重填资料" /></td>
    </tr>
  </form>
  <%End If%>
  <%
  If act="doReceive" Then
  SetSub = RequestS("SetSub",3,96)
  SetExt = RequestS("SetExt",3,96)
  res = smsReceive(SetSub,SetExt)
  %>
    <tr>
      <td align="right">接收短信 结果状态</td>
      <td align="left"><%=res(0)%>(<%=res(1)%>)</td>
    </tr>
    <tr>
      <td width="30%" align="right">接收短信 结果描述</td>
      <td align="left"><%=res(2)%></td>
    </tr>
  <%End If%>
  

  <tr>
    <td colspan="2" align="left">说明/提示：<span class="fntF00"><%=msg%></span><br />
      <span class="fnt00F">第一次使用,请设置好config/<%=smcSFile%>文件的smcUser,smcPass等相关参数</span>;</td>
  </tr>
  
  <form name="fmDir" method="post" action="" Xtarget="_blank">
  <tr bgcolor="#999999">
    <td colspan="2" align="left" bgcolor="#F0F0F0"><input name="sDir" type="text" id="sDir" value="xeucp.asp" size="45" maxlength="255">
      <input type="button" name="button" id="button" value="接口测试" onClick="goType()">
      &nbsp;</td>
  </tr>
  </form>
  
  <tr>
    <td colspan="2" align="left">短信会员 - 会员增加 - 认证只在Member表中<br />
短信会员 - 清理会员</td>
  </tr>
</table>

<script type="text/javascript">

function goType()
{
  var fm = document.fmDir;
  fm.action = fm.sDir.value;
  fm.submit();
}

var fm = document.fm01;
function chkData(){   
  eflag=1
  for(i=0;i<1;i++){  
   if (fm.CrdID.value.length==0){ 
     alert('[错误]\n 请输入 充值卡帐号!');
     fm.CrdID.focus();
     eflag=0; break;
   }
   if (fm.CrdPW.value.length==0){ 
     alert('[错误]\n 请输入 充值卡密码!');
     fm.CrdPW.focus();
     eflag=0; break;
   }
   if (fm.CrdCount.value.length==0){ 
     alert('[错误]\n 请输入 短信条数(条)!');
     fm.CrdCount.focus();
     eflag=0; break;
   }
   var tmp = chkF_Int(fm.CrdCount,"短信条数(条) 为整数!");
   if (tmp=="ER"){   
     eflag=0; break;
   }
   if (fm.CrdMoney.value.length==0){ 
     alert('[错误]\n 请输入 折合金额(元)!');
     fm.CrdMoney.focus();
     eflag=0; break;
   }
   var tmp = chkF_Dot(fm.CrdMoney,"折合金额(元) 为数字!");
   if (tmp=="ER"){   
     eflag=0; break;
   }
  }
  if (eflag==1){ fm.submit(); }
}
</script>
</body>
</html>
