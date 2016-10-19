<!--#include file="smscfg.asp"-->
<%

tm = outFmtTime(Now()) 'Now()
sn = outEncSN(id,tm,snOrg) 
urlAdd0 = smaAdda '(...&"msg/user/sapi.asp")
urlPar0 = "?tm="&tm&"&id="&id&"&sn="&sn&"" 'Peace_Sms_RndID="&Timer&"&
'以上公共参数


' //////////////////////////////// acs,acu:直接发送Start
ctu = "" 
acs = Request("acs") 'act_send发送方式
acu = Request("acu") 'act_user自定义
    If acs&acu<>"" Then
If smaState<>"isOpen" Then 
  Response.End()
End If
If inStr("(SOrder,SMember)",acu)>0 Then
  acs = "SForm" 'SFrame
  If acu="SOrder" Then
    OrdID = Request("OrdID")
	ctu = "A new order(No."&OrdID&") added in website."
  ElseIf acu="SMember" Then
    MemID = Request("MemID")
	ctu = "A new member(ID."&MemID&") added in website."
  End If
End If
If inStr("(SForm,SFrame)",acs)>0 Then
  'ChkPerm '注意, 请设置好访问权限
  ct1 = Request("ct1")
  If ct1="" And ctu<>"" Then
    ct1 = ctu
  ElseIf ct1="" Then
    '???
  End If
  tel = Request("tel")
  If tel="" Then
    tel = smaTels
  End If
  tel = Replace(Replace(tel,"-","")," ","")
  cs = Request("cs") 'outEncode(xStr)
  If acs="SFrame" Then
    If cs<>"" Then 
	  ct1 = outEncode(ct1)
	  csp = "&cs="&cs&""
	Else
	  ct1 = Server.URLEncode(ct1)
	  csp = ""
	End If
	urlPara = urlPar0&"&act=Send&tel="&tel&"&ct1="&ct1&csp&"" 
	'echo tel&ct1 :Response.End()
	Response.Redirect urlAdd0&urlPara
  Else
    If cs<>"" Then 
	  ct1 = outEncode(ct1)
	  csf = "  <input type='hidden' value='"&cs&"' name='cs' />"
	Else
	  csf = ""
	End If
	sSend =        "<form name='fmSms' action='"&urlAdd0&"' method='post' target='_self'>"
    sSend = sSend& "  <input type='hidden' value='"&tm&"' name='tm' />"
    sSend = sSend& "  <input type='hidden' value='"&id&"' name='id' />"
    sSend = sSend& "  <input type='hidden' value='"&sn&"' name='sn' />"
    sSend = sSend& "  <input type='hidden' value='Send' name='act' />"
    sSend = sSend& "  <input type='hidden' value='"&tel&"' name='tel' />"
    sSend = sSend& "  <input type='hidden' value='"&ct1&"' name='ct1' />"
	sSend = sSend& csf
    sSend = sSend& "</form>"
    sSend = sSend& "<script type='text/javascript'>"
    sSend = sSend& "document.fmSms.submit();"
    sSend = sSend& "</script>"
	'echo tel&ct1 :Response.End()
	Response.Write sSend
  End If
End If
Response.End()
    End If
' //////////////////////////////// acs,acu:直接发送End


ad = Request("ad") :If ad="" Then ad="Balance"
ac = Request("ac") :If ac="" Then ac="OutAPI" 'OutAPI
If smaState<>"isOpen" Then ac="Stop"


If ac="OutFrame" Then
  urlPar0 = urlPar0&"&act=SendOut"
ElseIf ad="Balance" Then
  urlPar0 = urlPar0&"&act=Balance&cs=(cs)&re=Xml"
ElseIf ad="Send" Then
  urlPar0 = urlPar0&"&act=Send&tel=&ct1=&cs=(cs)&re=Xml"
ElseIf ad="Edit" Then
  urlPar0 = urlPar0&"&nsn=(nsn)&cs=(cs)&re=Xml"
ElseIf ad="Read" Then
  urlPar0 = urlPar0&"&act=Readme&cs=(cs)&re=Xml"
End If
urlPar1 = Replace(Replace(urlPar0,sn,"(sn)"),tm,"(now)")


tel = "198-1234-5678 [Xie永顺], 198-1234-5648 [Xie永顺]" '用于测试
ct1 = ""&Time()&"测试@"&Get_CIP()&"("&Session("UsrID")&")" '用于测试
nsn = Get_GUID("","PEACE0ASP0")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>发送短信</title>
<link rel="stylesheet" type="text/css" href="<%=smsABase%>msg/inc/spub.css"/>
<style type="text/css">
.note {
	line-height:120%;
	padding-bottom:5px;
}
.not2 {
	line-height:108%;
	padding-bottom:5px;
}
.pTitle { 
    padding-left:50px;
}
</style>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
</head>
<body>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tdbg">
  <tr>
    <td align="left"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="20%" align="center"><strong>[<%=ac%>]</strong></td>
          <td align="center" nowrap="nowrap"><a href="?ac=OutAPI">外部API调用</a> | <a href="?ac=OutFrame">外部iFrame调用</a> | <a href="?ac=OutRead">外部调用说明</a> | <a href="?ac=OutPara">外部参数说明</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<%If ac="Stop" Then%>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tdbg">
  <tr>
    <td colspan="2" align="left"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong>禁用说明</strong></td>
          <td width="10%" align="right">&nbsp;</td>
          <td align="center" nowrap="nowrap">禁用说明</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="80" colspan="2" align="left" valign="top">&nbsp;&nbsp;当前设置：短信接口API 不能使用！</td>
  </tr>
</table>
<%End If%>
<%If ac="OutAPI" Then%>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <form action="<%=smsABase%>msg/user/sapi.asp" method="post" name="fm01" id="fm01">
    <tr>
      <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
          <tr>
            <td width="30%" align="center"><strong>外部API调用</strong></td>
            <td width="10%" align="right">&nbsp;</td>
            <td align="center" nowrap="nowrap"><a href="?ad=Balance">查询余额</a> | <a href="?ad=Send">发送短信</a> | <a href="?ad=Edit">修改SN</a> | <a href="?ad=Read">返回说明</a></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td colspan="2" align="left">调用地址：<%=urlAdd0%><br />
        调用参数：<%=urlPar1%></td>
    </tr>
    <tr>
      <td align="center">tm</td>
      <td><input name="tm" type="text" id="tNumb5" value="<%=tm%>" size="54" maxlength="48" /></td>
    </tr>
    <tr>
      <td width="15%" align="center">id</td>
      <td><input name="id" type="text" id="id" value="<%=id%>" size="54" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="center">sn</td>
      <td><input name="sn" type="text" id="sn" value="<%=sn%>" size="54" maxlength="48" /></td>
    </tr>
    <%If ad="Send" Then%>
    <tr>
      <td width="15%" align="center">tel<br />
        手机号码</td>
      <td align="left"><input name="tel" type="text" id="tel" value="<%=tel%>" size="54" maxlength="48" />
        <br />
        最多100个号码, 用半角[<span class="fntF00 fB f14px">,</span>]分开</td>
    </tr>
    <tr>
      <td align="center">ct1<br />
        短信内容<br />
        70字内</td>
      <td align="left"><textarea name="ctn" cols="36" rows="4" id="ctn" ><%=ct1%></textarea></td>
    </tr>
    <%End If%>
    <%If ad="Edit" Then%>
    <tr>
      <td align="center">nsn</td>
      <td><input name="nsn" type="text" id="nsn" value="<%=nsn%>" size="54" maxlength="48" /></td>
    </tr>
    <%End If%>
    <%If Request("reAct")="Return" Then%>
    <tr>
      <td align="center">reAct</td>
      <td>Return</td>
    </tr>
    <tr>
      <td align="center">reState</td>
      <td><%=Request("reState")%></td>
    </tr>
    <tr>
      <td align="center">reInt</td>
      <td><%=Request("reInt")%></td>
    </tr>
    <tr>
      <td align="center">reStr</td>
      <td><%=Request("reInfo")%></td>
    </tr>
    <tr>
      <td align="center">reInfo</td>
      <td><%=outDecode(Request("reInfo"))%></td>
    </tr>
    <%End If%>
    <tr>
      <td align="center">act</td>
      <td><select name="act" id="act">
          <%If ad="Read" Then%>
          <option value="Readme">Readme</option>
          <%ElseIf ad="Balance" Then%>
          <option value="Balance">Balance</option>
          <%ElseIf ad="Send" Then%>
          <option value="Send">Send</option>
          <%Else%>
          <option value="">(Null)</option>
          <%End If%>
        </select>
        &nbsp;&nbsp;cs
        <select name="cs" id="cs">
          <option value="UTF-8" selected="selected">UTF-8</option>
          <option value="GB2312">GB2312</option>
        </select>
        &nbsp;
        re
        <select name="re" id="re">
          <option value="Xml">Xml</option>
          <option value="Report" selected="selected">Report</option>
          <option value="goBack">goBack</option>
        </select></td>
    </tr>
    <tr>
      <td align="right">&nbsp;</td>
      <td align="left"><input name="Button" type="button" class="btn60" onclick="chkOutSend()" value="提交" />
        <input name="Button2" type="reset" class="btn60" value="重填" />
        <input name="ct1" type="hidden" id="ct1" value="" />
        <a href="?ad=<%=ad%>&amp;<%=Rnd_ID("",8)%>=<%=Rnd_ID("",8)%>">重载</a></td>
    </tr>
  </form>
</table>
<script type="text/javascript">
var fm = document.fm01;
function chkOutSend(){   
  var eflag=1, ctv="";
  <%If ad="Send" Then%>
  for(i=0;i<1;i++){  
   if (fm.tel.value.length==0){ 
     alert('[错误]\n 请输入 手机号码!');
     fm.tel.focus();
     eflag=0; break;
   }
   if (fm.ctn.value.length==0){ 
     alert('[错误]\n 请输入 短信内容!');
     fm.ctn.focus();
     eflag=0; break;
   }
   if (fm.ctn.value.length>70){ 
     alert('[错误]\n 短信内容 请少于70字!');
     fm.ctn.focus();
     eflag=0; break;
   }
  }
  var ctv = fm.ctn.value; //保存信息
  if((fm.cs.value!="UTF-8") &&(eflag==1))
    ctv = outEncode(ctv);  //编码信息
  <%End If%>
  if (eflag==1){ 
    fm.ct1.value = ctv; //传递信息
	fm.submit(); 
  }
}

function outEncode(xStr){ 
  var n = xStr.length;
  var s = "";
  for(i=0;i<n;i++){	
    s += xStr.charCodeAt(i) + ";";
  } 
  return s
}

</script>
<%End If%>
<%If ac="OutFrame" Then%>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td align="left"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong>外部iFrame调用</strong></td>
          <td width="10%" align="right">&nbsp;</td>
          <td align="center" nowrap="nowrap">仅用于发信息, 更多见[外部API调用]</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td align="left">调用地址：<%=urlAdd0%><br />
      调用参数：<%=urlPar1%></td>
  </tr>
  <tr>
    <td align="center"><IFRAME name=LeftMenu src="<%=urlAdd0%><%=urlPar0%>&tNumb=<%=tel%>&tCont=<%=ct1%>"
      frameBorder=0 width="100%" scrolling="no" height="290"></IFRAME></td>
  </tr>
</table>
<%End If%>
<%If ac="OutRead" Then%>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tdbg">
  <tr>
    <td colspan="2" align="left"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong>外部调用说明</strong></td>
          <td width="10%" align="right">&nbsp;</td>
          <td align="center" nowrap="nowrap">调用说明如下</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td colspan="2" align="left"><table width="100%" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td width="55%">&nbsp;&nbsp;<strong>1. tm,id,sn为必选参数：</strong></td>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2" align="left" valign="top" class="note">* tm为当前时间, 每次调用, 请取系统最新时间, 否则出现超时等认证错误; <br />
            格式:yyyy-mm-dd hh:nn:ss, 上述(now)请自行替换; <br />
            * id为申请的帐号, 
            上述demo或tester请自行替换;<br />
            * sn为加密串, md5(id&amp;&quot;+&quot;&amp;sn&amp;&quot;+&quot;&amp;tm&amp;&quot;+&quot;&amp;url), 其中sn,url为申请帐号时分配的;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td width="55%">&nbsp;&nbsp;<strong>2. 
            act,cs,re为可选参数：</strong></td>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td align="left" valign="top" class="note">* act参数见[3];<br />
            * cs为语言设置, 建议用支持UTF-8,如使用其它语言,内容请编码[5],[6];<br />
            * re为返回设置, Xml表示以xml文件返回, 适合Http,Ajax调用; 
            Report表示网页报告返回, 如表单提交到新窗口等可使用; goBack表示返回原来的页, 并返回如右参数:</td>
          <td align="left" valign="top" class="note">* reAct=Return<br />
            * reState=状态<br />
            * reInt=状态码<br />
            * reInfo=Server.URLEncode(提示信息)，提示信息含中文，已经编码；请自行解码，函数ChrW(a[i])。 <br />
            * reTime=系统时间 </td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td width="55%">&nbsp;&nbsp;<strong>3. act参数：</strong></td>
          <td>&nbsp;&nbsp;<strong>4. 发信息/修改SN的参数：</strong></td>
        </tr>
        <tr>
          <td align="left" valign="top" class="note">* 
            Send: (发信息)外部API调用; <br />
            * 
            SendOut: (发信息)外部iFrame调用;<br />
            * 
            Balance: 查询余额;<br />
            * Readme: 返回说明;</td>
          <td align="left" valign="top" class="note">* tel: 手机号码, 最多20个号码;<br />
            * ct1: 短信内容; 70字内;<br />
            * nsn: 新SN; </td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td width="55%"><strong>&nbsp;&nbsp;5. js内容编码函数:</strong></td>
          <td><strong>&nbsp;&nbsp;6. asp内容编码函数:</strong></td>
        </tr>
        <tr>
          <td align="left" valign="top" class="note">function outEncode(xStr){ <br />
            &nbsp;&nbsp;var n = xStr.length;<br />
            &nbsp;&nbsp;var s = &quot;&quot;;<br />
            &nbsp;&nbsp;for(i=0;i&lt;n;i++){ <br />
            &nbsp;&nbsp;s += xStr.charCodeAt(i) + &quot;;&quot;;<br />
            &nbsp;&nbsp;} <br />
            &nbsp;&nbsp;return s<br />
            }</td>
          <td align="left" valign="top" class="not2">Function outEncode(xStr)<br />
            &nbsp;&nbsp;Dim i,ch,cx,s<br />
            &nbsp;&nbsp;For i=1 to len(xStr)<br />
            &nbsp;&nbsp;ch=Mid(xStr,i,1)<br />
            &nbsp;&nbsp;cx=Hex(AscW(ch))<br />
            &nbsp;&nbsp;s=s&amp;cx&amp;&quot;;&quot;<br />
            &nbsp;&nbsp;Next<br />
            &nbsp;&nbsp;outEncode = s<br />
            End Function</td>
        </tr>
      </table></td>
  </tr>
</table>
<%End If%>
<%If ac="OutPara" Then%>
<div class="line08">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="left" class="fB pTitle">RequireInfo-接收参数</td>
  </tr>
  <tr>
    <td align="center">tm</td>
    <td>客户端系统时间</td>
  </tr>
  <tr>
    <td align="center">tel</td>
    <td>电话号码表</td>
  </tr>
  <tr>
    <td align="center">ct1</td>
    <td>内容</td>
  </tr>
  <tr>
    <td align="center">npw</td>
    <td>新SN</td>
  </tr>
  <tr>
    <td align="center">cs</td>
    <td>编码</td>
  </tr>
  <tr>
    <td align="center">act</td>
    <td>操作类型</td>
  </tr>
  <tr>
    <td align="center">re</td>
    <td>返回类型</td>
  </tr>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">ResponseInfo-返回参数</td>
  </tr>
  <tr>
    <td align="center">reState</td>
    <td>状态</td>
  </tr>
  <tr>
    <td align="center">reInt</td>
    <td>状态码/值</td>
  </tr>
  <tr>
    <td align="center">reInfo</td>
    <td>提示信息</td>
  </tr>
  <tr>
    <td align="center">reTime</td>
    <td>短信系统时间 </td>
  </tr>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">Notes-返回值说明</td>
  </tr>
  <tr>
    <td width="20%" align="center">0</td>
    <td>OK! 成功！</td>
  </tr>
  <!--#include file="../inc/inc_error.asp"-->
</table>
<%End If%>
<div class="line08">&nbsp;</div>
</body>
</html>
