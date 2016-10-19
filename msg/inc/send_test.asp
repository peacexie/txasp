<link rel="stylesheet" type="text/css" href="spub.css"/>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td colspan="2" align="center"><!--#include file="inc_send.asp"--></td>
    </tr>
    <tr>
      <td width="15%" align="center">余额</td>
      <td><input name="tBalance" type="text" id="tBalance" value="<%=sndMaxs%>" size="8" readonly="readonly" style="color:#CCC" />
        条(<%=smcDMod%>), 自动获得,不能修改；</td>
    </tr>
    <%If act="Test" Then%>
    <tr>
      <td width="15%" align="center">手机号码</td>
      <td align="left"><input name="tNumb" type="text" id="tNumb" value="<%=tNumb%>" size="54" maxlength="48" />
        <br />
        最多两个号码, 用半角[<span class="fntF00 fB f14px"><%=smcTSplit%></span>]分开</td>
    </tr>
    <tr>
      <td align="center">短信内容<br />
        1200字内</td>
      <td align="left"><textarea name="tCont" cols="36" rows="4" id="tCont" onChange="chkChars(this,'nowChars',1200)" ><%=tCont%></textarea></td>
    </tr>
    <tr>
      <td align="right"><div id="api_debug_area"></div>
        <input name="act" type="hidden" id="act" value="doTest" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Button" type="button" class="btn60" onclick="chkTest()" value="发送短信" />
        <input name="Button2" type="reset" class="btn60" value="重填资料" />
        <a href="?act=<%=act%>&<%=Rnd_ID("",8)%>=<%=Rnd_ID("",8)%>">重载</a></td>
    </tr>
    <%End If%>
    <tr>
      <td colspan="2" align="left">说明/提示：<span id="nowChars" class="fntF00"><%=msg%></span><br />
        测试发送信息，手机号码，短信内容不做任何检查；</td>
    </tr>
  </form>
</table>
<script type="text/javascript">
var fm = document.fm01;
function chkTest(){   
  eflag=1
  for(i=0;i<1;i++){  
   if (fm.tNumb.value.length==0){ 
     alert('[错误]\n 请输入 手机号码!');
     fm.tNumb.focus();
     eflag=0; break;
   }
   if (fm.tCont.value.length>1200){ 
     alert('[错误]\n 短信内容 请少于1200字!');
     fm.tCont.focus();
     eflag=0; break;
   }
  }
  if (eflag==1){ fm.submit(); }
}
</script>
