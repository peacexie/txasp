<link rel="stylesheet" type="text/css" href="spub.css"/>
<%
If PrmFlag="(Mem)" Then
  sqlS = " TypMod='userSign' And LogAUser='"&Session("MemID")&"' "
Else
  sqlS = " TypMod='sysSign' "
End If
%>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <form action="?" method="post" name="fm01" id="fm01">
    <tr>
      <td colspan="2" align="center"><!--#include file="inc_send.asp"--></td>
    </tr>
    <%If act="doSend" Or act="doGroup" Then%>
    <tr>
      <td align="center">发送报告</td>
      <td><%=msg%><br />
        <a href="?" class="cRed">请返回&gt;&gt;</a> <%=Now()%></td>
    </tr>
    <%Else%>
    <tr>
      <td width="15%" align="center">余额</td>
      <td><input name="tBalance" type="text" id="tBalance" value="<%=sndMaxs%>" size="8" readonly="readonly" style="color:#CCC" />
        条(<%=smcDMod%>), 自动获得,不能修改；</td>
    </tr>
    <%If inStr(act,"Send")>0 Then%>
    <input name="act" type="hidden" id="act" value="doSend" />
    <tr>
      <td align="center">手机号码<br />
        (号码表)<br />
        <a href="#" onclick="fmtTels()">整理号码</a><br /></td>
      <td align="left"><div style="float:right; padding:1px 2px 0px 0px"> <span class="fntF0F">[选择号码]</span><br />
          <%If PrmFlag<>"(Mem)" Then%>
          <%If inStr(Session(UsrPStr),"{Admin}")>0 Then%>
          *.<a href="#" onclick="openWin('../page/stels.asp?tMod=UsrAdmin')">从[管理员]</a><br />
          <%End If%>
          <%If inStr(Session(UsrPStr),"{Admin}")>0 Or inStr(Session(UsrPStr),"ModDocs")>0 Then%>
          *.<a href="#" onclick="openWin('../page/stels.asp?tMod=InnAdmin')">从内部会员</a><br />
          <%End If%>
          <%If inStr(Session(UsrPStr),"{Admin}")>0 Or inStr(Session(UsrPStr),"ModSms")>0 Then%>
          *.<a href="#" onclick="openWin('../page/stels.asp?tMod=SMember')">从短信会员</a><br />
          <%End If%>
          <%End If%>
          *.<a href="#" onclick="openWin('../page/stels.asp?tMod=SmsTels&tPrm=<%=PrmFlag%>')">从[电话薄]</a></div>
        <textarea name="tNumb" cols="36" rows="5" id="tNumb"><%=tNumb%></textarea>
        <div> 最多<%=smcTMax2%>(建议<%=smcTMax1%>)个号码，一行一个或用半角[<span class="fntF00 fB f14px"><%=smcTSplit%></span>]隔开。 </div></td>
    </tr>
    <%Else%>
    <input name="act" type="hidden" id="act" value="doGroup" />
    <tr>
      <td align="center">手机号码<br />
      (群组)<br /></td>
      <td align="left" valign="top">
      
<%
  ModID = RequestS("ModID","C",48) 
  If ModID="" Then ModID="sysTels"
  sqlK = " TelMod='"&ModID&"' "
  sqlUBak = " And LogAUser='(user)' "
  If PrmFlag="(Mem)" Then
	sqlK = sqlK&Replace(sqlUBak,"(user)",PrmUser)
  End If
  sqlO = " ORDER BY TelName,TelID ASC "
  sql = " SELECT * FROM [SmsTelq] WHERE "&sqlK&sqlO
  Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
  rs.Open Sql,conn,1,1 ':echo sql
  Do While NOT rs.EOF
	TelName = rs("TelName")
	TelNums = rs("TelNums")
	'TelNums = Replace(TelNums,vbcrlf,",")
	TelCount = rs("TelCount")
    si = "<div class='itmName' title='("&TelCount&")个号码'><input name='tNumb' type='checkbox' value='"&TelNums&"'>"&TelName&"("&TelCount&")</div>"
    If TelCount="0" Then
    si = "<div class='itmNamu' title='("&TelCount&")个号码'><input name='tNumb' type='checkbox' value='' disabled>"&TelName&"("&TelCount&")</div>"
    End If
	Response.Write si
    rs.MoveNext
  Loop
  rs.Close
  set rs = nothing
%>
      
      </td>
    </tr>
    <%End If%>
    <tr>
      <td align="center">内容模式</td>
      <td><div style="float:right; padding:1px 2px 0px 0px"> [<a href="#" onclick="openWin('../page/message.asp?mType=contMode')">模式说明</a>]&nbsp;&nbsp;</div>
        <select name="tMode" id="tMode" onchange="setMode()">
          <option value="One" selected="selected">兼容模式 --- 推荐！一次只发70个字以内</option>
          <option value="More">手动分割 --- 推荐！最多4条信息,每条70个字以内</option>
          <option value="Long">[长文本] --- 最多255个字(兼容性差)</option>
        </select></td>
    </tr>
    <tr>
      <td align="center">短信内容</td>
      <td><div style="float:right; padding:1px 2px 0px 0px"> &nbsp;<span class="fntF0F">[选择内容]</span><br />
          &nbsp;1.<a href="#" onclick="openWin('../page/stemp.asp?tMod=sysTemp&PrmFlag=<%=PrmFlag%>&tTyp=0BA94E30882DA012')">从常用范本</a><br />
          &nbsp;2.<a href="#" onclick="openWin('../page/stemp.asp?tMod=userTemp&PrmFlag=<%=PrmFlag%>')">从用户范本</a><br />
          &nbsp;3.<a href="#" onclick="openWin('../page/stemp.asp?tMod=sysTemp&PrmFlag=<%=PrmFlag%>')">从系统范本</a></div>
        <div id="tConts" style="width:314px; overflow:hidden;"></div>
        <div id="addBar" style="display:none;">
          <div onclick="addCont()" style="cursor:pointer; color:#06C; text-align:right;">&gt;&gt;&gt;增加第(<span id="addNO">3</span>)条信息</div>
        </div></td>
    </tr>
    <tr>
      <td align="center">内容签名</td>
      <td><div style="float:right; padding:1px 2px 0px 0px"><a href='../user/sign.asp?PrmFlag=<%=PrmFlag%>'>[设置签名]&nbsp;&nbsp;</a></div>
        <select name="tSign" id="tSign" onchange="setSign(this)" style="width:133px">
        <option value="" selected="selected">(选择签名)</option>
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE "&sqlS&" ORDER BY TypID",TmpType)%>
      </select>
        (需要先设置签名)</td>
    </tr>
    <%If smcSTime Then%>
    <tr>
      <td align="center">定时发送</td>
      <td><input name="tTime" type="text" id="tTime" size="20" maxlength="20" />
        (格式:<span class="fntF0F"><%=Now()%></span>, 为空则不定时)</td>
    </tr>
    <%End If%>
    <%If smcSPort Then%>
    <tr>
      <td align="center">子端口</td>
      <td><input name="tPort" type="text" id="tPort" size="20" maxlength="20" />
        (范围:数字<span class="fntF0F">1~65535</span>, 为空则不使用子端口)</td>
    </tr>
    <%End If%>
    <tr>
      <td><div id="api_debug_area"></div>
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td><input name="send" type="button" value="发送短信" onclick="<%=chkFuncs%>" class="btn60" />
        <input name="reset" type="reset" value="重填资料" class="btn60" />
        <a href="?<%=Rnd_ID("",8)%>=<%=Rnd_ID("",8)%>">重载</a></td>
    </tr>
    <tr>
      <td colspan="2" align="left">说明/提示：<span id="nowChars" class="fntF00"><%=msg%></span>
      <div id="bCont1"></div>
      <div id="bCont2"></div>
      <div id="bCont3"></div>
      <div id="bCont4"></div>
      </td>
    </tr>
    <%End If%>
  </form>
</table>
<div id="tmpCont" style="display:none">
  <fieldset style="padding:0px 0px;">
    <legend>内容(n): 最多(70)个字。</legend>
    <textarea name="tCont1" cols="35" rows="4" id="tCont1" onblur="chkChars(this,'nowChars',70)"></textarea>
  </fieldset>
</div>
<script src="../out/smskeys.asp" type="text/javascript"></script>
<script type="text/javascript">
var fm = document.fm01; 
var tmp = document.getElementById("tmpCont").innerHTML;
var objConts = document.getElementById("tConts");
var objBar = document.getElementById("addBar");
var objNO = document.getElementById("addNO");
var cnMax = 4;
try{setMode();}catch(err){}

var telCont="";
var tmpCont="";
var telMaxs="<%=smcTMax2%>";
//var keyCont="<%=xxxParFilALLKeys%>";

</script>
