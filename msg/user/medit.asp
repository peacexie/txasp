<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.rGap {	text-align:right;
	padding-right:90px;
}
-->
</style>
</head>
<body>
<%

send = Request("send") 
act = Request("act") 
ID = RequestS("ID",3,48)

PG = RequestS("PG","N",1)
KW = RequestS("KW",3,48)
KT = RequestS("KT",3,48)
ID = RequestS("ID",3,48)

If send = "ins" Then
  MemUrl = Replace(Replace(RequestS("MemUrl","C",255),"http://",""),"https://","")
  sql = " UPDATE [SmsMember] SET " 
  'sql = sql& " MemID = '" & RequestS("MemID",C,48) &"'"  
  sql = sql& " MemCode = '" & RequestS("MemCode","C",96) &"'" 
  sql = sql& ",MemName = '" & RequestS("MemName","C",69) &"'"
  sql = sql& ",MemMobile = '" & RequestS("MemMobile","C",48) &"'"
  'sql = sql& ",MemBalance = '" & RequestS("MemBalance","N",0) &"'"
  sql = sql& ",MemUrl = '" & MemUrl &"'" 
  sql = sql& ",MemFlag = '" & RequestS("MemFlag","C",2) &"'" 
  sql = sql & ",LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
  sql = sql& " WHERE MemID='"&ID&"' "
  'Response.Write sql
  Call rs_DoSql(conn,sql)
  Response.Write js_Alert("修改成功！","Redir","member.asp?KW="&KW&"&KT="&KT&"&Page="&PG&"")
ElseIf act="doCharge" Then 
  Call chkReload()
  CrdID = RequestS("CrdID",3,72)
  CrdPW = RequestS("CrdPW",3,72)
  CrdMoney = RequestS("CrdMoney","N",0)
  CrdCount = RequestS("CrdCount","N",0) 
  CrdNote = RequestS("CrdNote",3,96)
  sql = " UPDATE [SmsMember] SET MemBalance=MemBalance+ " & CrdCount &""
  sql = sql & ",LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
  sql = sql& " WHERE MemID='"&ID&"' "
  'Response.Write sql
  Call rs_DoSql(conn,sql)
  Call logCharge(ID,CrdMoney,CrdCount,"充值"&":("&CrdID&"/"&CrdPW&"):"&CrdNote)
  Response.Write js_Alert("充值成功！","Redir","member.asp?KW="&KW&"&KT="&KT&"&Page="&PG&"")

ElseIf act="doFile" Then  'Else

  s0 = File_Read("../out/smsapi.asp","utf-8") ':Response.Write sSend
  p0 = inStr(s0,"<%")+2
  s0 = Replace(s0,"nsn = Get_GUID(","'nsn = Get_GUID(")
  s0 = "<%"&Mid(s0,p0)
  s0 = "<"&"!--#include file=""config.asp""--"&">"&vbcrlf&s0
  Call File_Add2(Config_Path&"upfile/#dbf#/smsapi.asp."&ID&".Peace!Del",s0,"utf-8")
  
  s0 = File_Read("../out/smscfg.asp","utf-8") ':Response.Write sSend
  p0 = inStr(s0,"<%")+2
  s0 = Replace(s0,"Call Chk_Perm1(","'Call Chk_Perm1(")
  s0 = Replace(s0,"(id)",ID)
  s0 = Replace(s0,"(sn)",Request("sn"))
  s0 = Replace(s0,"(ad1)",Config_URL&Mid(Config_Path,2)&"msg/user/sapi.asp")
  s0 = Replace(s0,"(ad2)",Request("ad2"))
  s0 = Replace(s0,"(tel)",Request("tel"))
  s0 = Replace(s0,"'sms",Request("sms"))
  s0 = "<%"&Mid(s0,p0)
  s0 = "<"&"!--#include file=""md5.asp""--"&">"&vbcrlf&s0
  Call File_Add2(Config_Path&"upfile/#dbf#/smscfg.asp."&ID&".Peace!Del",s0,"utf-8")
 
  Call fil_copy(Config_Path&"sadm/func1/md5_func.asp",Config_Path&"upfile/#dbf#/smsmd5.asp."&ID&".Peace!Del")
  
  's0 = "短信接口API配置说明"&vbcrlf
  s0 = File_Read("../out/smsread.asp","utf-8") ':Response.Write sSend
  s0 = Replace(s0,"(id)",ID)
  s0 = Replace(s0,"(sn)",Request("sn"))
  s0 = Replace(s0,"(ad1)",Config_URL&Mid(Config_Path,2)&"msg/user/sapi.asp")
  s0 = Replace(s0,"(ad2)",Request("ad2"))
  s0 = Replace(s0,"(tel)",Request("tel"))
  's0 = Replace(s0,"<"&"%","")
  's0 = Replace(s0,"%"&">","")
  Call File_Add2(Config_Path&"upfile/#dbf#/smsread.txt."&ID&".Peace!Del",s0,"utf-8")

End If

sql = "SELECT "
sql = sql & " * FROM [SmsMember] "
sql =sql& " WHERE MemID='"&ID&"' "

   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
if NOT rs.EOF then

MemID = rs("MemID")
MemMod = rs("MemMod")
MemCode = rs("MemCode")
MemName = rs("MemName")
MemMobile = rs("MemMobile")
MemBalance = rs("MemBalance")
MemFlag = rs("MemFlag")
MemUrl = rs("MemUrl")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")

end if
  rs.close()
  set rs = nothing
  
Session("ChkCode") = Rnd_ID("",8)

%>
<br>
<%If act="doFile" Then%>
<table width="540" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
      <tr>
        <td width="30%" align="center"><strong>&nbsp;调用代码 生成提示</strong></td>
        <td width="10%" align="right">&nbsp;</td>
        <td align="right"><font color="#FF0000"><%=Msg%></font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" align="left">　　目录/upfile/#dbf#/下，后缀为.<%=ID%>.Peace!Del的4个文件：<br>
      send.asp, config.asp, md5.asp, readme.txt <br>
      为生成的文件，请把后缀.<%=ID%>.Peace!Del去掉，即可使用。</td>
  </tr>

  <tr>
    <td colspan="2" align="left">　　请经常清理生成的文件：系统与设置 &gt;&gt; 数据库管理 &gt;&gt; 进入清理界面。</td>
  </tr>
  <tr>
    <td colspan="2" align="left">　　说明/提示：<span class="fntF00">生成成功！<a href="?ID=<%=ID%>">请返回</a></span></td>
  </tr>
</table>
<div class="line15">&nbsp;</div>
<%Else%>
<table width="540" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
      <tr>
        <td width="30%" align="center"><strong>&nbsp;短信会员 充值</strong></td>
        <td width="10%" align="right">&nbsp;</td>
        <td align="right"><font color="#FF0000"><%=Msg%></font></td>
      </tr>
    </table></td>
  </tr>
  <form action="?" method="post" name="fm02" id="fm02">
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
      <td align="left"><input name="CrdCount" type="text" id="CrdCount" size="24" maxlength="9" /></td>
    </tr>
    <tr>
      <td align="right">折合金额(元)</td>
      <td align="left"><input name="CrdMoney" type="text" id="CrdMoney" size="24" maxlength="9" /></td>
    </tr>
    <tr>
      <td align="right">说明(48字内)</td>
      <td align="left"><input name="CrdNote" type="text" id="CrdNote" size="24" maxlength="48" /></td>
    </tr>
    <tr>
      <td align="right"><input name="act" type="hidden" id="act" value="doCharge" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td align="left"><input name="Button2" type="button" class="btn05" onClick="chkData()" value="接口充值" />
        <input type="reset" name="Button3" value="重填资料" />
        <input name='ID' type='hidden' id='ID' value='<%=ID%>'>
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="KT" type="hidden" id="KT" value="<%=KT%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>"></td>
    </tr>
  </form>
  <tr>
    <td colspan="2" align="left">说明/提示：<span class="fntF00"><%=msg%></span></td>
  </tr>
</table>
<div class="line15">&nbsp;</div>
<table width="540" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <form name='fm01' method='post' action='?'>

    <tr bgcolor="f4f4f4">
      <td height="24" colspan="2"><table width="100%" border="0" cellpadding="1" cellspacing="1">
          <tr>
            <td width="30%" align="center"><strong>&nbsp;短信会员 修改</strong></td>
            <td width="10%" align="right">&nbsp;</td>
            <td align="right"><font color="#FF0000"><%=Msg%></font></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td width="20%" align="right">帐号</td>
      <td><input name='MemID' type='text' value="<%=MemID%>" size='36' maxlength='24' readonly style="color:#CCC">
        (不能修改)</td>
    </tr>
    <tr>
      <td align="right">SN</td>
      <td><input name='MemCode' type='text' id="MemCode" value="<%=MemCode%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td align="right">名称</td>
      <td><input name='MemName' type='text' value="<%=MemName%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td align="right">手机</td>
      <td><input name='MemMobile' type='text' id="MemMobile" value="<%=MemMobile%>" size='36' maxlength='24'></td>
    </tr>
    <tr>
      <td align="right">余额</td>
      <td><input name='MemBalance' type='text' id="MemBalance" value="<%=MemBalance%>" size='36' maxlength='6' readonly style="color:#CCC">
        (条,数字,不能修改)</td>
    </tr>
    <tr>
      <td align="right">Url</td>
      <td><input name='MemUrl' type='text' id="MemUrl" value="<%=MemUrl%>" size='36' maxlength='120'>
        (外部调用Url)</td>
    </tr>
    <tr>
      <td align="right">审核</td>
      <td><select name="MemFlag" id="MemFlag" style="width:120px; ">
          <%=Get_SOpt("Y;N","已审;未审",MemFlag,"")%>
        </select></td>
    </tr>
    <tr bgcolor="f4f4f4">
      <td align="center"><a href="member.asp">返回</a></td>
      <td bgcolor="f4f4f4"><input name='Button' type='button' class="btn60" onClick="chkInfo()" value='提交'>
        
        <input name='Reset' type='reset' class="btn60" value='重设'>
        <a href="?act=doFile&ID=<%=ID%>&sn=<%=MemCode%>&ad2=<%=MemUrl%>&tel=<%=MemMobile%>">生成调用代码</a>
        <input name='send' type='hidden' id='send' value='ins'>
        <input name='ID' type='hidden' id='ID' value='<%=ID%>'>
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="KT" type="hidden" id="KT" value="<%=KT%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>"></td>
    </tr>
  </form>
</table>
<div class="line15">&nbsp;</div>
<%End If%>
<script type="text/javascript">

var fm = document.fm02;
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

 function chkInfo()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////

 if (document.fm01.MemID.value.length<4) 
   {   
     alert("帐号 至少4位");
	 document.fm01.MemID.focus();
	 eflag = 1; break;
   }

 if (document.fm01.MemName.value.length==0) 
   {   
     alert(" 姓名 不能为空！"); 
     document.fm01.MemName.focus();
     eflag = 1; break;
   }

 //tmv = chkF_Date(document.fm01.CrdTime,"日期 不规范！");
 //if (tmv=='ER') { eflag = 1; break; }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>