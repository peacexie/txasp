<!--#include file="config.asp"-->

<!--#include file="../../upfile/sys/para/keywords.asp"-->

<html>
<head>
<title>维护会员资料</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/mem_inc/mem_style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>

<% 

send = Request("send")
fAct = Request("fAct") 

If send = "send" Then

 Call Chk_URL()
 SetShow = ""
 InfCont = RequestS("InfCont",3,512) 
 InfSubj = RequestS("InfSubj",3,2400) 
  fSubj = Chr_Fil2(InfSubj)
  fCont = Chr_Fil2(InfCont)
  SetShow = "Y"
  fMsgShow = ""
  If fSubj<>InfSubj or fCont<>InfCont then
    SetShow = "N"
    fMsgShow = "名称或内容中 含有关键字,需要审核!"
    InfSubj = fSubj
    InfCont = fCont
  End If

If fAct="Ins" Then
sql = "INSERT INTO [TradeCorp] ("
sql = sql& " KeyID,KeyMod,KeyCode"
sql = sql& ",InfType,InfTyp2,InfSubj,InfCont"
sql = sql& ",SetSubj,SetHot,SetTop,SetShow,SetRead"
sql = sql& ",LnkName,LnkMobile,LnkTel,LnkFax"
sql = sql& ",LnkEmail,LnkQQ,LnkAddr,LnkUrl"
sql = sql& ",LogAddIP,LogAUser,LogATime"
sql = sql& ")VALUES("
sql = sql& " '"&Get_AutoID(24)&"','"&RequestS("MemType",3,24)&"','"&rs_AutID(conn,"TradeCorp","KeyCode","Corp-","1","")&"'"
sql = sql& ",'"&RequestS("InfType",3,255)&"','"&RequestS("InfTyp2",3,255)&"','"&InfSubj&"','"&InfCont&"'"
sql = sql& ",'000000','N',888,'"&SetShow&"',0"
sql = sql& ",'"&RequestS("LnkName",3,48)&"','"&RequestS("LnkMobile",3,48)&"','"&RequestS("LnkTel",3,48)&"','"&RequestS("LnkFax",3,48)&"'"
sql = sql& ",'"&RequestS("LnkEmail",3,255)&"','"&RequestS("LnkQQ",3,48)&"','"&RequestS("LnkAddr",3,255)&"','"&RequestS("LnkUrl",3,255)&"'"
sql = sql& ",'"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"'"
sql = sql& ")"
Else
sql = " UPDATE [TradeCorp] SET " 
sql = sql& " InfType = '" & RequestS("InfType",3,255) &"'" 
sql = sql& ",InfTyp2 = '" & RequestS("InfTyp2",3,255) &"'" 
sql = sql& ",InfSubj = '" & InfSubj &"'" 
sql = sql& ",InfCont = '" & InfCont &"'" 
sql = sql& ",SetShow = '" & SetShow &"'" 
sql = sql& ",LnkName = '" & RequestS("LnkName",3,48) &"'" 
sql = sql& ",LnkMobile = '" & RequestS("LnkMobile",3,48) &"'" 
sql = sql& ",LnkTel = '" & RequestS("LnkTel",3,48) &"'" 
sql = sql& ",LnkFax = '" & RequestS("LnkFax",3,48) &"'" 
sql = sql& ",LnkEmail = '" & RequestS("LnkEmail",3,255) &"'" 
sql = sql& ",LnkQQ = '" & RequestS("LnkQQ",3,48) &"'" 
sql = sql& ",LnkAddr = '" & RequestS("LnkAddr",3,255) &"'" 
sql = sql& ",LnkUrl = '" & RequestS("LnkUrl",3,255) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("MemID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE LogAUser='" & Session("MemID") &"' "
End If

 ChkCode = uCase(Request("ChkCode"))
 If Len(InfSubj)>0 And Len(InfCont)>0 And Session("ChkCode")=ChkCode And Len(ChkCode)>1 Then
   Call rs_DoSql(conn,sql)
   Response.Write js_Alert("修改成功!\n"&fMsgShow,"Redir","set_corp.asp")
 Else
   Msg = "认证码错误！或数据提交错误!"
 End If
  
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
KeyID = ""

  rs.Open "SELECT * FROM [TradeCorp] WHERE LogAUser='"&Session("MemID")&"'",conn,1,1 
  If NOT rs.eof then 
KeyID = rs("KeyID")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = Show_Form(rs("InfCont"))
LnkName = rs("LnkName")
LnkMobile = rs("LnkMobile")
LnkTel = rs("LnkTel")
LnkFax = rs("LnkFax")
LnkEmail = rs("LnkEmail")
LnkQQ = rs("LnkQQ")
LnkUrl = rs("LnkUrl")
LnkAddr = rs("LnkAddr")
ImgName = rs("ImgName")
  End If
  rs.Close()

  If KeyID="" Then
rs.Open "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"'",conn,1,1 
  If NOT rs.eof then 
LnkName = rs("MemName")
LnkAddr = rs("MemFrom")
LnkMobile = rs("MemMobile")
LnkTel = rs("MemTel")
LnkEmail = rs("MemEmail")
MemType = rs("MemType")
  End If 
rs.Close()
fAct = "Ins"
  Else
fAct = "Edt"  
  End If
If MemType&""="" Then
  MemType = rs_Val("","SELECT MemType FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"'")
End If
If MemType="Corp" Then
  MsgName = "公司名称"
ElseIf MemType="Privy" Then
  MsgName = "个人姓名"
  InfSubj = LnkName
ElseIf MemType="Gov" Then
  MsgName = "单位名称"
ElseIf MemType="Org" Then
  MsgName = "单位名称"
Else
  MsgName = "名称"
End If
  
SET rs=Nothing


%>
<!--#include file="../../inc/mem_inc/mem_top.asp"-->
            <table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="CCCCCC">
              <form name='fm01' method='post' action='?'>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF"><%=MsgName%>:</td>
                  <td nowrap bgcolor="#FFFFFF"><input name='InfSubj' type='text' id="InfSubj" value="<%=InfSubj%>" size='48' maxlength='120'> 
                    (<%=Get_SOpt(mCfgCode,mCfgName,MemType,"Val")%>用户)
                    <input name='KeyMod' type='hidden' id='KeyMod' value='<%=MemType%>'></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">类 &nbsp;&nbsp;&nbsp;别:</td>
                  <td nowrap bgcolor="#FFFFFF"><select name="InfType" id="InfType" style="width:180px ">
                      <option value="">[请选择]</option>
					  <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='ComTyp' ORDER BY TypID",InfType)%>
                    </select>
                    &nbsp;&nbsp;行 &nbsp;&nbsp;业:
                    <select name="InfTyp2" id="InfTyp2" style="width:180px ">
                      <option value="">[请选择]</option>
					  <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='Fields' ORDER BY TypID",InfTyp2)%>
                    </select></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">主要产品,<br>
                    服务项目,<br>
                    职能范围,<br>
                    或介绍等;<br>
                    1200字内.</td>
                  <td nowrap bgcolor="#FFFFFF"><textarea name="InfCont" cols="48" rows="8" id="InfCont"><%=InfCont%></textarea></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">联系人:</td>
                  <td nowrap bgcolor="#FFFFFF"><input name='LnkName' type='text' id="LnkName" value="<%=LnkName%>" size='24' maxlength='24'>
                    &nbsp;&nbsp;手 &nbsp;&nbsp;&nbsp;机:
                    <input name='LnkMobile' type='text' id="LnkMobile" value="<%=LnkMobile%>" size='24' maxlength='24'></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">电 &nbsp;&nbsp;&nbsp;话:</td>
                  <td nowrap bgcolor="#FFFFFF"><input name='LnkTel' type='text' id="LnkTel" value="<%=LnkTel%>" size='24' maxlength='24'>
                    &nbsp;&nbsp;传 &nbsp;&nbsp;&nbsp;真:
                    <input name='LnkFax' type='text' id="LnkFax" value="<%=LnkFax%>" size='24' maxlength='24'></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">邮 &nbsp;&nbsp;&nbsp;件:</td>
                  <td nowrap bgcolor="#FFFFFF"><input name='LnkEmail' type='text' id="LnkEmail" value="<%=LnkEmail%>" size='24' maxlength='120'>
                    &nbsp;&nbsp;Ｑ &nbsp;&nbsp;&nbsp;Ｑ:
                    <input name='LnkQQ' type='text' id="LnkQQ" value="<%=LnkQQ%>" size='24' maxlength='24'></td>
                </tr>
                <input name='iii_xx' type='hidden' id='iii_xx' value='33'>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">网 &nbsp;&nbsp;&nbsp;址: </td>
                  <td nowrap bgcolor="#FFFFFF"><input name='LnkUrl' type='text' id="LnkUrl" value="<%=LnkUrl%>" size='60' maxlength='120'></td>
                </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">地 &nbsp;&nbsp;&nbsp;址:</td>
                  <td nowrap bgcolor="#FFFFFF"><input name='LnkAddr' type='text' id="LnkAddr" value="<%=LnkAddr%>" size='60' maxlength='120'></td>
                </tr>
            <tr bgcolor="#FFFFFF">
              <td align="right" nowrap>认证码:</td>
              <td align="left" nowrap><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12">
                <img src="../../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../../')"/></td>
            </tr>
                <tr>
                  <td align="right" nowrap bgcolor="#FFFFFF">&nbsp;</td>
                  <td nowrap bgcolor="#FFFFFF"><input type='button' name='Submit' value='提交' onClick='chkData()'>
                    &nbsp;&nbsp;
                    <input type='reset' name='Reset' value='重设'>
                    <input name='send' type='hidden' id='send' value='send'>
                    <font color="#FF00FF"><%=msg%></font>
                    <input name='fAct' type='hidden' id='fAct' value='<%=fAct%>'></td>
                </tr>
              </form>
            </table>
            <script type="text/javascript">


 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 公司名称 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 if (document.fm01.LnkName.value.length==0) 
   {   
     alert(" 姓名 不能为空！"); 
     document.fm01.LnkName.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfType.value.length==0) 
   {   
     alert(" 类别 不能为空！"); 
     document.fm01.InfType.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfTyp2.value.length==0) 
   {   
     alert(" 行业 不能为空！"); 
     document.fm01.InfTyp2.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 公司名称 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length==0) 
   {   
     alert(" 详细内容 不能为空！"); 
     document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length>1200) 
   {   
     alert(" 详细内容 文字多于1200 个字！"); 
     document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 if (document.fm01.ChkCode.value.length==0) 
   {   
     alert(" 认证码 不能为空！"); 
     document.fm01.ChkCode.focus();
     eflag = 1; break;
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
            <br>
            <!--#include file="../../inc/mem_inc/mem_bot.asp"-->
</body>
</html>
