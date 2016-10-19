<!--#include file="_config.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%
'Response.Write vMem_Btn01&vMem_Btn02
'' /// =========================== 
' 注意，注册信息部分，为了尽力杜绝注册机恶意注册，有些东西跳转和认证很复杂；
' 强烈建议保留这些复杂 跳转和认证，请与后台 系统与设置 >> 参数 >> 文字 共同设置
' Peace(XieYS) 2009-08-31 '
'' /// =========================== 


Call sf_Guard()
Call ChkSpider(1)
Call Chk_Url()
goPage = Request("goPage")
Session.Timeout = 1440 '1440=1天


If Request("act")="Start" Then '点[注册会员]
  If Get_A30Chk(get30Time,get30TSN,720,"")="" Then '半天
    Response.Redirect "../inc/form/form_js.asp?act=Check1&goPage="&goPage&"&verMemb="&verMemb&""&url30Par
  Else
    echo "31.Start!" :Response.End()
  End If
Else '返回[mu_read,mu_app2.asp]
  rPage = Request.ServerVariables("HTTP_REFERER")
  'If inStr(rPage,"/member/mu_"&AppRandom&".asp")>0 Or inStr(rPage,"/member/mu_app2.asp")>0 Then
  If inStr(rPage,Config_Path&"inc/adm_inc/")>0_ 
    Or inStr(rPage,Config_Path&"bbs/")>0_
    Or inStr(rPage,Config_Path&"member/")>0 Then 'inc/adm_inc/adm_menu.asp
    If inStr(rPage,Config_Path&"member/")>0 And Get_A30Chk(get30Time,get30TSN,60,"")<>"" Then '超时:Stop
    'If Get_A30Chk(get30Time,get30TSN,60,"")<>"" Then '超时:Stop
	  echo "32.Start!" :Response.End()
	End If
  Else
      echo "33.Start!" :Response.End()
  End If
End If


't20a = Timer()
Randomize 
RndLen = Int(20*Rnd())+10
RndCode = Rnd_ID("KEY",RndLen)


Dim App23RStr,App23RLen,App23RS01,App23RS02
Call App23Set()
Call App25Set()


rDirBase = Config_Path&"member/mu_"&AppRandom&".asp"
rDir = rDirBase&"?goPage="&goPage&"&verMemb="&verMemb&""
If Chk_URL3(rDirBase)="eUrl" Then 
  Response.Redirect rDir
End If


App22Data = Request.Form(App22Code)
ChkFlag = Request.Form("ChkFlag") ' XX
ChkFRes = Request.Form("ChkFRes") ' OK 
MemType = Request.Form("MemType") ' Corp,
ChkCod11 = Request(get30Tab(1)) ' CSQW2818F2VW13VX2QUP5MWH
ChkCod12 = Request(get30Tab(3)) ' XXXXXXXXXXX...
ChkCod21 = Request(get30Tab(5)) ' 2009-8-25 12:59:44
ChkCod22 = Request(get30Tab(7)) ' XXXXXXXXXXX...
AppRCode = Request("AppRCode") ' CV6-CU18X-AQ2YHBT
App25AR = Request(App25A)
App25CR = Request(App25C)

'Response.Write ChkFlag&" - "&ChkFRes&" - "&MemType&" - "&ChkCod11&" - "&ChkCod12&" - "&ChkCod21&" - "&ChkCod22
If ChkFRes<>"OK" Then '' ////////////////////////////////////////////////////// 认证上页 是否同意会员协议
  Response.Redirect rDir&"&Msg="&vMem_AP1A1
ElseIf MemType="" Then '' ///////////////////////////////////////////////////// 认证上页 是否选择用户类别
  Response.Redirect rDir&"&Msg="&vMem_AP1A2
ElseIf inStr(mCfgCode,MemType)<=0 Then '' ///////////////////////////////////// 认证上页 用户类别是否合法
  Response.Redirect rDir&"&Msg="&vMem_AP1A3
ElseIf ChkCod11="" Or ChkCod21="" Or AppRCode="" Then '' ////////////////////// 三个参数 不能为空 
  Response.Redirect rDir&"?Msg="&vMem_AP1A4
ElseIf DateDiff("n",RequestSafe(ChkCod21,"D","1900-12-31"),Now())>60 Then '' // 60分钟内合法 
  Response.Redirect rDir&"&Msg="&vMem_AP1A5
ElseIf ChkCod12&""<>MD5_User(ChkCod11,AppRandom) Then '' ////////////////////// 随机RSA加密认证1 
  Response.Redirect rDir&"&Msg="&vMem_AP1A6
ElseIf ChkCod22&""<>MD5_User(ChkCod21,"Set1234") Then '' ////////////////////// 随机RSA加密认证2 注意设置Set1234参数  
  Response.Redirect rDir&"&Msg="&vMem_AP1A7
ElseIf AppRCode&""<>App24Read Then '' ////////////////////////////////////////// 认证-防注册机阅读协议:后台设置
  Response.Redirect rDir&"&Msg="&vMem_AP1A8
ElseIf ChkFlag&""<>"XX" Then '' /////////////////////////////////////////////// 认证-防注册机阅读协议:后台设置
  Response.Redirect rDir&"&Msg="&vMem_AP1A9
ElseIF Session("App24Code")&""="" Or Session("App24Code")&""<>Request("App24Code") Then
 If SwhCodRead = "Y" Then
  Response.Redirect rDir&"&Msg="&vMem_AP1AA
 End If
'ElseIF Session(App22Code)&""="" Or Session(App22Code)&""<>App22Data&"" Then
  'Response.Redirect rDir&"&Msg="&vMem_AP1AB&":"&Session(App22Code)&"-"&App22Data&""
ElseIF App25AR="" Or App25AR<>App25B Then
  Response.Redirect rDir&"&Msg="&vMem_AP1AC
ElseIF App25CR="" Or App25CR<>App25D Then
  Response.Redirect rDir&"&Msg="&vMem_AP1AD

ElseIf Get_A30Chk(app30Time,app30TSN,60,"s")<>"" Then
    msg = "错误！超时错误 或 安全认证错误！"

End If
'Response.Write vbcrlf&vbcrlf&"<hr><br>"&App25A
'Response.Write vbcrlf&vbcrlf&"<hr><br>"&App25B
'Response.Write vbcrlf&vbcrlf&"<hr><br>"&App25C
'Response.Write vbcrlf&vbcrlf&"<hr><br>"&App25D


sys27_RVal = Request(App27Random)&""
Dim sys27_Rnd(10)
If sys27_RVal&"" = "" Then
  Response.End()
Else
  For i = 1 To 9
   sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If


 s1=""
 sql = " SELECT * FROM [MemSyst] WHERE SysType='Type' ORDER BY SysID " 
 rs.Open Sql,conn,1,1
 ChkBoxStr = ""
 If NOT rs.EOF Then
 Do While NOT rs.EOF
  SysID = rs("SysID")
  SysDef = rs("SysDef")
  s1 = s1& vbcrlf & "<input type='hidden' name='dType' checked value='"&SysID&"'>"
  s1 = s1& vbcrlf & "<input type='hidden' name='dGrade' checked value='"&SysDef&"'>"
 rs.movenext
 Loop
 End If
 rs.close()



MemID = LCase(RequestS("MemID","C","48"))
MemQu = RequestS("MemQu","C",255)
MemAn = RequestS("MemAn","C",255) 
MemName = RequestS("MemName","C",96)
MemNam2 = RequestS("MemNam2","C",96)
MemCard = RequestS("MemCard","C","48")
MemBirth = RequestS("MemBirth","D","1900-12-31")
MemEmail = RequestS("MemEmail","C","255")
MemTel = RequestS("MemTel","C","48")
MemMobile = RequestS("MemMobile","C","48")

'MemTypeOpt = Get_rsOpt(conn,"SELECT MTPID,MTPName FROM MemType WHERE MTPID LIKE 'Mod%'","ModJob")
'MemTypeOpt = "<select id='MemType' name='MemType'>"&MemTypeOpt&"</select>"

urlEvent = Server.HTMLEncode("class='lgnBtn1 lgnInpt' xtabindex=1 onblur=""ChkAjMem('ChkAjID');""") 
urlParas = "?act=setInput&n="&Get_IDEnc(AppRMemID&sys27_Rnd(1),0)&"&s=24&m=64&e="&urlEvent&url30Par&""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE><%=vMem_Btn01%>-<%=vPMsg_WName%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../inc/home/jsPass.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</HEAD>
<BODY>
<%
If SwhMemApp<>"Y" And Session("UsrID")&""="" Then 
  Response.Write "<br><br><center>"&vMem_RD21&"</center><br><br>"
  Response.End()
End If
TPName = vMem_Btn01&"("&Get_SOpt(mCfgCode,mCfgName,MemType,"Val")&")"
Response.Write "<!--"&RndCode&"-->"&vbcrlf
%>
<!--#include file="_ftop2.asp"-->
      <table width="90%" align="center" cellpadding="2" cellspacing="1">
        <form name='fm01<%=sys27_RVal%>' method='post' action='mu_app2.asp'>
          <input type="hidden" name="<%=App23Date%>RFlg" id="<%=App23Date%>RFlg" value="<%=App23Date%><%=App23RLen%>" />
          <input type="hidden" name="<%=App23Date%>RChk" id="<%=App23Date%>RChk" value="<%=App23RS01%><%=App23RS02%>" />
          <input type="hidden" name="<%=App23Date%>RStr" id="<%=App23Date%>RStr" value="<%=App23RStr%>" />
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE">&nbsp;</td>
            <td align="left" nowrap bgcolor="#FFFFEE">&nbsp;</td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE">
            
                <div style="display:none;">
                  xxID(帐号)<input name="xxID" type="text" id="xxID" value="(xxID)" size="12" onChange="sys27_stop()" maxlength="8">
                  Name(姓名)<input name="Name" type="text" id="Name" value="(Name)" size="12" maxlength="8">
                </div>
            
            <FONT color=#ff0000>*</FONT><%=vMem_AP1D1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE">
              <script src="../inc/form/form_js.asp<%=urlParas%>" type='text/javascript'></script>
              <input name='<%=AppRMemID%><%=sys27_Rnd(1)%><%=Timer()%>' type='hidden' value="<%=MemID%>" size='24' maxlength='64'>
              <input type="button" name="Button" value="<%=vMem_AP1DD%>" onClick="CheckID('check_id.asp');" tabindex="999">
              <br>
              <%=vMem_AP1DA%></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><FONT color=#ff0000>*</FONT><%=vMem_AP1D2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name="MemPW<%=sys27_Rnd(2)%>" type="password" id="MemPW<%=sys27_Rnd(2)%>" value="" size="24" maxlength="24" onKeyUp="ChkPassStrength(this.value);" onblur="ChkAjMem('ChkAjPW');" style="float:left; margin-right:1px">
              
<div id="pws" class="memb_pws">
  <div id="idSM1" class="memb_pws0"><span style="font-size:1px">&nbsp;</span><span id="idSMT1">弱</span></div>
  <div id="idSM2" class="memb_pws0" style="border-left:solid 1px #DEDEDE"><span style="font-size:1px">&nbsp;</span><span id="idSMT2">中</span></div>
  <div id="idSM3" class="memb_pws0" style="border-left:solid 1px #DEDEDE"><span style="font-size:1px">&nbsp;</span><span id="idSMT3">强</span></div>
</div>
              
              <div style="clear:both">
              <%=vMem_AP1DB%></div> </td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><FONT 
                              color=#ff0000>*</FONT><%=vMem_AP1D3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name="MemPW2<%=sys27_Rnd(2)%>" type="password" id="MemPW2<%=sys27_Rnd(2)%>" value="" size="24" maxlength="24" onblur="ChkAjMem('ChkAjPW')">
              <br>
              <%=vMem_AP1DC%> </td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><%=vMem_AP1D4%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name='MemQu<%=sys27_Rnd(3)%>' type='text' value="<%=MemQu%>" size='24' maxlength='18'>
              (<%=vMem_AP1DE%>)</td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><%=vMem_AP1D5%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name='MemAn<%=sys27_Rnd(3)%>' type='text' value="<%=MemAn%>" size='24' maxlength='18'>
(<%=vMem_AP1DF%>)</td>
          </tr>
          <input type='hidden' name='MemType' id="MemType" size='24' value="<%=MemType%>">
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">公司----------------------------</td>
          </tr>
          <%End If%>
          <%If MemType="Corp" Then%>
          <tr>
            <td align="right" nowrap><%=vMem_InfA1%>:</td>
            <td align="left" nowrap><input name='<%=AppRName%><%=sys27_Rnd(4)%>' type='text' value="<%=MemName%>" size='48' maxlength='48'></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfA2%>:</td>
            <td align="left" nowrap><select name="MemTyp2" id="MemTyp2" style="width:150px;">
                <option value=""><%=vMem_InfAA%></option>
				<%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='Fields'",MemTyp2)%>
              </select></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfA3%>:</td>
            <td align="left" nowrap><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'> 
              &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfA4%>:</td>
            <td align="left" nowrap><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfA5%>:</td>
            <td align="left" nowrap><input name='MemNam2<%=sys27_Rnd(4)%>' type='text' id="MemNam2<%=sys27_Rnd(4)%>" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80px">
                <option value="F" selected><%=vMem_InfAC%></option>
                <option value="M"><%=vMem_InfAD%></option>
                <option value="X"><%=vMem_InfAE%></option>
              </select></td>
          </tr>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">个人----------------------------</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Privy" Then%>
          <input type='hidden' name='MemTyp2' id="MemTyp2" size='24' value="<%=MemTyp2%>">
          <input type='hidden' name='MemNam2<%=sys27_Rnd(4)%>' id="MemNam2<%=sys27_Rnd(4)%>" size='24' value="<%=MemNam2%>">
          <tr>
            <td align="right" nowrap><%=vMem_InfB1%>:</td>
            <td align="left" nowrap><input name='<%=AppRName%><%=sys27_Rnd(4)%>' type='text' value="<%=MemName%>" size='24' maxlength='24'>
            <select name="MemSex" id="MemSex" style="width:80px">
                <option value="F" selected><%=vMem_InfAC%></option>
                <option value="M"><%=vMem_InfAD%></option>
                <option value="X"><%=vMem_InfAE%></option>
            </select></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfB2%>:</td>
            <td align="left" nowrap><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfB3%>:</td>
            <td align="left" nowrap><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">机构----------------------------</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Gov" Then%>
          <tr>
            <td align="right" nowrap><%=vMem_InfC1%>:</td>
            <td align="left" nowrap><input name='<%=AppRName%><%=sys27_Rnd(4)%>' type='text' value="<%=MemName%>" size='48' maxlength='48'></td>
          </tr>
          <input type='hidden' name='MemTyp2' id="MemTyp2" size='24' value="<%=MemTyp2%>">
          <tr>
            <td align="right" nowrap><%=vMem_InfC2%>:</td>
            <td align="left" nowrap><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'>
            &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfC3%>:</td>
            <td align="left" nowrap><input name='MemNam2<%=sys27_Rnd(4)%>' type='text' id="MemNam2<%=sys27_Rnd(4)%>" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80px">
                <option value="F" selected><%=vMem_InfAC%></option>
                <option value="M"><%=vMem_InfAD%></option>
                <option value="X"><%=vMem_InfAE%></option>
              </select></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfC4%>:</td>
            <td align="left" nowrap><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>

          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">团体----------------------------</td>
          </tr>
          <%End If%>
          <%ElseIf MemType="Org" Then%>
          <tr>
            <td align="right" nowrap><%=vMem_InfD1%>:</td>
            <td align="left" nowrap><input name='<%=AppRName%><%=sys27_Rnd(4)%>' type='text' value="<%=MemName%>" size='48' maxlength='48'></td>
          </tr>
          <input type='hidden' name='MemTyp2' id="MemTyp2" size='24' value="<%=MemTyp2%>">
          <tr>
            <td align="right" nowrap><%=vMem_InfD2%>:</td>
            <td align="left" nowrap><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='24'>
            &nbsp;<%=vMem_InfAB%></td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfD3%>:</td>
            <td align="left" nowrap><input name='MemBirth' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
              (<%=vMem_InfAF%>:1900-12-31)</td>
          </tr>
          <tr>
            <td align="right" nowrap><%=vMem_InfD4%>:</td>
            <td align="left" nowrap><input name='MemNam2<%=sys27_Rnd(4)%>' type='text' id="MemNam2<%=sys27_Rnd(4)%>" value="<%=MemNam2%>" size='24' maxlength='12'>
              <select name="MemSex" id="MemSex" style="width:80px">
                <option value="F" selected><%=vMem_InfAC%></option>
                <option value="M"><%=vMem_InfAD%></option>
                <option value="X"><%=vMem_InfAE%></option>
              </select></td>
          </tr>
          
          <%End If%>
          <%If 1=2 Then%>
          <tr>
            <td colspan="2" bgcolor="#FFFFEE">联系----------------------------</td>
          </tr>
          <%End If%>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><%=vMem_AP1E1%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name='MemTel<%=sys27_Rnd(5)%>' type='text' id="MemTel<%=sys27_Rnd(5)%>" value="<%=MemTel%>" size='24' maxlength='24'></td>
          </tr>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><%=vMem_AP1E2%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name='MemMobile<%=sys27_Rnd(5)%>' type='text' id="MemMobile<%=sys27_Rnd(5)%>" value="<%=MemMobile%>" size='24' maxlength='24'></td>
          </tr>
          
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE"><font color="#ff0000">*</font><%=vMem_AP1E3%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name='<%=app30Tab(9)%><%=sys27_Rnd(5)%>' type='text' value="<%=MemEmail%>" size='24' maxlength='120'></td>
          </tr>
          <tr>
            <td width="30%" align="right" nowrap bgcolor="#FFFFEE"><%=vMem_AP1E4%>:</td>
            <td align="left" nowrap bgcolor="#FFFFEE"><input name="<%=App21Code%>" type="text" id="<%=App21Code%>" size="6" maxlength="12" onblur="ChkAjMem('ChkAjCode')">
              <img src="../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../')"/></td>
          </tr>
          <tr>
            <td colspan="2" align="center" nowrap bgcolor="#FFFFEE">
            <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
            <input name='Chk123' type='hidden' id='Chk123' value='12' />
              <input type='button' name='SendOK' value='<%=vMem_Btn14%>' onClick='chkData()' xdisabled="disabled" id="SendOK">
              &nbsp;<a href="mu_<%=AppRand12%>.asp?goPage=<%=goPage%>&verMemb=<%=verMemb%>"><%=vMem_AP1EA%></a> &nbsp;
              <input type='reset' name='Reset' value='<%=vMem_Btn12%>'>
              <input name='send' type='hidden' id='send' value='ins'>
              <input name='goPage' type='hidden' id='goPage' value='<%=goPage%>'>
              </td>
          </tr>
          <tr>
            <td colspan="2" align="left" nowrap bgcolor="#FFFFEE">&nbsp;&nbsp;&nbsp;&nbsp;<%=vMem_AP1EB%> <a href="<%=Config_Path&"member/mu_"&AppRandom&".asp"%>"><font color="#0000FF"><%=vMem_AP1EC%></font></a><%=vMem_AP1ED%></td>
          </tr>
          <%=s1%>
          <tr>
            <td align="right" nowrap bgcolor="#FFFFEE">&nbsp;</td>
            <td align="left" nowrap bgcolor="#FFFFEE">&nbsp;</td>
          </tr>
          <input name='<%=App25A%>' type="hidden" value='<%=App25D%>' />
          <input name="<%=App25B%>" type="hidden" value="<%=App25C%>" />
          <input name="<%=App28Code%>" type="hidden" value="<%=Get_AHour(0)%>" />
          <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>" />
          <div id="App26Div" style="display:none"></div>
          <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
          <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
          <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
          <input name="_trakeTime" type="hidden" value='<%=Request("_trakeTime")%>' />
        </form>
      </table>
      <script type="text/javascript">

function DispPwdStrength(iN,sHL){
  if(iN>3){ iN=3;}
  for(var i=1;i<4;i++){
	var sHCR="memb_pws0";
	if(i<=iN){ sHCR=sHL;}
	if(iN>0){
	getElmID("idSM"+i).className=sHCR;
	}
	//getElmID("idSMT"+i).className="ob_pwfont2";
	if (iN>0){
		if (i<=iN){
		getElmID("idSMT"+i).style.display=((i==iN)?"inline":"none");
		}
	}
	else{
	getElmID("idSMT"+i).style.display=((i==iN)?"none":"inline");
	}
  }
}
//out_upwd1();this.className='input_onBlur'

var fmObj = document.fm01<%=sys27_RVal%>;
var fmMemID = fmObj.<%=AppRMemID%><%=sys27_Rnd(1)%>;
var fmMName = fmObj.<%=AppRName%><%=sys27_Rnd(4)%>;
var fmChkFG = fmObj.Chk123;
var fmEmail = fmObj.<%=app30Tab(9)%><%=sys27_Rnd(5)%>;

function sys27_stop()
{
  //alert(11); //暂时未用
  fmObj.<%=App21Code%>.readonly = true;
  //alert(12); 
}

var xmlHttp = getXmlHttp(); //ChkAjID,ChkAjCode
function ChkAjMem(xType) {
	//if(xType=='ChkAjCode'){ 
	  //fmObj.SendOK.disabled = true; // 还原
	  //fmChkFG.value = fmChkFG.value.replace('A','1'); 
	//}
	//if(xType=='ChkAjID'){   
	  //fmObj.SendOK.disabled = true; // 还原
	  //fmChkFG.value = fmChkFG.value.replace('B','2'); 
	//}
  if(xType=='ChkAjPW'){return;} //2011-04-21:未使用
  var url = "check_id.asp?yAct="+xType+"&MemID="+fmMemID.value+"&ChkCode="+fmObj.<%=App21Code%>.value; 
  xmlHttp.open("GET", url, true); //这里的true代表是异步请求 ChkAjCode
  xmlHttp.onreadystatechange = ChkAjUpd;
  xmlHttp.send(null);
}
function ChkAjUpd(){
  if (xmlHttp.readyState == 4) {
	var rData = xmlHttp.responseText; 
	var rMsg = "";
	if(rData=="N.Code"){ rMsg = "<%=vMem_AP1B1%>";}
	if(rData=="N.<4"){   rMsg = "<%=vMem_AP1B2%>";}
	if(rData=="N.Adm"){  rMsg = "<%=vMem_AP1B3%>";}
	if(rData=="N.Key"){  rMsg = "<%=vMem_AP1B4%>";}
	if(rData=="N.Mem"){  rMsg = "<%=vMem_AP1B5%>";}
	if(rData.substring(0,2)=='N.'){alert(rMsg);}
	if(rData.substring(0,6)=='<input'){ 
	  //App26Div.innerHTML = rData; //2011-04-21:未使用
	  //alert(rData) 
	}
	//if(rData=='Y.Code'){ 
	  //fmChkFG.value = fmChkFG.value.replace('1','A'); 
	//}
	//if(rData=='Y.ID'){   
	  //fmChkFG.value = fmChkFG.value.replace('2','B');  
	//}
	//if(fmChkFG.value=="AB"){ 
	  //fmObj.SendOK.disabled = false; 
	//}else{
	  //if(rData.substring(0,2)=='N.'){alert(rMsg);}
	//}
  }
}

function CheckID(pDir) //
{
  //alert(fmObj.User_ID.value);
  window.open(pDir + '?MemID='+fmMemID.value+'&verMemb=<%=verMemb%>', 'popUpWin', 'toolbar=no,location=no,directories=no,status=no,menub ar=no,scrollbar=no,resizable=no,copyhistory=yes,width=480,height=360,left=240, top=120');
}

 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////
 var tmp = chkF_ID(fmMemID,"<%=vMem_AP1C1%>")
 if ( (fmMemID.value.length<4) || (tmp=="ER") )
   {   
     eflag = 1; break;
   }
 var tmp = chkF_PW(fmObj.MemPW<%=sys27_Rnd(2)%>,"<%=vMem_AP1C2%>")
 if ( (fmObj.MemPW<%=sys27_Rnd(2)%>.value.length<5) || (tmp=="ER") )
   {   
     eflag = 1; break;
   }
 if (fmObj.MemPW<%=sys27_Rnd(2)%>.value==fmMemID.value)
   {   
     alert('<%=vMem_AP1C3%>');
     fmObj.MemPW<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }

 if ( !(fmObj.MemPW<%=sys27_Rnd(2)%>.value==fmObj.MemPW2<%=sys27_Rnd(2)%>.value) )
   {   
     alert('<%=vMem_AP1C4%>');
     fmObj.MemPW2<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }
 if (fmObj.MemQu<%=sys27_Rnd(3)%>.value.length==0) 
   {   
     //alert(" 提示问题 不能为空！"); 
     //fmObj.MemQu.focus();
     //eflag = 1; break;
   }
 if (fmObj.MemAn<%=sys27_Rnd(3)%>.value.length==0) 
   {   
     //alert(" 问题答案 不能为空！"); 
     //fmObj.MemAn.focus();
     //eflag = 1; break;
   }
 if (fmMName.value.length==0) 
   {   
     alert("<%=vMem_AP1C5%>"); 
     fmMName.focus();
     eflag = 1; break;
   }
 //tmv = chkF_Date(fmObj.MemBirth,"日期 不规范！");
 //if (tmv=='ER') 
   { 
     //eflag = 1; break;
   }
 tmv = chkF_Mail(fmEmail,"<%=vMem_AP1C6%>");
 if (tmv=='ER') 
   { 
     eflag = 1; break;
   }
 
 <%If MemType="Corp" Then%>
 if(fmObj.MemTyp2.value.length==0)
   {   
     alert('<%=vMem_AP1C7%>');
     fmObj.MemTyp2.focus();
     eflag = 1; break;
   }
  <%End If%>
 
 if(fmObj.<%=App21Code%>.value.length==0)
   {   
     alert('<%=vMem_AP1C8%>');
     fmObj.<%=App21Code%>.focus();
     eflag = 1; break;
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ fmObj.submit(); }
}</script>
      <!--#include file="_fbot2.asp"-->
<%
't20b = Timer()
'Response.Write t20b - t20a
%>
</BODY>
</HTML>
