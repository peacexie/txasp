<!--#include file="../../upfile/sys/pcfg/jmail.asp"--> 
<!--#include file="../../upfile/sys/pcfg/link.asp"--> 
<!--#include file="../../sadm/func2/func_obj.asp"--> 

<%

    ModTab = "GboSend"'rel_ModTab(MD)
    upPart = "defdt"'rel_TabPath(ModTab)
    upRoot = Config_Path&"upfile/"

If verNow="2" Then

 vJApp_mSubj = "Job Apply"
 vJApp_mClose = "Close"
 vJApp_mPrint = "Print"
 vJApp_mTime = "Time"
 vJApp_mInp = " Input "
 vJApp_mBot = " Please Input the Details, We will Contact you SOON! "
 vJApp_Subj = "Subject"
 vJApp_Cont = "Content"
 vJApp_Name = "Name"
 vJApp_Tel = "Tel"
 vJApp_Order = "Order: "
 vJApp_BtnOK = "OK Send"
 vJApp_BtnNO = "Reset"
 vJApp_MsgNull = "Error, Null Subject or Contenct!"
 vJApp_MsgCode = "Check Code Error,Please Click the Code Picture And Send Again! "
 vJApp_MsgOK = "Send OK!! \n We will Contact you SOON! \nClick Close"
 vJApp_jsSubj = "Job Name is Null!"
 vJApp_jsCont = "Resume is Null!"
 vJApp_jsName = "Name is Null!"
 vJApp_jsTel = "Tel is Null!"
 
 vJApp_tBase = "Base Info."
 vJApp_tJob = "Job intentions"
 vJApp_tResume = "Resume"
 vJApp_tStudy = "Study Experience"
 vJApp_tWork = "Work Experience"
 
 vJApp_t1Name = "Name"
 vJApp_t1Code = "Pos ID"
 vJApp_t1S_e_x = "S e x"
 vJApp_t1S_e_1 = "Female"
 vJApp_t1S_e_2 = "Male"
 vJApp_t1Edu = "Education"
 vJApp_t1Birth = "Birth"
 vJApp_t1Prof = "Professional"
 vJApp_t1Native = "Native land"
 vJApp_t1IDCard = "ID.Card"
 
 vJApp_t1Tel = "Tel"
 vJApp_t1Mob = "Mobile"
 vJApp_t1Mail = "E-Mail"
 vJApp_t1Addr = "Addr."
 vJApp_t1Else = "Else"
 
 vJApp_t2Pos = "Position"
 vJApp_t2Field = "Field"
 vJApp_t2Addr = "Job Addr."
 vJApp_t2Wage = "Wage"
 vJApp_t2Addr2 = "Dongguan"
 vJApp_t2Wage2 = "Interview"
 
Else

 vJApp_mSubj = "职位应聘"
 vJApp_mClose = "关闭"
 vJApp_mPrint = "打印"
 vJApp_mTime = "时间"
 vJApp_mInp = " 请输入 "
 vJApp_mBot = " 请认真输入，提交资料后，我们尽快处理！"
 vJApp_Subj = "标题"
 vJApp_Cont = "内容"
 vJApp_Name = "姓名"
 vJApp_Tel = "电话"
 vJApp_Order = "订购: "
 vJApp_BtnOK = "提交"
 vJApp_BtnNO = "重填"
 vJApp_MsgNull = "输入为空！请点击认证码图片，刷新认证码再提交一次！"
 vJApp_MsgCode = "认证码错误！请点击认证码图片，刷新认证码再提交一次！ "
 vJApp_MsgOK = "提交成功!如果有需要，我们会尽快联络您！\n点击关闭"
 vJApp_jsSubj = "职位 不能为空!"
 vJApp_jsCont = "自我简介不能为空!"
 vJApp_jsName = "姓名不能为空!"
 vJApp_jsTel = "电话不能为空!"
 
 vJApp_tBase = "基本信息"
 vJApp_tJob = "职位信息"
 vJApp_tResume = "自我介绍"
 vJApp_tStudy = "学习经历"
 vJApp_tWork = "工作经历"
 
 vJApp_t1Name = "姓　名"
 vJApp_t1Code = "简历编号"
 vJApp_t1S_e_x = "性　别"
 vJApp_t1S_e_1 = "女"
 vJApp_t1S_e_2 = "男"
 vJApp_t1Edu = "学　历"
 vJApp_t1Birth = "出生日期"
 vJApp_t1Prof = "专　业"
 vJApp_t1Native = "籍贯地"
 vJApp_t1IDCard = "身份证号"
 
 vJApp_t1Tel = "联系电话"
 vJApp_t1Mob = "手　机"
 vJApp_t1Mail = "电子邮件"
 vJApp_t1Addr = "联系地址"
 vJApp_t1Else = "其他方式"
 
 vJApp_t2Pos = "职　位"
 vJApp_t2Field = "行　业"
 vJApp_t2Addr = "工作地点"
 vJApp_t2Wage = "期望工资"
 vJApp_t2Addr2 = "东莞"
 vJApp_t2Wage2 = "面议"
 
End If

ObjSubj = RequestS("ObjSubj","C",255) 
ObjID = RequestS("ObjID","C",48)

sSubj = RequestS("InfSubj","C",48)
sType = ObjSubj
sCont = RequestS("sysCont","C",240000)
'sPage = Request.Servervariables("HTTP_REFERER")
send = Request("send")
ChkCode = Request("ChkCode")

If send="send" Then

 If Config_Cont="DB" Then
  xxxCont = sCont
 Else
  xxxCont = ""
 End If

  Call Chk_URL()	
  'Response.End()
  IP = Get_CIP()
  KeyID = rs_AutID(conn,"GboInfo","KeyID",upPart,"1","") 'Get_AutoID(24)
   'Call Send_Mails(Config_Name,xTo,sSubj,hCont,"gb2312")
sql = " INSERT INTO [GboSend] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont"  
sql = sql& ", LogAddIP" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & MD & "'" 
sql = sql& ", '" & RequestS("InfCode","C",48) &"'" 
sql = sql& ", '" & ObjSubj &"'" 
sql = sql& ", '" & sSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  If Trim(sSubj)="" Or Trim(sCont)="" Then
    fMsgShow = vJApp_MsgNull
	PicReLoad = "PicReLoad('../');" 
	Response.Write js_Alert(fMsgShow,"Alert","-1")
  ElseIf Session("ChkCode")<>uCase(ChkCode) Then
    fMsgShow = vJApp_MsgCode
	PicReLoad = "PicReLoad('../');" 
	Response.Write js_Alert(fMsgShow,"Alert","-1")
  Else
    Call rs_Dosql(conn,sql)
	InfCont = sCont
    upPath = upRoot&Replace(KeyID,"-","/") 
    Call add_sfFile()
    Response.Write js_Alert(vJApp_MsgOK,"Close","?") 
	Response.End()
  End If
End If

'Response.Write send&sSubj&sCont
'MD = "GboK124"

InfCode = rs_AutID(conn,"GboSend","KeyCode",upPart,"1","")
ObjSubj = Show_Form(ObjSubj)
ObjID = Show_Form(ObjID)
%>

<form method="post" name='fmMail' action='?'>
  <input type="hidden" name='send' value="send" />
  <INPUT type=hidden name='sysSubj' value='[<%=vJApp_mSubj%>]'>
  <INPUT type=hidden name='sysTo' value='xpigeon@163.com'>
  <INPUT type=hidden name='sysFrom' value='[<%=vPMsg_WName%>]'>
  <INPUT type=hidden name='sysCont'>
  <INPUT type=hidden name='sysEnd' value="Close">
  <div id="SMailID">
<table width="640" border="0" align="center">
  <tr>
    <td align="center"><table width='100%' border="0" align="center" cellpadding="2" cellspacing="0">
      <tbody>
        <tr align="middle">
          <th height="30" align="center" nowrap="nowrap">[<%=vPMsg_WName%>]<%=vJApp_mSubj%></th>
          <td width="30%" align="center" nowrap="nowrap">&nbsp;</td>
          <td width="10%" align="center" nowrap="nowrap"><span style="cursor:hand" onclick="window.close();" title="点击关闭窗口"><%=vJApp_mClose%></span></td>
          <td width="10%" align="center" nowrap="nowrap"><span style="cursor:hand" onclick="javascript:window.print();" title="点击打印本页"><%=vJApp_mPrint%></span></td>
        </tr>
      </tbody>
    </table></td>
  </tr>
</table>
<table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><FIELDSET style="padding:12px;">
          <LEGEND><span class="Tit01"><%=vJApp_tBase%></span></LEGEND>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD vAlign=top><TABLE width='100%' border=0 align=center cellPadding=2 cellSpacing=0>
                    <TBODY>
                      <TR align=middle>
                        <td align="center" nowrap><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
                            <tr bgcolor="#FFFFFF">
                              <td width="15%" align="center"><%=vJApp_t1Name%></td>
                              <td width="22%" align="left"><input name="InfSubj" type="text" id="InfSubj" size="18" maxlength="24"></td>
                              <td width="15%" align="center"><%=vJApp_t1Code%></td>
                              <td width="22%" align="left"><input name="InfCode" type="text" id="InfCode" value="<%=InfCode%>" size="18" maxlength="24"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1S_e_x%></td>
                              <td align="left">
                                <input name="InfXBie" type="radio" id="radio" value="<%=vJApp_t1S_e_1%>" checked="checked" />
                                <%=vJApp_t1S_e_1%>&nbsp;&nbsp;
                                <input type="radio" name="InfXBie" id="radio2" value="<%=vJApp_t1S_e_2%>" />
                              <%=vJApp_t1S_e_2%></td>
                              <td align="center"><%=vJApp_t1Edu%></td>
                              <td align="left"><input name="InfXXX1" type="text" id="InfXXX1" size="18" maxlength="24"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Birth%></td>
                              <td align="left"><input name="InfXXX2" type="text" id="InfXXX2" size="18" maxlength="24"></td>
                              <td align="center"><%=vJApp_t1Prof%></td>
                              <td align="left"><input name="InfXXX3" type="text" id="InfXXX3" size="18" maxlength="24"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Native%></td>
                              <td align="left"><input name="InfXXX4" type="text" id="InfXXX4" size="18" maxlength="24"></td>
                              <td align="center"><%=vJApp_t1IDCard%></td>
                              <td align="left"><input name="InfXXX5" type="text" id="InfXXX5" size="18" maxlength="24"></td>
                            </tr>                           
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Tel%></td>
                              <td align="left"><input name="LnkTel" type="text" id="LnkTel" size="18" maxlength="24"></td>
                              <td align="center"><%=vJApp_t1Mob%></td>
                              <td align="left"><input name="LnkMobile" type="text" id="LnkMobile" size="18" maxlength="24"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Mail%></td>
                              <td colspan="3" align="left"><input name="LnkMail" type="text" id="LnkMail" size="48" maxlength="60"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Addr%></td>
                              <td colspan="3" align="left"><input name="LnkAddr" type="text" id="LnkAddr" size="48" maxlength="60"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t1Else%></td>
                              <td colspan="3" align="left"><input name="LnkElse" type="text" id="LnkElse" size="48" maxlength="60"></td>
                            </tr>
                          </table></td>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
          </FIELDSET></td>
      </tr>
    </table>
    <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><FIELDSET style="padding:12px;">
          <LEGEND><span class="Tit01"><%=vJApp_tJob%></span></LEGEND>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD vAlign=top><TABLE width='100%' border=0 align=center cellPadding=2 cellSpacing=0>
                    <TBODY>
                      <TR align=middle>
                        <td align="center" nowrap><table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
                            <tr bgcolor="#FFFFFF">
                              <td width="20%" align="center"><%=vJApp_t2Pos%></td>
                              <td width="30%"><input name="ObjSubj" type="text" id="ObjSubj" value="<%=ObjSubj%>" size="18" maxlength="24"></td>
                              <td width="20%" align="center"><%=vJApp_t2Field%></td>
                              <td><input name="InfXXY1" type="text" id="InfXXY1" size="18" maxlength="24"></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td align="center"><%=vJApp_t2Addr%></td>
                              <td><input name="InfXXY2" type="text" id="InfXXY2" value="<%=vJApp_t2Addr2%>" size="18" maxlength="24"></td>
                              <td align="center"><%=vJApp_t2Wage%></td>
                              <td><input name="InfXXY3" type="text" id="InfXXY3" value="<%=vJApp_t2Wage2%>" size="18" maxlength="24"></td>
                            </tr>
                          </table></td>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
          </FIELDSET></td>
      </tr>
    </table>
    <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><FIELDSET style="padding:12px;">
          <LEGEND><span class="Tit01"><%=vJApp_tResume%></span></LEGEND>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD vAlign=top><TABLE border=0 align=center cellPadding=2 cellSpacing=0>
                    <TBODY>
                      <TR align=middle>
                        <td align="center" [Flag]><textarea name="InfCont" cols="72" rows="8" id="InfCont"></textarea></td>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
          </FIELDSET></td>
      </tr>
    </table>
    <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><FIELDSET style="padding:12px;">
          <LEGEND><span class="Tit01"><%=vJApp_tStudy%></span></LEGEND>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD vAlign=top><TABLE border=0 align=center cellPadding=2 cellSpacing=0>
                    <TBODY>
                      <TR align=middle>
                        <td align="center" [Flag]><textarea name="InfCont2" cols="72" rows="8" id="InfCont2"></textarea></td>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
          </FIELDSET></td>
      </tr>
    </table>
    <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><FIELDSET style="padding:12px;">
          <LEGEND><span class="Tit01"><%=vJApp_tWork%></span></LEGEND>
          <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
            <TBODY>
              <TR>
                <TD vAlign=top><TABLE border=0 align=center cellPadding=2 cellSpacing=0>
                    <TBODY>
                      <TR align=middle>
                        <td align="center" [Flag]><textarea name="InfCont3" cols="72" rows="8" id="InfCont3"></textarea></td>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            </TBODY>
          </TABLE>
          </FIELDSET></td>
      </tr>
    </table>
    <table width="640" border="0" align="center">
      <tr>
        <td width="50%" align="center"><%=vJApp_mTime%>: <%=Now()%></td>
        <td width="50%" align="right">[<%=vPMsg_WName%>]<%=vJApp_mSubj%></td>
      </tr>
    </table>
    <!--KillA-->
    <table width="640" border="0" align="center">
      <tr>
        <td height="30" colspan="3" align="left"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td width="10%" align="center"><%=vPMsg_ChkCode%></td>
            <td width="40%" nowrap="nowrap">&nbsp;&nbsp;
              <input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../');"/>
              <img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/></td>
            <td width="40%"><%=vJApp_mInp%><%=vPMsg_ChkCode%> <%=fMsgShow%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td align="center"><input type="button" name="button" id="button" value="<%=vJApp_BtnOK%>" onclick="ChkFData(document.fmMail)"/></td>
        <td align="center"><input type="reset" name="button2" id="button2" value="<%=vJApp_BtnNO%>" /></td>
        <td width="50%" height="30" align="left"><%=vJApp_mBot%>
        <input name="MD" type="hidden" id="MD" value="<%=MD%>" />
        <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
        </td>
      </tr>
    </table>
    <!--KillB-->
  </div>
</form>
<script type="text/JavaScript">
function ChkFData(fmName)
{
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fmName.InfSubj.value.length==0) 
   {   
     alert(" <%=vJApp_jsName%>"); 
     fmName.InfSubj.focus();
     eflag = 1; break;
   }
 if (fmName.LnkTel.value.length==0) 
   {   
     alert(" <%=vJApp_jsTel%>"); 
     fmName.LnkTel.focus();
     eflag = 1; break;
   }
 if (fmName.ObjSubj.value.length==0) 
   {   
     alert(" <%=vJApp_jsSubj%>"); 
     fmName.ObjSubj.focus();
     eflag = 1; break;
   }
 if (fmName.InfCont.value.length==0) 
   {   
     alert(" <%=vJApp_jsCont%>"); 
     fmName.InfCont.focus();
     eflag = 1; break;
   }
   // LnkTel JobName,InfRem1
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ ChkFSend(fmName); }
}
 
function ChkFSend(fmName)
{ 
  fmName.sysCont.value=SMailID.innerHTML;
  fmName.submit();
}
</script>
