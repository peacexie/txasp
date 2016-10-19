<!--#include file="../page/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Send Mail(Order)</title>
<link href="inc/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="apis/play/jsPlayer.js" type="text/javascript"></script>
</head>
<body>
<table align="center">
  <form id="fms01" name="fms02" method="post" action="../page/osearch.asp" target="_blank">
    <tr>
      <td>OrderNO</td>
      <td><textarea name="KeyWD" rows="4" class="fm" id="KeyWD" style="font-size:12px;width:150px; "></textarea></td>
    </tr>
    <tr>
      <td>Search</td>
      <td><input type="submit" name="button" id="button" style="width:120px;" value="Order Search"/></td>
    </tr>
    <tr>
      <td colspan="2">注意：可一次填入多个定单号，<br />
        如输入多个定单号，则一行一个；<br />
        一次最多可查询10个定单；<br />
        eg: <a href="ord_view.asp?ID=0BA8A839F7D008237A187B3T" target="_blank">108RB-JK82K-984</a>, <a href="ord_view.asp?ID=0BA89D325A3615FFKH8X25HV" target="_blank">8DA24S-GDW1K6</a>, <a href="ord_view.asp?ID=0BA73B5EBA162DSVN71DFRRF" target="_blank">Ord-0998H-1B8FR</a> 。</td>
    </tr>
  </form>
</table>
<!--#include file="../page/_head.asp"-->
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle"><%=MDName%></td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td align="right" class="SysTit03" height="30"><%=vPMsg_WSite%> <a href="../index.asp"><%=vPMsg_WHome%></a> &gt;&gt; <a href="../trade/">供求黄页</a> &gt;&gt; <a href="corp.asp?UsrID=<%=US%>" class="cRed"><%=MemSubj%></a> &gt;&gt; <%=MDName%> &gt;&gt;&nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><table cellspacing="0" cellpadding="0" width="99%" border="0">
        <tr>
          <td height="500" colspan="4" align="left" valign="top" style="padding:5px;"><!--Item Start-->
<!--#include file="../upfile/sys/pcfg/jmail.asp"--> 
<!--#include file="../upfile/sys/pcfg/link.asp"--> 
<!--#include file="../sadm/func2/func_obj.asp"-->
            <%

If verNow="2" Then
 vPOrd_mSubj = "Order"
 vPOrd_mClose = "Close"
 vPOrd_mPrint = "Print"
 vPOrd_mTime = "Time"
 vPOrd_mInp = " Input "
 vPOrd_mBot = " Please Input the Details, We will Contact you SOON! "
 vPOrd_Subj = "Subject"
 vPOrd_Cont = "Content"
 vPOrd_Name = "Name"
 vPOrd_Tel = "Tel"
 vPOrd_Order = "Order: "
 vPOrd_BtnOK = "OK Send"
 vPOrd_BtnNO = "Reset"
 vPOrd_MsgNull = "Error, Null Subject or Contenct!"
 vPOrd_MsgCode = "Check Code Error,Please Click the Code Picture And Send Again! "
 vPOrd_MsgOK = "Send OK!! \n We will Contact you SOON! "
 vPOrd_jsSubj = "Subject is Null!"
 vPOrd_jsCont = "Content is Null!"
 vPOrd_jsName = "Name is Null!"
Else
 vPOrd_mSubj = "定单提交"
 vPOrd_mClose = "关闭"
 vPOrd_mPrint = "打印"
 vPOrd_mTime = "时间"
 vPOrd_mInp = " 请输入 "
 vPOrd_mBot = " 请认真输入，提交资料后，我们尽快处理！ "
 vPOrd_Subj = "标题"
 vPOrd_Cont = "内容"
 vPOrd_Name = "姓名"
 vPOrd_Tel = "电话"
 vPOrd_Order = "订购: "
 vPOrd_BtnOK = "提交"
 vPOrd_BtnNO = "重填"
 vPOrd_MsgNull = "输入为空！请点击认证码图片，刷新认证码再提交一次！"
 vPOrd_MsgCode = "认证码错误！请点击认证码图片，刷新认证码再提交一次！ "
 vPOrd_MsgOK = "提交成功!如果有需要，我们会尽快联络您！"
 vPOrd_jsSubj = "主题不能为空!"
 vPOrd_jsCont = "内容不能为空!"
 vPOrd_jsName = "姓名不能为空!"
End If

ObjSubj = RequestS("ObjSubj","C",255) 
ObjID = RequestS("ObjID","C",48)

sSubj = RequestS("InfSubj","C",48)
sType = ObjSubj
sCont = RequestS("sysCont","C",240000)
'sPage = Request.Servervariables("HTTP_REFERER")
'rPage = Request("rPage")
send = Request("send")
ChkCode = Request("ChkCode")

If send="send" Then

  Call Chk_URL()	
  'Response.End()
  IP = Request.ServerVariables("REMOTE_HOST")
  KeyID = Get_AutoID(24) 
  ''//////////////////
      emSubj = sSubj
      'emSubj = "=?UTF-8?B?"&sSubj&"?="
      emCont = sCont
	  'emCont = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"&Cont
      Call Send_Mails(Config_Name,"xpigeon@163.com",emSubj,emCont,"gb2312")
      'Call Send_Mails(Config_Name,"xpigeon@163.com",emSubj,emCont,"utf-8")
   ''//////////////////
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
sql = sql& ", '" & rs_AutID(conn,"GboSend","KeyCode","Ord-","1","") &"'" 'ObjID
sql = sql& ", '" & ObjSubj &"'" 
sql = sql& ", '" & sSubj &"'" 
sql = sql& ", '" & sCont &"'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  If Trim(sSubj)="" Or Trim(sCont)="" Then
    fMsgShow = vPOrd_MsgNull
	PicReLoad = "PicReLoad('../');" 
	Response.Write js_Alert(fMsgShow,"Alert","-1")
  ElseIf Session("ChkCode")<>uCase(ChkCode) Then
    fMsgShow = vPOrd_MsgCode
	PicReLoad = "PicReLoad('../');" 
	Response.Write js_Alert(fMsgShow,"Alert","-1")
  Else
    Call rs_Dosql(conn,sql)
    Response.Write js_Alert(vPOrd_MsgOK,"Redir","pview.asp?KeyID="&ObjID&"") 
  End If
End If

'Response.Write send&sSubj&sCont

'InfCode = rs_AutID(conn,"GboSend","KeyCode","Ord-","1","")
ObjSubj = vPOrd_Order&Show_Form(ObjSubj)
ObjID = Show_Form(ObjID)
%>
            <form method="post" name='fmMail' action='?'>
              <input type="hidden" name='send' value="send" />
              <INPUT type=hidden name='sysSubj' value='[<%=vPOrd_mSubj%>]'>
              <INPUT type=hidden name='sysTo' value='xpigeon@163.com'>
              <INPUT type=hidden name='sysFrom' value='[<%=vPMsg_WName%>]'>
              <INPUT type=hidden name='sysCont'>
              <INPUT type=hidden name='sysEnd' value="Close">
              <div id="SMailID">
                <table width="640" border="0" align="center">
                  <tr>
                    <td align="center"><TABLE width='100%' border=0 align=center cellPadding=2 cellSpacing=0>
                        <TBODY>
                          <TR align=middle>
                            <th height="30" align="center" nowrap>[<%=vPMsg_WName%>]<%=vPOrd_mSubj%></th>
                            <TD width="30%" align="center" noWrap>&nbsp;</TD>
                            <TD width="10%" align="center" noWrap><span style="cursor:hand" onClick="window.close();" title="点击关闭窗口"><%=vPOrd_mClose%></span></TD>
                            <TD width="10%" align="center" noWrap><span style="cursor:hand" onClick="javascript:window.print();" title="点击打印本页"><%=vPOrd_mPrint%></span></TD>
                          </TR>
                        </TBODY>
                      </TABLE></td>
                  </tr>
                </table>
                <table width="640" border="0" align="center" cellpadding="3" cellspacing="1">
                  <tr>
                    <td width="10%" align="center" nowrap="nowrap"><%=vPOrd_Subj%></td>
                    <td align="left"><input name="InfSubj" type="text" id="InfSubj" value="<%=ObjSubj%>" size="48" maxlength="48" /></td>
                  </tr>
                  <tr>
                    <td align="center" nowrap="nowrap"><%=vPOrd_Cont%></td>
                    <td align="left"><textarea name="InfCont" cols="48" rows="8" id="InfCont"></textarea></td>
                  </tr>
                  <tr>
                    <td align="center" nowrap="nowrap"><%=vPOrd_Name%></td>
                    <td align="left"><input name="ObjSubj" type="text" id="ObjSubj" size="18" maxlength="24" />
                      <%=vPOrd_Tel%>
                      <input name="sTel" type="text" id="sTel" size="18" maxlength="24" /></td>
                  </tr>
                  <tr>
                    <td align="center" nowrap="nowrap">E-Mail</td>
                    <td align="left"><input name="sEmail" type="text" id="sEmail" size="18" maxlength="60" />
                      ＱＱ
                      <input name="sQQ" type="text" id="sQQ" size="18" maxlength="24" /></td>
                  </tr>
                </table>
                <table width="640" border="0" align="center">
                  <tr>
                    <td width="50%" align="center"><%=vPOrd_mTime%>: <%=Now()%></td>
                    <td width="50%" align="right">[<%=vPMsg_WName%>]<%=vPOrd_mSubj%></td>
                  </tr>
                </table>
                <!--KillA-->
                <table width="640" border="0" align="center">
                  <tr>
                    <td height="30" colspan="3" align="left"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
                        <tr>
                          <td width="10%" align="center"><%=vPMsg_ChkCode%></td>
                          <td width="40%" align="left" nowrap="nowrap"><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../');"/>
                            <img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/></td>
                          <td width="40%"><%=vPOrd_mInp%><%=vPMsg_ChkCode%> <%=fMsgShow%></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr>
                    <td width="30%" align="center"><input type="button" name="button" id="button" value="<%=vPOrd_BtnOK%>" onclick="ChkFData(document.fmMail)"/></td>
                    <td width="20%" align="center"><input type="reset" name="button2" id="button2" value="<%=vPOrd_BtnNO%>" /></td>
                    <td width="50%" height="30" align="left"><%=vPOrd_mBot%>
                      <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
                      <input name="ObjID" type="hidden" id="ObjID" value="<%=ObjID%>" />
                      <input name="ObjSubj2" type="hidden" id="ObjSubj2" value="<%=ObjSubj%>" />
                      <input name="ObjUrl" type="hidden" id="ObjUrl" value="<%=ObjUrl%>" /></td>
                  </tr>
                </table>
                <!--KillB-->
              </div>
            </form>
            <script language="JavaScript" type="text/JavaScript">
function ChkFData(fmName)
{
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fmName.InfSubj.value.length==0) 
   {   
     alert(" <%=vPOrd_jsSubj%>"); 
     fmName.InfSubj.focus();
     eflag = 1; break;
   }
 if (fmName.ObjSubj.value.length==0) 
   {   
     alert(" <%=vPOrd_jsName%>"); 
     fmName.ObjSubj.focus();
     eflag = 1; break;
   }
 if (fmName.InfCont.value.length==0) 
   {   
     alert(" <%=vPOrd_jsCont%>"); 
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
            <!--Item End--></td>
        </tr>
      </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240"><!--Side Start-->
      <!--Side End--></td>
  </tr>
</table>
<!--#include file="../page/_foot.asp"-->
</body>
</html>
