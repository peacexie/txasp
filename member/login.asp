<!--#include file="_config.asp"-->
<%				
'Response.Write vMem_Btn01
'Session.Timeout = 1440 

verCSet = Request.Cookies("PeaceCSet")
If verCSet="Big5" And verMemb="1" Then
  verMemb="0"
End If
'Response.Write verMemb


Call sf_Guard()	  
send = Request("send")
goPage = Request("goPage") ' <a href="../member/login.asp?goPage=goRef" target="_blank">登陆再[动作]</a>
If goPage="goRef" Then
  goPage = Request.Servervariables("HTTP_REFERER")
End If


If send = "send" Then
  Call Chk_URL()
  MemID = Trim(RequestS(get30Tab(1),3,128))
  MemPW = Trim(RequestS(get30Tab(3),3,48))
  	Session("MemID") = ""
    Session("MemPerm") = ""
  If Len(MemID&"")=0 Then MemID=" "
  If Len(MemPW&"")=0 Then MemPW=" "
  MemPWE= MD5_Mem(MemPW&MemID) 
  MemPWA= RequestS("MemPWA",3,512)
  ChkCode = uCase(Request("ChkCode"))
   Msg = ""
   
 If Session("ChkMember")&""="" OR Session("ChkMember")<>ChkCode Then
   msg = vMem_MsgB1&" "&Session("ChkMember")&":"&ChkCode
 
 Elseif Get_A30Chk(get30Time,get30TSN,60,"")<>"" Then
    msg = "错误！超时错误 或 安全认证错误"

 Else 
   if Session(UsrPStr)&""<>"" AND Len(RequestS("MemPWA",3,512))>16 Then
     whPW = "MemPW='"&MemPWA&"'"
     goPage = "[Admin]" '***
   else
     whPW = "MemPW='"&MemPWE&"'" 
   end if
   sql = "SELECT MemID,MemType,MemGrade,MemFlag FROM [Member"&Mem_aMemb&"] WHERE MemID='"&MemID&"' AND "&whPW&""
   set rs = server.CreateObject("adodb.recordset")
   rs.Open sql,conn,1,1
   if not rs.EOF then
	  MemType = rs("MemType")
	  MemGrade = rs("MemGrade")
	  MemFlag = rs("MemFlag")
	'/////////////////// Check for 子站
	If Len(Config_Path)>5 And lCase(Left(Config_Path,3))="/u/" Then
     Session("Pub_Subs") = Config_Path
	End If
	'//////////////////////////////////
	  If MemFlag="Y" Or goPage="[Admin]" Then
	    Session("MemID")  = MemID
		Session("MemPerm") = "{("&MemID&":"&MemGrade&"("&MemType&");}"&gMemPerm(conn,MemGrade)&"{FileInfo,FileEditor,FileView,FileAdmin}" 
	    Msg = "Welcom !"
	    If goPage<>"[Admin]" Then
	      sql = "UPDATE [Member"&Mem_aMemb&"] SET LogETime='"&Now()&"',LogEditIP='"&Get_CIP()&"' WHERE MemID='"&MemID&"' "
	      Call rs_DoSql(conn,sql)
	      sql = "UPDATE [OrdItem] SET LogAUser='"&MemID&"' WHERE LogAUser='("&Session.SessionID&")' "
	      Call rs_DoSql(conn,sql)
	    End If
	  Else
	    Msg = vMem_MsgB2
	  End If
	else
	  Msg = vMem_MsgB3
	end if
    rs.Close()
    set rs = nothing
	'Response.Write goPage :Response.End()
	if Msg = "Welcom !" then
	   if goPage = "[Admin]" Then
		 Response.Redirect "../inc/mem_inc/mem_main.asp"
       elseif goPage <> "" Then
	     'Session("goPage") = ""
	     Response.Redirect goPage
	   else
		 Response.Redirect "../inc/mem_inc/mem_main.asp?verMemb="&verMemb&""
	   end if  
	end if
 end if
ElseIf send = "del" Then
	Session("MemID") = ""
    Session("MemPerm") = ""
    Msg = "-=> "&vMem_MsgB4
ElseIf send = "TimOut" Then
    Msg = "-=> "&vMem_MsgB5
ElseIf send = "out" Then
    Msg = "-=> "&vMem_MsgB6
	Session("MemID") = ""
    Session("MemPerm") = ""
	if goPage <> "" Then
	  Response.Redirect goPage
	end if
ElseIf send = "NoLogin" Then
    Msg = "-=> "&vMem_MsgB7
ElseIf send = "NoPerm" Then
    Msg = "-=> "&vMem_MsgB8
Else ' NoLogin
	Msg = "-=> "&vMem_MsgB9
	'Msg = Msg&vbcrlf&"<!--"&vbcrlf&EncOffset&"-->"&vbcrlf
    'Response.Write "<!--出错啦~&*%#!"&"<br>"&EncOffset&"[PeaceXieYS]-->"
End If


'// Test 
urlEvent = Server.HTMLEncode("class='lgnBtn1 lgnInpt' style='width:150px;' xtabindex=1") 
urlParas = "?act=setInput&n="&Get_IDEnc(app30Tab(1),0)&"&s=20&m=64&v="&MemID&"&e="&urlEvent&url30Par&""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE><%=vMem_PLogin%>-<%=vPMsg_WName%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<link rel="stylesheet" type="text/css" href="../inc/adm_img/login.css"/>
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</HEAD>
<BODY>
<%TPName=vMem_PLogin%>
<!--#include file="_ftop2.asp"-->
      <table width="100%" border="0" cellpadding="2" cellspacing="0" style="margin:auto">
        <tr>
          <td style="padding:12px 0px"><table width="450" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td colspan="2" class="lgnCnr5"></td>
                <td class="lgnCnr6"></td>
              </tr>
              <tr>
                <td colspan="3" style="padding:0px 2px 0px 4px"><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="#1C78B8">
                    <tr>
                      <td align="right" class="lgnComp"><%=vMem_PLogin%></td>
                      <td align="right" class="lgnLogoSite">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <%If Session("MemID")&""<>"" Then%>
              <tr>
                <td class="lgnSid1">&nbsp;</td>
                <td valign="top" class="lgnCont"><table width="320" border="0" align="center" cellpadding="2" cellspacing="1">
                    <%If verMemb="2" Then%>
                    <tr>
                      <td nowrap="nowrap" class="SysCont"><font color="#0000FF" style="font-weight:bold; "><%=Session("MemID")%></font>, Welcome
                        to use this system! <br />
                        &nbsp;&nbsp; Please <a href="../inc/mem_inc/mem_main.asp?verMemb=<%=verMemb%>"><font color="#0000FF">Into Member Center</font></a> Or <a href="login.asp?send=out&amp;verMemb=<%=verMemb%>"><font color="#0000FF">Logout</font></a></td>
                    </tr>
                    <tr>
                      <td nowrap="nowrap" class="SysCont">Not:<%=Now()%><br />
                        &nbsp; IP:<%=Request.Servervariables("REMOTE_ADDR")%></td>
                    </tr>
                    <%Else%>
                    <tr>
                      <td nowrap="nowrap" class="SysCont"><font color="#0000FF" style="font-weight:bold; "><%=Session("MemID")%></font> 您好, <br />
                        您已经登入系统! <br />
                        &nbsp;&nbsp; 请 <a href="../inc/mem_inc/mem_main.asp?verMemb=<%=verMemb%>"><font color="#0000FF">进入后台</font></a> 或 <a href="login.asp?send=out&amp;verMemb=<%=verMemb%>"><font color="#0000FF">安全登出</font></a></td>
                    </tr>
                    <tr>
                      <td nowrap="nowrap" class="SysCont">当前时间:<%=Now()%><br />
                        当前ＩＰ:<%=Request.Servervariables("REMOTE_ADDR")%></td>
                    </tr>
                    <%End If%>
                  </table></td>
                <td class="lgnSid2">&nbsp;</td>
              </tr>
              <%Else%>
              <tr>
                <td class="lgnSid1">&nbsp;</td>
                <td valign="top" class="lgnCont"><table width="100%" border="0" 
                        align="center" cellpadding="2" cellspacing="1" class="lgnTab">
                    <form action="?" method="post" name="ff" id="ff">
                      <tr>
                        <td align="right" nowrap="nowrap">&nbsp;<%=vMem_MemID%>:</td>
                        <td height="27" align="left" nowrap="nowrap">
                        <script src="../inc/form/form_js.asp<%=urlParas%>" type='text/javascript'></script>
                        <input name="send" type="hidden" id="send" value="send" /></td>
                      </tr>
                      <tr>
                        <td align="right" nowrap="nowrap">&nbsp;<%=vMem_MemPW%>:</td>
                        <td height="27" align="left" nowrap="nowrap"><input name="<%=app30Tab(3)%>" type="password" class="lgnBtn1 lgnInpt" id="<%=app30Tab(3)%>" size="20" maxlength="24"  style="width:150px; " />
                        <input name="goPage" type="hidden" id="goPage" value="<%=goPage%>" /></td>
                      </tr>
                      <tr>
                        <td align="right" nowrap="nowrap"><%=vPMsg_ChkCode%>:</td>
                        <td align="left" nowrap="nowrap"><input name="ChkCode" type="text" id="ChkCode" size="6" class="lgnInpt" maxlength="12" />
                          <img src="../sadm/pcode/img_frnd.asp?xConfig_PCode=hij&Config_PSess=ChkMember" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../','','ChkMember')"/></td>
                      </tr>
                      <tr>
                        <td width="25%" align="left" nowrap="nowrap"><input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>" /></td>
                        <td nowrap="nowrap"><input type="submit" name="Submit3" value="<%=vMem_Btn11%>" class="SysBtnM1" />
                          &nbsp;
                          <input id="Submit1" type="reset" value="<%=vMem_Btn12%>" name="Submit1" class="SysBtnM1" /></td>
                      </tr>
                      
        <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
        <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
        <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
                      
                    </form>
                  </table>
                  <div class="line05">&nbsp; </div>
                  <table width="100%" 
            border="0" cellpadding="2" cellspacing="0" class="lgnRBox">
                    <tr align="center">
                      <td colspan="5" align="left" nowrap="nowrap" class="lgnMsg"><%=vMem_MsgA0%> <%=msg%></td>
                    </tr>
                    <tr>
                      <td colspan="5" align="left" valign="top" nowrap="nowrap" class="lgnRead13">&middot; <%=vMem_MsgA3%><br />
                        &middot; <%=vMem_MsgA1%> <br />
                        &middot; <%=vMem_MsgA2%> <a href="?goPage=<%=goPage%>&verMemb=<%=verMemb%>"><span class="col00F"><%=vMem_Btn13%></span></a> <%=vMem_Btn11%>；</td>
                    </tr>
                    <tr>
                      <td width="80">&nbsp;</td>
                      <td align="center" nowrap="nowrap" class="lgnMsg"><a href="mu_<%=AppRand12%>.asp?goPage=<%=goPage%>&verMemb=<%=verMemb%>&act=Start<%=url30Par%>"><span class="colF00"><%=vMem_Btn01%></span></a></td>
                      <td width="50">&nbsp;</td>
                      <td align="center" nowrap="nowrap" class="lgnMsg"><a href="get_pw.asp?goPage=<%=goPage%>&verMemb=<%=verMemb%><%=url30Par%>"><span class="colF00"><%=vMem_Btn02%></span></a></td>
                      <td width="80">&nbsp;</td>
                    </tr>
                    <tr align="center">
                      <td colspan="5" align="left" nowrap="nowrap" class="line05">&nbsp;</td>
                    </tr>
                  </table></td>
                <td class="lgnSid2">&nbsp;</td>
              </tr>
              <%End If%>
              <tr>
                <td colspan="2" class="lgnCnr3"></td>
                <td class="lgnCnr4"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <!--#include file="_fbot2.asp"-->
</BODY>
</HTML>
