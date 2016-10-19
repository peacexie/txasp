<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<!--#include file="../sadm/func1/md5_func.asp"-->

<%		

Call ChkSpider(1)
UsrType = RequestS("UsrType",3,84) 'Docs,BBS
DefType = "Docs" 'DefDir  = ""
If UsrType="" Then UsrType=DefType

MDCodes = "BBS;Docs;Down"
MDNames = "内部论坛;内部公文;内部下载"
MDHomes = "../bbs/;../doc/;../page/down.asp"

MDName = Get_SOpt(MDCodes,MDNames,UsrType,"Val")

send = Request("send")
goPage = Request("goPage") 
If goPage="goRef" Then
  goPage = Request.Servervariables("HTTP_REFERER")
End If
'Response.Write send&":"&goPage


Dim sys27_Rnd(10)
If send = "send" Then
  sys27_RVal = Request(App27Random)&""
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If
Else
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If



UsrID   = RequestF("UsrID"&sys27_Rnd(1),3,96)
UsrPW   = RequestF("UsrPW"&sys27_Rnd(2),3,96)

If send = "send" Then
Call Chk_URL()

 if Len(UsrID&"")<2 or Len(UsrPW&"")<5 then
   Msg = "-=> 错误 ! 请输入帐号密码! "
 elseif Session("ChkCode")<>uCase(Request("ChkCode")) then
   Msg = "-=> 认证码错误! ["&Session("ChkCode")&"]-["&uCase(Request("ChkCode"))&"]"
 'elseif UsrType <> "" then 
 else  
   eStr = MD5_Adm(UsrPW&UsrID)
   sql  = "SELECT * FROM [AdmUser"&Adm_aUser&"] "
   sql = sql& " WHERE UsrID='" &UsrID&"' AND UsrPW='"&eStr&"' "
   sql = sql& " AND (UsrType LIKE 'Inn%')  " 'AND UsrExp>#"&Date()&"# 
   set rs = Server.CreateObject("adodb.recordset") ': Response.Write sql
   rs.Open sql,conn,1,3
   if not rs.EOF then
	 If rs("UsrType")="InnStop" Then
	  Msg = "禁止登陆 !"
	 Else
	  Session("InnID") = UsrID
	  Session("InnPerm") = "{(MemFEditor);(MemFUpload)}"&rs("UsrPerm")
	  Session("UsrLTime") = rs("UsrLTime")
      rs("UsrLogIP") = Get_CIP()
	  rs("UsrLTime") = Now()
	  rs.UPDATE()
	  Msg = "Welcom !"
	 End If
	else
	  Msg = "-=> 帐号不存在; 密码错误; 帐号禁用 或帐号到期 ! "
	end if
	rs.Close()
    set rs = nothing
	'/////////////////// Check for 子站
	If Len(Config_Path)>5 And lCase(Left(Config_Path,3))="/u/" Then
     Session("Pub_Subs") = Config_Path
	End If
	'//////////////////////////////////
	if Msg = "Welcom !" then
		  Call Add_Log(conn,Session("InnID"),MDName,"[login]",Msg)

  		  MDHome = Get_SOpt(MDCodes,MDHomes,UsrType,"Val")
		  If goPage&""<>"" Then
			Response.Redirect goPage
		  ElseIf MDHome&""<>"" Then
		    Response.Redirect MDHome
		  Else
		    Msg = "登陆成功！"
		  End If
		  
	end if
    
 end if
 
 
ElseIf send = "TimOut" Then
    Msg = "-=> 超时，请重新登陆! "
ElseIf send = "out" Then
    Msg = "-=> 已经安全退出!"
	Session("InnID") = ""
	Session("InnPerm") = ""
ElseIf send = "NoLogin" Then
    Msg = "-=> 请登陆!"
ElseIf send = "NoPerm" Then
    Msg = "-=> 没有权限，请联系管理员！"
ElseIf send = "UsrID_ERROR" Then
    Msg = "-=> 错误!请登陆！"
ElseIf send = "getPW" Then 
    Msg = "-=> 取回密码 成功！谢谢使用!<br>请登陆系统修改您好记忆的密码！"
Else ' NoLogin
    Msg = "-=> 请输入帐号密码！" 
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE>Demo-<%=Config_Name%></TITLE>
<META http-equiv=Pragma content=no-cache>
<meta name="robots" content="noindex, nofollow">
<meta name="robots" content="noarchive">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="../inc/adm_img/login.css"/>
</HEAD>
<BODY>
<div class="line50">&nbsp;</div>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:auto">
  <tr>
    <td colspan="2" class="lgnCnr5"></td>
    <td class="lgnCnr6"></td>
  </tr>
  <tr>
    <td colspan="3" style="padding:0px 2px 0px 4px"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#1C78B8" style="border:1px solid #1C78B8">
        <tr>
          <td align="left" valign="top" class="lgnLogoMember"><div class="lgnMod"><%=MDName%></div></td>
          <td align="center" class="lgnComp" style="overflow:hidden;">会员登陆<%=xPub_CName%></td>
          <td align="right" class="lgnLogoSite">&nbsp;</td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td class="lgnSid1">&nbsp;</td>
    <td valign="top" class="lgnCont"><table width="100%" class="lgnTab" 
            border="0" cellpadding="1" cellspacing="1">
        <form action="?" method="post" name="fm1<%=sys27_RVal%>" id="fm1<%=sys27_RVal%>" >
          <tr>
            <td align="right" valign="bottom" nowrap="nowrap">&nbsp;用户名</td>
            <td valign="bottom" nowrap="nowrap"><input name="UsrID<%=sys27_Rnd(1)%>" id="UsrID<%=sys27_Rnd(1)%>" 
                   class="lgnBtn1 lgnInpt"  
                  tabindex="1" value="<%=Request("UsrID")%>" size="30" maxlength="16" />
              <input name="send" type="hidden" id="send" value="send" /></td>
          </tr>
          <tr>
            <td align="right" valign="bottom" nowrap="nowrap">密　码</td>
            <td valign="bottom" nowrap="nowrap"><input 
              name="UsrPW<%=sys27_Rnd(2)%>" type="password" id="UsrPW<%=sys27_Rnd(2)%>" class="lgnBtn1 lgnInpt"  
                  tabindex="2" size="30" maxlength="16" /></td>
          </tr>
          <tr>
            <td align="right" valign="bottom" nowrap="nowrap">模　块</td>
            <td valign="bottom" nowrap="nowrap"><select name="UsrType" id="UsrType" class="lgnBtn1">
                <option value="">[---请选择----]</option>
                <%=Get_SOpt(MDCodes,MDNames,UsrType,"")%>
              </select></td>
          </tr>
          <tr>
            <td align="right" nowrap="nowrap">认证码</td>
            <td nowrap="nowrap"><input name="ChkCode" type="text" id="ChkCode" size="6" class="lgnInpt" maxlength="12" />
              <img src="../sadm/pcode/img_frnd.asp?xConfig_PCode=hij" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')" ('../','hij') />
              <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>" /></td>
          </tr>
          <tr>
            <td width="25%" align="center" nowrap="nowrap"><span class="col666"><a href="../" target="_blank">首页</a></span></td>
            <td nowrap="nowrap"><input type="submit" name="Submit" value="提 交" class="lgnBtn2" />
              <span class="lgnBtnSp">&nbsp;</span>
              <input type="submit" name="Submit2" value="提 交" class="lgnBtn2" />
              <input name="goPage" type="hidden" id="goPage" value="<%=goPage%>" /></td>
          </tr>
        </form>
      </table>
      <div class="line05"> &nbsp; </div>
      <table width="100%" 
            border="0" cellpadding="2" cellspacing="0">
        <tr align="center">
          <td colspan="2" align="left" nowrap="nowrap" class="lgnMsg">提示: <%=Msg%></td>
        </tr>
        <tr>
          <td colspan="2" align="left" valign="top" nowrap="nowrap" class="lgnRBox"><p class="lgnRead lgnLogo-Admin">&middot;  推荐请使用IE7,IE8浏览器，其它浏览器，请自行设置好；<br />
              &middot; <a href="../tools/help/xhelp.asp" target="_blank"><span class="colF0F">第一次使用，请查看 《后台管理帮助文件》&gt;&gt;&gt;</span></a>；<br />
              &middot; 帐号密码区分大小写;帐号&gt;=2位,密码&gt;=5位；<br />
              &middot; 认证码不分大小写,如错误,请<a href="?"><span class="col00F">刷新</span></a>登陆；<br />
              &middot; 技术支持: <a href="http://www.dg.gd.cn" target="_blank">东莞网(www.dg.gd.cn)</a>。</p></td>
        </tr>
      </table></td>
    <td class="lgnSid2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" class="lgnCnr3"></td>
    <td class="lgnCnr4"></td>
  </tr>
</table>
<script type="text/javascript">if(top.location!==self.location){top.location=self.location;}</script>
</BODY>
</HTML>

