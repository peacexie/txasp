<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%

  rfUrl1 = Chk_URL3(bbsPath&"badm/bbs_list.asp")
  rfUrl2 = Chk_URL3(bbsPath&"buser.asp")
  rfUrl3 = Chk_URL3(bbsPath&"bview.asp")
  rfUrl4 = Chk_URL3(bbsPath&"bedit.asp")
  'Response.Write rfUrl1&rfUrl2&rfUrl3&rfUrl4
  If inStr(rfUrl1&rfUrl2&rfUrl3&rfUrl4,"OK")>0 Then
  Else
    Response.Write "rfUrl1~rfUrl4"
	Response.End()
  End If

send = Request("send")
ID = RequestS("ID",C,48)
eID = RequestS("eID",C,48)

'Response.Write "f-03:"&Chk_Perm1("ModBBS","")
If rfUrl1="OK" AND Chk_Perm1("ModBBS","") Then
ElseIf rfUrl2="OK" AND (bbsChkUser(eID,bbsUser,"Self") OR bbsChkUser(eID,bbsUser,"Master")) Then
ElseIf rfUrl3="OK" AND bbsChkUser(eID,bbsUser,"Self") Then
ElseIf rfUrl4="OK" Then
Else
  Response.Write "rfUrl-x4"
  Response.End()
End If


If send = "send" Then

  InfSubj = RequestS("InfSubj",C,255)
  InfCont = RequestS("InfCont",C,48000) 
  InfSubj = Chr_Fil2(InfSubj)
  InfCont = Chr_Fil2(InfCont)
If Config_Cont="DB" Then
  xxxCont = InfCont
Else
  xxxCont = ""
End If
sql = " UPDATE [BBSInfo] SET " 
sql = sql& " InfSubj = '" & InfSubj &"'" 
sql = sql& ",InfType = '" & RequestS("InfType",C,255) &"'" 
sql = sql& ",InfCont = '" & xxxCont &"'" 
sql = sql& ",ImgName = '" & Session("ImgList")&"" &"'"
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & bbsUser &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&eID&"' "
  'If Session("UsrID")&""="" Then
    'sql2 = "SELECT TypID FROM [WebTyps] WHERE TypID='"&RequestF("TypOld",C,255)&"' AND TypNam2 LIKE '"&Session("MemID")&"' "
    'If rs_Val("",sql2)="" Then
      'sql = sql& " AND LnkName='"&Session("MemID")&"' "    
	'End If
  'End If
Call rs_DoSql(conn,sql)
KeyID = eID
upPath = upRoot&Replace(KeyID,"-","/")&"/" 
Call add_sfFile()
Session("ImgList")=""
Response.Redirect "bview.asp?ID="&ID
  Msg = "修改成功! "&msgPW
  Response.Write sql
End If  

rs.Open "SELECT * FROM [BBSInfo] WHERE KeyID='"&eID&"'",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = Show_Form(rs("InfType"))
InfSubj = Show_Form(rs("InfSubj"))
InfCont = rs("InfCont")
LnkFace = Show_Form(rs("LnkFace"))
LnkName = Show_Form(rs("LnkName"))
LnkEmail = Show_Form(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
SetSAdm = rs("SetSAdm")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
ImgName = rs("ImgName")
Session("ImgList")=ImgName&""
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
xxxCont = rs("InfCont") 
InfCont = Show_sfRead(eID,"/fcont.htm")
Session("KeyID") = KeyID
end if 
rs.Close()
  'set rs = nothing

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>帖子修改 -<%=bbsName%></title>
<link rel="stylesheet" type="text/css" href="bimg/style.css">
<script type="text/javascript" charset="utf-8" src="../inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="../inc/home/jsInfo.js"></script>
<script type="text/javascript" charset="utf-8" src="../smod/file/edt_api.asp?edtAct=mainLoad"></script>
</head>
<body>
<!--#include file="_itop.asp"-->
<div align="center" style="width:980px; height:auto; margin:auto; background-color:#F2F6FB">
  <table width="980" border="0" align="center" cellpadding="1" cellspacing="8">
    <tr>
      <td align="left" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; <a href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>"><%=TName%></a> &gt;&gt; 编辑帖子 &gt;&gt;</td>
      <td width="50%" align="right" bgcolor="#FFFFFF">&nbsp;<a href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>">返回列表</a>&nbsp;</td>
    </tr>
  </table>
  <table width="720"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="fm01" id="fm01" action="?" method="post">
      <tr>
        <td align="center" bgcolor="#F2F6FB">主题</td>
        <td align="left" bgcolor="#F2F6FB"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="36" maxlength="60" /></td>
        <td rowspan="3" align="center" valign="top" bgcolor="#F2F6FB"><img src="<%=LnkFace%>" width="75" height="85" id="LnkFac2"/></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">类别</td>
        <td align="left" bgcolor="#F2F6FB" ><select name="InfType" id="InfType" send="send">
            <option value="<%=InfType%>">(默认)</option>
            <%If Session("UsrID")&""<>"" Then%>
            <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
            <%End If%>
          </select>
          &nbsp;&nbsp;头像
          <select name="LnkFace" size=1 id="LnkFace" xonChange=document.all.LnkFac2.src=document.all.LnkFace.value>
            <option value='<%=LnkFace%>'>[ <%=Replace(LnkFace&"","../img/Head/","")%> ]</option>
          </select>
          <input name="TypOld" type="hidden" id="TypOld" value="<%=TypID%>"></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">姓名</td>
        <td align="left" bgcolor="#F2F6FB"><input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="36"></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">内容
          </td>
        <td colspan="2" align="left" bgcolor="#F2F6FB">
        
<textarea id="InfCont" name="InfCont" style="width:580px;height:360px;visibility:hidden;display:none"><%=Show_Form(InfCont)%></textarea>
<script type="text/javascript" charset="utf-8" src="../smod/file/edt_api.asp?edtID=EditID01&edtCont=InfCont"></script>

        </td>
      </tr>
      <tr>
        <td width="10%" bgcolor="#F2F6FB">&nbsp;</td>
        <td colspan="2" align="left" bgcolor="#F2F6FB"><input type="button" name="Button" value="提交" onClick="chkData()">
          <input type="reset" name="Submit2" value="重置">
          <input name="send" type="hidden" id="send" value="send">
          <input name="ChkCode" type="hidden" id="ChkCode" value="<%=chkCode%>">
          <input name="ID" type="hidden" id="ID" value="<%=ID%>">
          <input name="eID" type="hidden" id="eID" value="<%=eID%>">
          <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
          <%=fMsg%></td>
      </tr>
    </form>
  </table>
  <div style="line-height:8px;">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td height="24" align="left" bgcolor="#FFFFFF">发贴说明: </td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<script type="text/javascript">
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 //if (document.fm01.InfCont.value.length==0)
 document.fm01.InfCont.value = apiGetValEditID01();
 if (document.fm01.InfCont.value.length==0)
   {   
     alert(" 内容 不能为空！"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length>=24000) 
   {   
     alert(" 内容 不能超过 24 K字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
