<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%

TypID = RequestS("TypID","C",48)
TypLay = RequestS("TypLay","C",255) 
aParent = Split(TypLay,";")
TypParent = aParent(0)
TName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&TypID&"'")
send = Request("send")

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

if send="send" then
  Call Chk_URL()
  If Len(Session("KeyID"))<15 Then 
    KeyID = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
  Else
    KeyID = Session("KeyID")
  End If
  upPath = upRoot&Replace(KeyID,"-","/")&"/" 
  InfSubj = RequestS("InfSubj"&sys27_Rnd(1),C,255)
  InfCont = RequestS("InfCont"&sys27_Rnd(2),C,48000) 
  If bbsUser&""="" Then
   InfCont = Show_Text(InfCont)
  End If
  '//////////////////////////
  fSubj = Chr_Fil2(InfSubj)
  fCont = Chr_Fil2(InfCont)
  SetShow = "Y"
  fMsgShow = "感谢您发帖!"
  If SwhBBSChk="N" Then 'rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='SWRemark'")="N" Then
    SetShow = "N"
    fMsgShow = "感谢您发帖,我们回尽快审核!"
  ElseIf fSubj<>InfSubj or fCont<>InfCont then
    SetShow = "N"
    fMsgShow = "感谢您发帖,帖子含有关键字:\n["&fSubj&fCont&"]需要审核!"
  Else
  End If
  InfSubj = fSubj
  InfCont = fCont
  If Config_Cont="DB" Then
    xxxCont = InfCont
  Else
    xxxCont = ""
  End If
  '//////////////////////////
  ChkCode = uCase(RequestS("ChkCode","C",255)) 

sql = " INSERT INTO [BBSInfo] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyRe" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyFlag" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", LnkName" 
sql = sql& ", LnkFace" 
sql = sql& ", LnkEmail" 
sql = sql& ", SetRead" 
sql = sql& ", SetShow" 
sql = sql& ", SetSAdm" 
sql = sql& ", SetHot" 
sql = sql& ", SetTop" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ", LogEditIP" 
sql = sql& ", LogEUser" 
sql = sql& ", LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", ''"  'KeyRe
sql = sql& ", '"&ModID&"'" 
sql = sql& ", 'Y'" 'KeyFlag
sql = sql& ", '" & RequestS("InfType",C,255) &"'"
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & RequestS("LnkName"&sys27_Rnd(4),C,24) &"'" 
sql = sql& ", '" & RequestS("LnkFace",C,48) &"'" 
sql = sql& ", '" & RequestS("LnkEmail",C,255) &"'" 
sql = sql& ", 0" 
sql = sql& ", '" & SetShow & "'" ' SetShow
sql = sql& ", 'Y'"  
sql = sql& ", 'N'" 
sql = sql& ", '888'" 
sql = sql& ", '" & Session("ImgList")&"" &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & bbsUser &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '-'" 
sql = sql& ", '-'" 
sql = sql& ", '1900-12-31'" 
sql = sql& ")"

 If Len(InfSubj)>0 And Len(InfCont)>0 And Session("ChkCode")=ChkCode And Len(ChkCode)>1 Then
  Call rs_DoSql(conn,sql) 
  Call add_sfFile()
  Session("ImgList")=""
  Response.Write js_Alert(fMsgShow,"Redir","blist.asp?TypID="&TypID&"&TypLay="&TypLay&"") 
 Else
  fMsg = "帖子发布失败,请重新提交！\n可能认证码错误或内容为空！" ': Response.Write sql  
  Response.Write js_Alert(fMsg,"Redir","badd.asp?TypID="&TypID&"&TypLay="&TypLay&"") 
 End If

else
  Session("KeyID") = rs_AutID(conn,"BBSInfo","KeyID","dtbbs","1","")
end if

'Response.Write fMsg&"<br>"&InfSubj&"<br>"&InfCont&"<br>"&chkCod2&"<br>"&Session("ChkCode")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=bbsName%></title>
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
      <td align="left" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; <a href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>"><%=TName%></a> &gt;&gt; 发帖 &gt;&gt;</td>
      <td width="50%" align="right" bgcolor="#FFFFFF">&nbsp;<a href="blist.asp?TypID=<%=TypID%>&TypLay=<%=TypLay%>">返回列表</a>&nbsp;</td>
    </tr>
  </table>
  <!--AAAA-->
  <fieldset style="padding:3px; width:950px;">
  <legend> &nbsp; *贴子发布： </legend>
  <table width="720"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="fm01" id="fm01" action="?" method="post">
      <tr>
        <td align="center" bgcolor="#F2F6FB">主 题</td>
        <td align="left" bgcolor="#F2F6FB"><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="36" maxlength="60"></td>
        <td rowspan="3" align="center" valign="top" bgcolor="#F2F6FB"><img src="../img/Head/1.gif" width="75" height="85" id="LnkFac2"/></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">类 别</td>
        <td align="left" bgcolor="#F2F6FB" ><select name="InfType" id="InfType" send="send">
            <%=Get_rsOpt(conn,"SELECT TypLayer,TypName FROM WebTyps WHERE TypMod='"&ModID&"' AND TypLayer LIKE '"&TypParent&"%' AND TypDeep=2 ORDER BY TypLayer",TypLay)%>
          </select>
          &nbsp;&nbsp;头像
          <select name="LnkFace" size=1 id="LnkFace" onChange=document.all.LnkFac2.src=document.all.LnkFace.value>
            <option value=../img/Head/1.gif>Head1</option>
            <option value=../img/Head/2.gif>Head2</option>
            <option value=../img/Head/3.gif>Head3</option>
            <option value=../img/Head/4.gif>Head4</option>
            <option value=../img/Head/5.gif>Head5</option>
            <option value=../img/Head/6.gif>Head6</option>
            <option value=../img/Head/7.gif>Head7</option>
            <option value=../img/Head/8.gif>Head8</option>
            <option value=../img/Head/9.gif>Head9</option>
            <option value=../img/Head/10.gif>Head10</option>
            <option value=../img/Head/11.gif>Head11</option>
            <option value=../img/Head/12.gif>Head12</option>
            <option value=../img/Head/13.gif>Head13</option>
            <option value=../img/Head/14.gif>Head14</option>
            <option value=../img/Head/15.gif>Head15</option>
            <option value=../img/Head/16.gif>Head16</option>
            <option value=../img/Head/17.gif>Head17</option>
            <option value=../img/Head/18.gif>Head18</option>
            <option value=../img/Head/19.gif>Head19</option>
            <option value=../img/Head/20.gif>Head20</option>
          </select></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">姓 名</td>
        <td align="left" bgcolor="#F2F6FB"><input name="LnkName<%=sys27_Rnd(4)%>" type="text" id="LnkName<%=sys27_Rnd(4)%>" value="<%=bbsUser%>" size="36"></td>
      </tr>
      <%If bbsUser&""<>"" Then%>
      <tr>
        <td align="center" bgcolor="#F2F6FB">内 容
        </td>
        <td colspan="2" align="left" bgcolor="#F2F6FB">
        
<textarea id="InfCont<%=sys27_Rnd(2)%>" name="InfCont<%=sys27_Rnd(2)%>" style="width:580px;height:360px;visibility:hidden;display:none"></textarea>
<script type="text/javascript" charset="utf-8" src="../smod/file/edt_api.asp?edtID=EditID01<%=sys27_Rnd(2)%>&edtCont=InfCont<%=sys27_Rnd(2)%>"></script>

        </td>
      </tr>
      <%Else%>
      <tr>
        <td align="center" bgcolor="#F2F6FB">内 容        </td>
        <td colspan="2" align="left" bgcolor="#F2F6FB"><textarea name="InfCont<%=sys27_Rnd(2)%>" cols="60" rows="12" id="InfCont<%=sys27_Rnd(2)%>"></textarea></td>
      </tr>
      <%End If%>
            <tr>
              <td align="right" nowrap bgcolor="#F2F6FB">认证码</td>
              <td colspan="2" align="left" nowrap bgcolor="#F2F6FB"><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12">
              <img src="../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../')"/>
              <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>"></td>
            </tr>
      <tr>
        <td width="10%" bgcolor="#F2F6FB">&nbsp;</td>
        <td colspan="2" align="left" bgcolor="#F2F6FB"><input type="button" name="Button" value="提交" onClick="chkData()">
          <input type="reset" name="Submit2" value="重置">
          <input name="send" type="hidden" id="send" value="send">
          <input name="ID" type="hidden" id="ID" value="<%=ID%>">
          <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
          <input name="TypLay" type="hidden" id="TypLay" value="<%=TypLay%>" />
          <%=fMsg%></td>
      </tr>
    </form>
  </table>
  </fieldset>
  <div style="line-height:8px;">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td height="24" align="left" bgcolor="#FFFFFF">发贴说明: </td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
<script type="text/javascript">
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
 //if (document.fm01.InfCont.value.length==0)
 <%If bbsUser&""<>"" Then%>
   document.fm01.InfCont<%=sys27_Rnd(2)%>.value = apiGetValEditID01<%=sys27_Rnd(2)%>();
 <%End If%>
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length==0)
   {   
     alert(" 内容 不能为空！"); 
	 document.fm01.InfCont<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length>=24000) 
   {   
     alert(" 内容 不能超过 24 K字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
<!--BBBB-->
</body>
</html>
