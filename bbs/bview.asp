<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%

ID = RequestS("ID","C",48)
send = Request("send")
Page = RequestS("Page","N",1)

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

'/////////////////////////// 增加资料
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
  fMsgShow = "感谢您回帖!"
  If SwhBBSChk="N" Then 'rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='SWRemark'")="N" Then
    SetShow = "N"
    fMsgShow = "感谢您回帖,我们回尽快审核!"
  ElseIf fSubj<>InfSubj or fCont<>InfCont then
	SetShow = "N"
    fMsgShow = "感谢您回帖,帖子含有关键字:\n["&fSubj&fCont&"]需要审核!"
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
sql = sql& ", '" & ID & "'"  'KeyRe
sql = sql& ", '"&ModID&"'" 
sql = sql& ", 'Y'" 'KeyFlag
sql = sql& ", '" & RequestS("InfType",C,255) &"'"
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & RequestS("LnkName"&sys27_Rnd(4),C,24) &"'" 
sql = sql& ", '" & RequestS("LnkFace",C,48) &"'" 
sql = sql& ", '" & RequestS("LnkEmail",C,255) &"'" 
sql = sql& ", 0" 
sql = sql& ", 'Y'" ' SetShow
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

 If Len(InfSubj)>0 And Len(InfCont&xxxCont)>0 And Session("ChkCode")=ChkCode And Len(ChkCode)>1 Then
  Call rs_DoSql(conn,sql) 
  Call add_sfFile()
  Session("ImgList")=""
  Response.Write js_Alert(fMsgShow,"Redir","bview.asp?ID="&ID) 
  'Response.Redirect "bview.asp?ID="&ID
 Else
  fMsg = "帖子回复失败,请重新提交！\n可能认证码错误或内容为空！" ': Response.Write sql  
  Response.Write js_Alert(fMsg,"Redir","bview.asp?ID="&ID&"") 
 End If

elseif send="Del" then
  'Response.Write bbsPath&"bview.asp"
  If Chk_URL3(bbsPath&"bview.asp")="eUrl" Then
    Response.End()
  End If
 dID = RequestS("dID","C",48)
 dName = bbsUser
 if dID=ID then
   If bbsChkUser(dID,bbsUser,"Self") Then
     Call bbsDelID(dID,"") 
     aTP = Split(Request("TP"),";")
     Response.Write js_Alert("删除主帖成功！","Redir","blist.asp?TypID="&aTP(1)&"&TypLay="&aTP(0)&";"&aTP(1)&";")
   End If 
 else
   If bbsChkUser(dID,bbsUser,"Self") Then
     Call bbsDelID(dID,"Rep") 
     Response.Write js_Alert("删除回复帖成功！","Redir","bview.asp?ID="&ID)
   End If 
   Response.End()
 end if

else
  
  Session("KeyID") = rs_AutID(conn,"BBSInfo","KeyID","dtbbs","1","")

end if

'/////////////////////////// 取出资料
	sql = " SELECT BBSInfo.* FROM [BBSInfo] WHERE KeyID='"&ID&"' "
   rs.Open Sql,conn,1,1
if Not rs.EOF then
KeyRe = rs("KeyRe")
KeyFlag = rs("KeyFlag")
InfSubj = Show_Text(rs("InfSubj"))
InfSubjB = InfSubj
InfTypeB = rs("InfType")
xxxCont = Show_Html(rs("InfCont"))
LnkName = Show_Text(rs("LnkName"))
LnkFace = rs("LnkFace")&""
LnkEmail = Show_Text(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
SetSAdm = rs("SetSAdm")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
LnkFace = Replace(LnkFace,"src='img/","src='../img/")
Call rs_Dosql(conn,"UPDATE BBSInfo SET SetRead=SetRead+1 WHERE KeyID='"&ID&"' ")
Else
Response.End() '//////////////
End If
rs.Close()

If LogAUser&""<>"" Then
'LnkName = rs_Val("","SELECT MemName FROM Member_ABCDE WHERE MemID='"&LogAUser&"'")
LnkName = GetMemInfo(LogAUser)
mReply = rs_Count(conn," BBSInfo WHERE KeyMod='BBSP124' AND LnkName='"&LogAUser&"' AND LEN(KeyRE)>12 ")
mPub = rs_Count(conn," BBSInfo WHERE KeyMod='BBSP124' AND LnkName='"&LogAUser&"' ")-mReply
'Response.Write fMsg&"<br>"&InfSubj&"<br>"&InfCont&"<br>"&chkCod2&"<br>"&Session("ChkCode")
Else
mGrade = "<font color=gray>(匿名)</font>"
mReply = "<font color=gray>(0)</font>"
mPub = "<font color=gray>(0)</font>"
End If

'TypID = RequestS("TypID","C",48)
TypLay = InfTypeB 'RequestS("TypLay","C",255) 
aParent = Split(TypLay,";")
TypParent = aParent(0)
TypID = aParent(1) 
TName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&TypID&"'")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=InfSubjB%>-<%=bbsName%></title>
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
      <td align="left" nowrap="nowrap" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; <a href="blist.asp?TypID=<%=TypID%>&amp;TypLay=<%=TypLay%>"><%=TName%></a> &gt;&gt; 帖子查看</td>
      <td align="right" nowrap="nowrap" bgcolor="#FFFFFF">&nbsp;<a href="blist.asp?TypID=<%=TypID%>&amp;TypLay=<%=TypLay%>">返回列表</a>&nbsp;</td>
    </tr>
  </table>
  <!--AAAA-->
  <TABLE width="960" border="0" align="center" cellPadding="0" cellSpacing="1" bgcolor="#9C0E0C" class="sysTabC1">
    <TR bgColor="#fbfdff">
      <TD vAlign="top" width="160" bgColor="#F0F0F0"><TABLE cellSpacing="0" cellPadding="8" width="100%" border="0">
          <TR vAlign="top">
            <TD align="center"><TABLE border="0" align="center" cellPadding="2" cellSpacing="1">
                <TR>
                  <TD align="center"><img src="<%=LnkFace%>" align="absmiddle"></TD>
                </TR>
                <TR>
                  <TD align="left"><%=LnkName%></TD>
                </TR>
              </TABLE>
              <!--头衔: <%=mGrade%><BR>-->
              贴数: <%=mPub%>篇<BR>
              回复: <%=mReply%>篇<BR>
            </TD>
          </TR>
        </TABLE></TD>
      <TD vAlign="top" width="*"><TABLE cellSpacing="0" cellPadding="6" width="100%" border="0">
          <TR>
            <TD align="left"><font color="#000033"><b><%=InfSubjB%></b></font>
              <HR noShade SIZE="1">
              <TABLE width="96%" border="0" align="center" cellPadding="0" cellSpacing="0">
                <TR vAlign="top">
                  <TD align="left"><%Call Show_sfData(ID,"fcont.htm")%></TD>
                </TR>
              </TABLE>
              <HR noShade SIZE="1">
              <TABLE cellSpacing="0" cellPadding="0" width="96%" border="0">
                <TR>
                  <TD align="left"><%if bbsUser=LogAUser And bbsUser&""<>"" then%>
                    <a href="bedit.asp?ID=<%=ID%>&eID=<%=ID%>" xtarget="_blank"><img src="bimg/edit.gif" width="47" height="18" align="absmiddle" border="0"></a> <a href="bview.asp?ID=<%=ID%>&dID=<%=ID%>&TP=<%=InfTypeB%>&send=Del"><img src="bimg/del.gif" width="45" height="18" align="absmiddle" border="0"></a>
                    <%else%>
                    <img src="bimg/edit.gif" width="47" height="18" align="absmiddle" style='filter:Alpha(Opacity=50);'> <img src="bimg/del.gif" width="45" height="18" align="absmiddle" style='filter:Alpha(Opacity=50);'>
                  <%end if%>                  </TD>
                  <TD>发贴时间:<%=LogATime%></TD>
                  <TD width="30%" align="right"><img src="bimg/ip.gif" width="13" height="15" align="absmiddle">IP:<%=LogAddIP%></TD>
                </TR>
              </TABLE></TD>
          </TR>
        </TABLE></TD>
    </TR>
  </TABLE>
  <div style="line-height:8px;">&nbsp;</div>
  <%
    sql = " SELECT BBSInfo.* FROM [BBSInfo] "
	sql =sql& " WHERE KeyRE='"&ID&"' "
   rs.Open Sql,conn,1,1
   rs.PageSize = 6 
   If Int(Page)>rs.PageCount Or Int(Page)<1 Then
      Page = 1
   End If
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
KeyID = rs("KeyID") 
KeyFlag = rs("KeyFlag")
InfSubj = Show_Text(rs("InfSubj"))
InfType = rs("InfType")
xxxCont = Show_Html(rs("InfCont"))
LnkName = Show_Text(rs("LnkName"))
LnkFace = rs("LnkFace")&""
LnkEmail = Show_Text(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
SetSAdm = rs("SetSAdm")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
LnkFace = Replace(LnkFace,"src='img/","src='../img/")

If LogAUser&""<>"" Then
'mGrade = rs_Val("","SELECT MemName FROM Member_ABCDE WHERE MemID='"&LogAUser&"'")
LnkName = GetMemInfo(LogAUser)
mReply = rs_Count(conn," BBSInfo WHERE KeyMod='BBSP124' AND LnkName='"&LogAUser&"' AND LEN(KeyRE)>12 ")
mPub = rs_Count(conn," BBSInfo WHERE KeyMod='BBSP124' AND LnkName='"&LogAUser&"' ")-mReply
'Response.Write fMsg&"<br>"&InfSubj&"<br>"&InfCont&"<br>"&chkCod2&"<br>"&Session("ChkCode")
Else
mGrade = "<font color=gray>(匿名)</font>"
mReply = "<font color=gray>(0)</font>"
mPub = "<font color=gray>(0)</font>"
End If

%>
  <TABLE width="960" border="0" align="center" cellPadding="0" cellSpacing="1" bgcolor="#9C0E0C" class="sysTabC1">
    <TR bgColor="#fbfdff">
      <TD vAlign="top" width="160" bgColor="#F0F0F0"><TABLE cellSpacing="0" cellPadding="8" width="100%" border="0">
          <TR vAlign="top">
            <TD align="center"><TABLE border="0" align="center" cellPadding="2" cellSpacing="1">
                <TR>
                  <TD align="center"><img src="<%=LnkFace%>" align="absmiddle"></TD>
                </TR>
                <TR>
                  <TD align="left"><%=LnkName%></TD>
                </TR>
              </TABLE>
              <!--头衔: <%=mGrade%><BR>-->
              贴数: <%=mPub%>篇<BR>
              回复: <%=mReply%>篇<BR>
            </TD>
          </TR>
        </TABLE></TD>
      <TD vAlign="top" width="*"><TABLE cellSpacing="0" cellPadding="6" width="100%" border="0">
          <TR>
            <TD align="left"><font color="#000033"><b><%=InfSubj%></b></font>
              <HR noShade SIZE="1">
              <TABLE width="96%" border="0" align="center" cellPadding="0" cellSpacing="0">
                <TR vAlign="top">
                  <TD align="left"><%Call Show_sfData(KeyID,"fcont.htm")%></TD>
                </TR>
              </TABLE>
              <HR noShade SIZE="1">
              <TABLE cellSpacing="0" cellPadding="0" width="96%" border="0">
                <TR>
                  <TD><%if bbsUser=LogAUser And bbsUser&""<>"" then%>
                    <a href="bedit.asp?ID=<%=ID%>&eID=<%=KeyID%>" xtarget="_blank"><img src="bimg/edit.gif" width="47" height="18" align="absmiddle" border="0"></a> <a href="bview.asp?ID=<%=ID%>&dID=<%=KeyID%>&TP=<%=InfTypeB%>&send=Del"><img src="bimg/del.gif" width="45" height="18" align="absmiddle" border="0"></a>
                    <%else%>
                    <img src="bimg/edit.gif" width="47" height="18" align="absmiddle" style='filter:Alpha(Opacity=50);'> <img src="bimg/del.gif" width="45" height="18" align="absmiddle" style='filter:Alpha(Opacity=50);'>
                    <%end if%>
                  </TD>
                  <TD>发贴时间:<%=LogATime%></TD>
                  <TD width="30%" align="right"><img src="bimg/ip.gif" width="13" height="15" align="absmiddle">IP:<%=LogAddIP%></TD>
                </TR>
            </TABLE></TD>
          </TR>
        </TABLE></TD>
    </TR>
  </TABLE>
  <div style="line-height:8px;">&nbsp;</div>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <%= RS_Page(rs,Page,"?send=pag&ID="&ID&"",1)%>
  <div style="line-height:8px;">&nbsp;</div>
  <%  
  Else
%>
  <div style="line-height:8px;">&nbsp;</div>
  <TABLE width="960" border="0" align="center" cellPadding="2" cellSpacing="1" class="sysTabC1">
    <TR>
      <TD align="center" bgcolor="#FFFFFF"> 无回复信息 </TD>
    </TR>
  </TABLE>
  <div style="line-height:5px;">&nbsp;</div>
  <%
  End If
  rs.Close()
%>
  <fieldset style="padding:3px; width:950px;">
  <legend> &nbsp; *贴子回复： </strong>RE:<%=InfSubjB%></strong> </legend>

  <%
  If KeyRe&""="" Then
  %>
  <table width="720"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="fm01" id="fm01" action="?" method="post">
      <tr>
        <td align="center" nowrap bgcolor="#F2F6FB">主题</td>
        <td align="left" bgcolor="#F2F6FB"><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="36" maxlength="60" value="RE:<%=InfSubjB%>"></td>
        <td rowspan="3" align="center" valign="top" bgcolor="#F2F6FB"><img src="../img/Head/1.gif" width="75" height="85" id="LnkFac2"/></td>
      </tr>
      <tr>
        <td align="center" nowrap bgcolor="#F2F6FB">类别</td>
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
        <td align="center" nowrap bgcolor="#F2F6FB">姓名</td>
        <td align="left" bgcolor="#F2F6FB"><input name="LnkName<%=sys27_Rnd(4)%>" type="text" id="LnkName<%=sys27_Rnd(4)%>" value="<%=Session("MemID")%>" size="36"></td>
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
        <td colspan="2" align="left" bgcolor="#F2F6FB"><input type="button" name="Button" value="提交" onClick="chkData()" <%If KeyRe&""<>"" Then Response.Write("disabled")%>>
          &nbsp;&nbsp;
          <input type="reset" name="Submit2" value="重置">
          <input name="send" type="hidden" id="send" value="send">
          <input name="ID" type="hidden" id="ID" value="<%=ID%>">
          <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
          <%=fMsg%></td>
      </tr>
    </form>
  </table>
  <%Else%>
      <span class="fntF00">不能回复：本帖子本身是回复贴；　请
  <a href="blist.asp?TypID=<%=TypID%>&amp;TypLay=<%=TypLay%>">返回列表</a>
  </span>
  <%End If%>
  </fieldset>
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
  <div style="line-height:8px;"></div>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
</body>
</html>
