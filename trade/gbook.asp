<!--#include file="inc/_config.asp"-->
<!--#include file="inc/mem_info.asp"-->

<!--#include file="../upfile/sys/para/keywords.asp"-->
<%

'留言,评论,申请职位,订购产品
MD = RequestS("ModID",3,24) : If MD="" Then MD="TraG124"
If MD="TraG124" Then
  MDName = "在线留言"
  sqlK = " WHERE KeyMod='"&MD&"' AND KeyUser='"&US&"' AND SetShow='Y' "
  defAct1 = "给：" : defAct2 = "<a href='corp.asp?UsrID="&US&"' class='cRed'>"&MemSubj&"</a> 留言"
ElseIf MD="TraR124" Then
  MDName = "信息评论"
  ID = RequestS("ID",3,48)
  ObjSubj = Show_Text(Request("ObjSubj"))
  sqlK = " WHERE KeyMod='"&MD&"' AND KeyUser='"&US&"' AND SetShow='Y' AND InfReply='"&ID&"' "
  vGbo_SendOK1 = "感谢您评论!"
  vGbo_SendOK2 = "感谢您评论,我们回尽快处理!"
  vGbo_SendOK3 = "感谢您评论,评论含有敏感关键字,需要审核!"
  vGbo_SendNG = "评论失败!"
  defSubj = "评论:"&ObjSubj
  defAct1 = "评论" : defAct2 = "<a href='iview.asp?KeyID="&ID&"' class='cRed'>"&ObjSubj&"</a> "
  TP = "U11224"
ElseIf MD="TraA124" Then
  MDName = "申请职位"
  ID = RequestS("ID",3,48)
  ObjSubj = Show_Text(Request("ObjSubj"))
  sqlK = " WHERE KeyMod='"&MD&"' AND KeyUser='"&US&"' AND SetShow='Y' AND InfReply='"&ID&"' "
  vGbo_SendOK1 = "感谢您评论!"
  vGbo_SendOK2 = "感谢您评论,我们回尽快处理!"
  vGbo_SendOK3 = "感谢您评论,评论含有敏感关键字,需要审核!"
  vGbo_SendNG = "评论失败!"
  defSubj = "申请:"&ObjSubj&"职位"
  defAct1 = "申请：" : defAct2 = "<a href='iview.asp?KeyID="&ID&"' class='cRed'>"&ObjSubj&"</a> 职位"
  TP = "U11228"
ElseIf MD="TraO124" Then
  MDName = "订购产品"
  ID = RequestS("ID",3,48)
  ObjSubj = Show_Text(Request("ObjSubj"))
  sqlK = " WHERE 1=2 "
  vGbo_SendOK1 = "感谢您订购!"
  vGbo_SendOK2 = "感谢您订购,我们回尽快处理!"
  vGbo_SendOK3 = "感谢您订购,评论含有敏感关键字,需要审核!"
  vGbo_SendNG = "订购失败!"
  defSubj = "订购:"&ObjSubj
  defAct1 = "订购：" : defAct2 = "<a href='iview.asp?KeyID="&ID&"' class='cRed'>"&ObjSubj&"</a> 产品"
  TP = "U11232"
Else
  Response.End()
End If


Page = RequestS("Page","N",1)
ChkCode = Request("ChkCode")

send = Request("send")

Dim sys27_Rnd(10)
If send = "ins" Then
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

If send="ins" Then 
  Call Chk_Url()   'KeyID = Get_AutoID(24)
  KeyID = rs_AutID(conn,"TradeGbook","KeyID","dtbug","1","")
InfSubj = RequestS("InfSubj"&sys27_Rnd(1),C,255)
InfCont = RequestS("InfCont"&sys27_Rnd(2),C,500) 
fSubj = Chr_Fil2(InfSubj)
fCont = Chr_Fil2(InfCont)

  SetShow = "Y"
  fMsgShow = vGbo_SendOK1
If SwhGbkChk="N" Then 
  SetShow = "N"
  fMsgShow = vGbo_SendOK2
ElseIf fSubj<>InfSubj or fCont<>InfCont then
  SetShow = "N"
  fMsgShow = vGbo_SendOK3
Else

End If

  InfSubj = fSubj
  InfCont = fCont
'fMsgShow = fMsgShow&ParFilALLKeys

sql = " INSERT INTO [TradeGbook] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyUser" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfReply"
sql = sql& ", LnkName" 
sql = sql& ", LnkTel" 
sql = sql& ", LnkQQ" 
sql = sql& ", LnkEmail" 
sql = sql& ", SetRead" 
sql = sql& ", SetShow" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ", LogEditIP" 
sql = sql& ", LogEUser" 
sql = sql& ", LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & MD &"'" 
sql = sql& ", '" & US &"'" 
sql = sql& ", '" & RequestS("InfType",C,255) &"'"
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfCont &"'" 
sql = sql& ", '" & ID &"'" 
sql = sql& ", '" & RequestS("LnkName"&sys27_Rnd(4),C,24) &"'" 
sql = sql& ", '" & RequestS("LnkTel"&sys27_Rnd(5),C,48) &"'"
sql = sql& ", '" & RequestS("LnkQQ"&sys27_Rnd(5),C,48) &"'" 
sql = sql& ", '" & RequestS("LnkEmail"&sys27_Rnd(5),C,255) &"'" 
sql = sql& ", 0" 
sql = sql& ", '" & SetShow &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("MemID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '-'" 
sql = sql& ", '-'" 
sql = sql& ", '1900-12-31'" 
sql = sql& ")"

If  Session("ChkCode")<>uCase(ChkCode) Then
  fMsgShow = vMsg_ChkCErr
ElseIf Len(InfSubj)>0 And Len(InfCont)>0 Then
  Call rs_DoSql(conn,sql) 
  upPath = upRoot&Replace(KeyID,"-","/") 
  Call add_sfFile()
Else
  fMsgShow = vGbo_SendNG
End If

  Response.Write js_Alert(fMsgShow,"Redir","?UsrID="&US&"&ModID="&MD&"&ID="&ID&"&ObjSubj="&ObjSubj&"") 

End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-<%=vPMsg_WName%></title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../ext/api/play/jsPlayer.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="_head.asp"-->
<!--#include file="_menu.asp"-->
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
          <td height="500" colspan="4" align="left" valign="top" style="padding:12px 0px;"><!--Item Start-->
          
            <%
	sql = " SELECT * FROM [TradeGbook] "&sqlK
	sql =sql& " ORDER BY KeyID DESC " ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 5 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If
%>
            <table width="700" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
              <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize

KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&rs("InfType")&"'")
xxxCont =  Show_Text(rs("InfCont"))
xxxReply =  Show_Text(rs("InfReply"))
If xxxReply="" Then xxxReply="<span class=cRed>(暂无回复)</span>"
LnkName = Show_Text(rs("LnkName"))
LnkTel = rs("LnkTel")
LnkQQ = rs("LnkQQ")
LnkEmail = Show_Text(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LnkEmail = Replace(LnkEmail,"@","#")
If Len(LnkQQ)>=5 And isNumeric(LnkQQ) Then
  LnkQQ = "<span class='da01_qq'>&nbsp;</span><a href=tencent://message/?uin="&LnkQQ&"&Site="&MemSubj&"&Menu=yes>"&LnkQQ&"</a>"
Else
  If LnkQQ="" Then LnkQQ="00000"
  LnkQQ = "<span class=fntCCC>"&LnkQQ&"</span>"
End If
	  %>
              <tr>
                <td width="15%" align="center" valign="top" bgcolor="#FFFFFF">主题</td>
                <td width="40%" align="left" valign="top" bgcolor="#FFFFFF"><font color="#0000FF"><%=InfSubj%></font></td>
                <td width="15%" align="center" valign="top" bgcolor="#FFFFFF">时间</td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><%=LogATime%></td>
              </tr>
              <tr>
                <td align="center" valign="top" bgcolor="#FFFFFF">姓名</td>
                <td align="left" valign="top" bgcolor="#FFFFFF"><%=LnkName%></td>
                <td align="center" valign="top" bgcolor="#FFFFFF">电话</td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkTel%></td>
              </tr>
              <tr>
                <td align="center" valign="top" bgcolor="#FFFFFF">类别</td>
                <td align="left" valign="top" bgcolor="#FFFFFF"><%=TypName%></td>
                <td align="center" valign="top" bgcolor="#FFFFFF">ＱＱ</td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkQQ%></td>
              </tr>
              <tr>
                <td align="center" valign="top" bgcolor="#FFFFFF">IP</td>
                <td align="left" valign="top" bgcolor="#FFFFFF"><%=LogAddIP%></td>
                <td align="center" valign="top" bgcolor="#FFFFFF">邮件</td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><%=LnkEmail%></td>
              </tr>
              <tr>
                <td colspan="4" align="left" valign="top" bgcolor="#FFFFFF" style="padding:5px;"><fieldset style="paddingt: 5px; border:#F0F0F0 1px solid">
                    <legend class="fnt666"><%=vGbo_Cont%></legend>
                    <table width="100%" border="0" cellpadding="2" cellspacing="1">
                      <tr>
                        <td class="SysCont" style="padding:3px;"><%=xxxCont%></td>
                      </tr>
                    </table>
                    <%If MD="TraR124" Then%>
                    <div style="float:right; padding:5px 8px;"> 原文：<a href="iview.asp?KeyID=<%=ID%>" class="cRed"><%=ObjSubj%></a> </div>
                    <%End If%>
                  </fieldset></td>
              </tr>
              <%If MD="TraG124" Then%>
              <tr>
                <td colspan="4" align="left" valign="top" bgcolor="#FFFFFF" style="padding:5px;"><fieldset style="paddingt: 5px; border:#F0F0F0 1px solid">
                    <legend class="fnt666"><%=vGbo_Reply%></legend>
                    <table width="100%" border="0" cellpadding="2" cellspacing="1">
                      <tr>
                        <td class="SysCont" style="padding:8px;"><%=xxxReply%></td>
                      </tr>
                    </table>
                  </fieldset></td>
              </tr>
              <%End If%>
              <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
  
%>
              <tr>
                <td colspan="4" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?UsrID="&US&"&ModID="&MD&"&ID="&ID&"&ObjSubj="&ObjSubj&"&TP="&TP&"&KW="&KW&"",1)%></td>
              </tr>
              <%  
  Else
    If MD="TraG124" OR MD="TraR124" Then
  %>
              <tr>
                <td colspan="4" align="center" bgcolor="#FFFFFF"><%=vPMsg_InfNull%></td>
              </tr>
              <%
    End If
  End If
	  
  rs.Close()
	  
If Session("MemID")&""<>"" Then
  rs.Open "SELECT * FROM [TradeCorp] WHERE LogAUser='"&Session("MemID")&"'",conn,1,1 
  If NOT rs.eof then 
LnkName = rs("LnkName")
LnkMobile = rs("LnkMobile")
LnkTel = rs("LnkTel")
LnkFax = rs("LnkFax")
LnkEmail = rs("LnkEmail")
LnkQQ = rs("LnkQQ")
LnkUrl = rs("LnkUrl")
LnkAddr = rs("LnkAddr")
  End If
  rs.Close()
Else
LnkName = ""
LnkMobile = ""
LnkTel = ""
LnkFax = ""
LnkEmail = ""
LnkQQ = ""
LnkUrl = ""
LnkAddr = ""
End If
	  
	  %>
            </table>
            <a name="Write" id="Write"></a>
            <div class="line12">&nbsp;</div>
            <table width="700"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
              <form name="fm01" id="fm01" action="?" method="post">
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><%=defAct1%></td>
                  <td bgcolor="#FFFFFF"><%=defAct2%></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><%=vGbo_Subj%></td>
                  <td bgcolor="#FFFFFF"><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="36" maxlength="60" value="<%=defSubj%>">
                    &nbsp;<%=vInf_Type%>
                    <select name="InfType" id="InfType" style="width:210px;">
                      <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='BisU124'",TP)%>
                    </select></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><%=vGbo_Cont%><br />
                    (250字内)<br />
                    当前<span id="CntDiv">0</span>字</td>
                  <td bgcolor="#FFFFFF"><textarea name="InfCont<%=sys27_Rnd(2)%>" rows="4" style="width:450px" onblur="CntCont()"></textarea></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><%=vGbo_Name%></td>
                  <td bgcolor="#FFFFFF"><input name="LnkName<%=sys27_Rnd(4)%>" type="text" id="LnkName<%=sys27_Rnd(4)%>" value="<%=Session("MemID")%>" size="36" maxlength="24">
                    &nbsp;<%=vGbo_Email%>
                    <input name="LnkEmail<%=sys27_Rnd(5)%>" type="text" id="LnkEmail<%=sys27_Rnd(5)%>" value="<%=LnkEmail%>" size="36" maxlength="120" /></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF">电话</td>
                  <td bgcolor="#FFFFFF"><input name="LnkTel<%=sys27_Rnd(5)%>" type="text" id="LnkTel<%=sys27_Rnd(5)%>" value="<%=LnkTel%>" size="36" maxlength="15">
                    &nbsp;ＱＱ
                    <input name="LnkQQ<%=sys27_Rnd(5)%>" type="text" id="LnkQQ<%=sys27_Rnd(5)%>" value="<%=LnkQQ%>" size="36" maxlength="15" /></td>
                </tr>
                <tr>
                  <td width="12%" align="center" nowrap bgcolor="#FFFFFF"><%=vPMsg_ChkCode%></td>
                  <td nowrap bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="280" nowrap="nowrap"><input name="ChkCode" type="text" id="ChkCode" size="5" maxlength="12" style="width:50px" />
                          <img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/>
                          <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>"></td>
                        <td align="right" nowrap="nowrap"><input type="button" name="Button" value="<%=vGbo_Send%>" onclick="chkData()" />
                          &nbsp;&nbsp;
                          <input type="reset" name="Submit2" value="<%=vGbo_Reset%>" /></td>
                        <td width="150"><input name="send" type="hidden" id="send" value="ins" />
                          <input name="UsrID" type="hidden" id="UsrID" value="<%=US%>" />
                          <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
                          <input name="ID" type="hidden" id="ID" value="<%=ID%>" />
                          <input name="ObjSubj" type="hidden" id="ObjSubj" value="<%=ObjSubj%>" /></td>
                      </tr>
                    </table></td>
                </tr>
              </form>
            </table>
            <script type="text/javascript">

var dfm = document.getElementById('fm01');

function CntCont()
{
	var n = dfm.InfCont<%=sys27_Rnd(2)%>.value.length;
	getElmID('CntDiv').innerHTML = n;
	if(n>250) { alert('内容超过250个字了'); }
}

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (dfm.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" <%=vGbo_jsSubj%>"); 
     dfm.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
 if (dfm.InfCont<%=sys27_Rnd(2)%>.value.length==0) 
   {   
     alert(" <%=vGbo_jsCont%>"); 
	 dfm.InfCont<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }
 if (dfm.InfCont<%=sys27_Rnd(2)%>.value.length>=250) 
   {   
     alert(" <%=vGbo_jsCMax%> 250!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ dfm.submit(); }
}</script>
            <!--Item End-->
          
          </td>
        </tr>
      </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240"><!--Side Start-->
      <!-- #include file="_side.asp" -->
      <!--Side End--></td>
  </tr>
</table>
<!--#include file="_foot.asp"-->
</body>
</html>
