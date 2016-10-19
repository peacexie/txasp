<!--#include file="_config.asp"-->
<%

ID = RequestS("ID","C",48)
IP = Get_CIP()
send = Request.Form("send")



  rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID") : js="VotCount=0;"
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfRem1 = Show_Html(rs("InfRem1"))
InfTim1 = rs("InfTime1") :js=js&"VotTim1='"&InfTim1&"';"
InfTim2 = rs("InfTime2") :js=js&"VotTim2='"&InfTim2&"';"
InfCard = rs("InfCard")  :js=js&"InfCard='"&InfCard&"';" 'Y/N
InfVNum = rs("InfVNum")  :js=js&"InfVNum='"&InfVNum&"';" 'D1T1,DNT1,DXTN
InfNum1 = rs("InfNum1")  :js=js&"InfNum1="&InfNum1&";"
InfNum2 = rs("InfNum2")  :js=js&"InfNum2="&InfNum2&";"
ImgName = rs("ImgName")
  else
KeyID = ""  
  end if 
  rs.Close()
If KeyID = "" Then Response.Redirect("vlist.asp")



fVote = False 
If DateDiff("d",InfTim2,Date())>0 Then
   fVNot = "投票已结束!<br><a href='vres.asp?ID="&ID&"'>查看结果</a>"
   'Response.Redirect "vres.asp"
ElseIf DateDiff("d",Date(),InfTim1)>0 Then
   fVNot = "投票还未开始!"
Else
   fVNot = "投票进行中..." 
 If InfVNum="D1T1" Then
   sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND LogAddIP='"&IP&"' AND LogATime>"&Cfg_FTime&Date()&Cfg_FTime&" "
  If rs_Exist(conn,sql)="EOF" Then 
   fVote = True
  Else
   fVNot = "每个IP每天只可投票1次,此IP可能已经有投票!" 
  End If
 ElseIf InfVNum="DNT1" Then
   sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND LogAddIP='"&IP&"' "
  If rs_Exist(conn,sql)="EOF" Then 
   fVote = True
  Else
   fVNot = "每IP共可投票1次,此IP可能已经有投票!"  
  End If
 Else
   fVote = True
 End If
End If
dis="Disabled" :If fVote Then dis="xDisabled"
'Response.Write fVote&fVNot&dis



' 条件：投票时间，投票数量，IP,VNum，Request.Form("kItems").Count
If send="send" AND fVote Then

  Call Chk_Url()
  ChkCode = uCase(Request("ChkCode"))
  If Session("ChkCode")&""="" OR Session("ChkCode")<>ChkCode Then
   fMsg = "错误：认证码错误 或 超时错误！\n请重新提交!"&Session("ChkCode")&":"&ChkCode
   'Response.Write js_Alert(fMsg,"Alert","-1")
   Response.Write js_Alert(fMsg,"Redir","?ID="&ID&"")
   Response.End()
  End If

 vItems = RequestS("kItems",3,1200) :vItems = Replace(vItems," ","")
 vName = RequestS("InfName",3,48)
 vCard = RequestS("InfCard",3,48)
 If InfCard="Y" Then
  sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND InfCard='"&vCard&"' "
  If rs_Exist(conn,sql)="EOF" Then 
   Call InsItem(vItems,vName,vCard)
   Response.Write js_Alert("投票成功!","Redir","?ID="&ID)
  Else
   fVNot = "每个身份证号只可投票1次" 
   Response.Write js_Alert("投票失败：\n"&fVNot&"!","Redir","?ID="&ID)
  End If
 Else
   Call InsItem(vItems,vName,vCard)
   Response.Write js_Alert("投票成功!","Redir","?ID="&ID)
 End If
End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线调查投票系统-<%=Config_Name%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
<!--
.vIOut {
	width:135px;
	height:auto;
	margin:0px;
	padding:0px;
	float:left;
}
.vItem {
	width:126px;
	height:auto;
	margin:4px;
	padding:0px;
	font-size:12px;
	border:1px #CCCCCC solid;
}
.vH100 { line-height:100%; }
-->
</style>
<script type="text/javascript"><%=js%></script>
<script src="vote.js" type="text/javascript"></script>
</HEAD>
<BODY>
<!--#include file="_xtop.asp"-->
                  <!--////////////////////////////////////////Start主体-->
                  <table border="1" align="center" cellpadding="0" cellspacing="5">
                    <form action="?" method="post" name="fm01">
                      <tr>
                        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
                          <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> (<%=InfTim1%> ~ <%=InfTim2%>) &nbsp; </LEGEND>
                          <%=InfRem1%>
                          </FIELDSET></td>
                      </tr>
                      <tr>
                        <td width="650" align="left" valign="bottom" style="padding:5px 2px 5px 5px;"><%
sql = " SELECT * FROM [VoteItem] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY SetTop,KeyID " 
rs.Open Sql,conn,1,1
Do While NOT rs.EOF
iKeyID = rs("KeyID")
iInfSubj = rs("InfSubj")
iSetTop = rs("SetTop")
iSetVote = rs("SetVote")
iImage = rs("ImgName")&""
If iImage<>"" Then
 sImg = "../upfile/vote/"&iImage&""
Else
 sImg = "../img/logo/no_pic211.jpg"
End If
%>
                          <div class="vIOut">
                          <div class="vItem">
                            <table border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="120" height="95" align="center" valign="bottom" style="padding-top:2px;"><a 
                                href="#" onClick="javascript:window.open('vitem.asp?ID=<%=iKeyID%>','Win3','scrollbars=yes,width=600,height=480')"><img 
                                src="<%=sImg%>" alt="Image" width="110" height="90" vspace="1" border="0" 
                                onload="javascript:setImgSize(this);"></a></td>
                              </tr>
                              <tr>
                                <td align="center"><a 
                                href="#" onClick="javascript:window.open('vitem.asp?ID=<%=iKeyID%>','Win3','scrollbars=yes,width=600,height=480')"><%=iInfSubj%></a></td>
                              </tr>
                              <tr>
                                <td class="vH100">&nbsp;现有：<%=iSetVote%>票</td>
                              </tr>
                              <tr>
                                <td class="vH100"><input name="kItems" type="checkbox" id="kItems" value="<%=iKeyID%>"
                    onclick="javascript:return VoteItem(this)" style="vertical-align:middle;">
                                投此一票</td>
                              </tr>
                            </table>
                          </div>
                          </div>
                          <%
rs.MoveNext()
Loop
rs.Close()
%>
                        </td>
                      </tr>
                      <tr>
                        <td><TABLE width="96%" align="center">
                            <TBODY>
                              <%If dis="xDisabled" Then%>
                              <TR align=middle>
                                <TD 
                              height=30 colSpan=2 align="left">投票人资料：（便于与获奖者联系，主办方对投票人资料保密） </TD>
                              </TR>
                              <TR>
                                <TD width="200" nowrap>姓　名：
                                  <INPUT name=InfName id=InfName size="15" maxLength=12>
                                  <input name="send" type="hidden" id="send" value="send"></TD>
                                <TD nowrap>&nbsp;          电话：
                                  <input name=InfTel id=InfTel size="15" maxlength=15>
                                  <input name="ID" type="hidden" id="ID" value="<%=ID%>"></TD>
                              </TR>
                              <TR>
                                <TD colspan="2" nowrap>身份证：
                                  <INPUT name=InfCard id=InfCard size="48" maxLength=20>
                                </TD>
                              </TR>
            <tr>
              <td colspan="2" align="left" nowrap>认证码：
                <input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12">
                <img src="../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../')"/>
                <input name="goPage" type="hidden" id="goPage" value="<%=goPage%>"></td>
              </tr>
                              <TR>
                                <TD colspan="2" align="center" nowrap><input type="button" name="button" id="button" value="按钮"
					<%=dis%> onClick="chkData()">
                                  &nbsp;
                                  <input type="reset" name="button2" id="button2" value="重置"></TD>
                              </TR>
                              <%Else%>
                              <TR align=middle>
                                <TD 
                              height=30 colSpan=2 align="center" bgcolor="ffffDD">注意：<%=fVNot%> </TD>
                              </TR>
                              <%End If%>
                            </TBODY>
                          </TABLE></td>
                      </tr>
                    </form>
                  </table>
                  <!--////////////////////////////////////////End主体-->
                  <!--#include file="_xbot.asp"-->
</BODY>
</HTML>
