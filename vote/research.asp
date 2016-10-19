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
InfTim1 = rs("InfTime1")
InfTim2 = rs("InfTime2")
InfCard = rs("InfCard") 
InfVNum = rs("InfVNum") 
InfNum1 = rs("InfNum1") 
InfNum2 = rs("InfNum2") 
ImgName = rs("ImgName")
  else
KeyID = ""  
  end if 
  rs.Close()
If KeyID = "" Then Response.Redirect("rlist.asp")



fVote = False 
If DateDiff("d",InfTim2,Date())>0 Then
   fVNot = "调查已结束!<br><a href='rview.asp?ID="&ID&"'>查看结果</a>"
   'Response.Redirect "vres.asp"
ElseIf DateDiff("d",Date(),InfTim1)>0 Then
   fVNot = "调查还未开始!"
Else
   fVNot = "调查进行中..." 
 If InfVNum="D1T1" Then
   sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND LogAddIP='"&IP&"' AND LogATime>"&Cfg_FTime&Date()&Cfg_FTime&" "
  If rs_Exist(conn,sql)="EOF" Then 
   fVote = True
  Else
   fVNot = "每个IP每天只可提交1次,此IP可能已经有投票!" 
  End If
 ElseIf InfVNum="DNT1" Then
   sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND LogAddIP='"&IP&"' "
  If rs_Exist(conn,sql)="EOF" Then 
   fVote = True
  Else
   fVNot = "每IP共可提交1次,此IP可能已经有投票!"  
  End If
 Else
   fVote = True
 End If
End If
dis="Disabled" :If fVote Then dis="xDisabled"
'Response.Write fVote&fVNot&dis
NullMsg = "[必须填写]"
If InfCard="N" Then
NullMsg = "<font color=red>[可不填写]</font>"
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线调查投票系统-<%=Config_Name%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</HEAD>
<BODY>

<%

' 条件：投票时间，投票数量，IP,VNum，Request.Form("kItems").Count
If send="send" AND fVote Then

  Call Chk_Url()
  ChkCode = uCase(Request("ChkCode"))
  If Session("ChkCode")&""="" OR Session("ChkCode")<>ChkCode Then
   fMsg = "错误：认证码错误 或 超时错误！\n请刷新认证码后在再提交!"&Session("ChkCode")&":"&ChkCode
   Response.Write js_Alert(fMsg,"Alert","-1")
   Response.End()
  End If
 
 vItems = RequestS("vItems",3,8000) '//:vItems = Replace(vItems," ","")
 vName = RequestS("InfName",3,48)
 vCard = RequestS("InfCard",3,48)
 If InfCard="Y" Then
  sql="SELECT KeyID FROM VoteLogs WHERE KeyMod='"&ID&"' AND InfCard='"&vCard&"' "
  If rs_Exist(conn,sql)="EOF" Then 
   Call InsItem(vItems,vName,vCard)
   Response.Write js_Alert("提交成功!","Redir","?ID="&ID)
  Else
   fVNot = "每个身份证号只可提交1次" 
   Response.Write js_Alert("提交失败：\n"&fVNot&"!","Redir","?ID="&ID)
  End If
 Else
   Call InsItem(vItems,vName,vCard)
   Response.Write js_Alert("提交成功!","Redir","?ID="&ID)
 End If
End If

%>

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
i=0
Do While NOT rs.EOF
i=i+1
iKeyID = rs("KeyID")
iSubj = rs("InfSubj")
iRem = rs("InfRem")&""
aRem = Split(iRem,"|")
iTop = rs("SetTop")
iVote = rs("SetVote")
iImage = rs("ImgName")&""
jsItem = jsItem&iKeyID&"|"
jsITyp = jsITyp&iImage&"|"
If inStr("Select;SelSin",iImage)<=0 Then
 If iImage="FBlank" Then
  iImgExt = " (120个字以内 <font color=red>必填项！！！</font>)"
 Else
  iImgExt = " (120个字以内 <font color=gray>可空白</font>)"
 End If
End If
%>
                          <div style="padding-top:12px;">
                            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td align="left"> <b><span id="Msg<%=iKeyID%>"><%=i%>. <%=iSubj%></span></b> <%=iImgExt%> </td>
                              </tr>
                              <tr>
                                <td><%
								If iImage="Select" Or iImage="SelSin" Then
								  For j=0 To uBound(aRem)
								    jItem = Trim(Show_Text(aRem(j)))
								    jChar = Chr(65+j)
								%>
                                    &nbsp;&nbsp;
                                    <%If iImage="SelSin" Then%>
                                    <input name="iBox<%=iKeyID%>" type="radio"    id="iBox<%=iKeyID%>" value="1" style="vertical-align:middle;" onclick="javascript:return CheckItem(this)" />
                                    <%Else%>
                                    <input name="iBox<%=iKeyID%>" type="checkbox" id="iBox<%=iKeyID%>" value="1" style="vertical-align:middle;" onclick="javascript:return CheckItem(this)" >
                                    <%End If%>
									<%=jChar&". "&jItem%><br>
                                    <%
								  Next
								Else 
								%>
                                  &nbsp;&nbsp;&nbsp;
                                  <textarea name="iTxt<%=iKeyID%>" cols="56" rows="3" id="iTxt<%=iKeyID%>"></textarea>
                                  <%
								End If
								  %>
                                </td>
                              </tr>
                            </table>
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
                              <%If dis="xDisabled" Then%>
                              <TR align=middle>
                                <TD 
                              height=30 colSpan=2 align="left">被调查人资料：（便于与获奖者联系，主办方对调查人资料保密） <%=NullMsg%></TD>
                              </TR>
                              <TR>
                                <TD width="200" nowrap>姓　名：
                                  <INPUT name=InfName id=InfName size="15" maxLength=12>
                                  <input name="send" type="hidden" id="send" value="send"></TD>
                                <TD nowrap>&nbsp;          电　话：
                                  <input name=InfTel id=InfTel size="15" maxlength=15>
                                  <input name="ID" type="hidden" id="ID" value="<%=ID%>"></TD>
                              </TR>
                              <TR>
                                <TD colspan="2" nowrap>身份证：
                                  <INPUT name=InfCard id=InfCard size="48" maxLength=20>
                                <input name="vItems" type="hidden" id="vItems" />                                </TD>
                              </TR>
            <tr>
              <td colspan="2" align="left" nowrap>认证码：
                <input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12">
                <img src="../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../')"/>
                <input name="goPage" type="hidden" id="goPage" value="<%=goPage%>"></td>
              </tr>
                              <TR>
                                <TD colspan="2" align="left" nowrap><span style="float:right; padding-right:5%;">
                                  <input name="" type="button" value="查看结果" onclick="javascript:location.href='rview.asp?ID=<%=ID%>';"/></span>
                                  &nbsp;&nbsp;&nbsp;&nbsp;
                                  <input type="button" name="button" id="button" value="提交" <%=dis%> onClick="chkData()">
                                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <input type="reset" name="button2" id="button2" value="重置"></TD>
                              </TR>
                              <%Else%>
                              <TR align=middle>
                                <TD height=30 colSpan=2 align="center" bgcolor="ffffDD"><span style="float:right; padding-right:5%;"><input name="" type="button" value="查看结果" onclick="javascript:location.href='rview.asp?ID=<%=ID%>';"/></span>注意：<%=fVNot%> </TD>
                              </TR>
                              <%End If%>
                        </TABLE></td>
                      </tr>
                    </form>
                  </table>
                  <!--////////////////////////////////////////End主体-->
                  <!--#include file="_xbot.asp"-->
<script type="text/javascript">var fVote='<%=fVote%>'; var InfCard='<%=InfCard%>'; var jsItem='<%=jsItem%>';var jsITyp='<%=jsITyp%>';</script>
<script src="research.js" type="text/javascript"></script>
</BODY>
</HTML>
