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

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线调查投票系统-<%=Config_Name%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<script src="vote.js" type="text/javascript"></script>
</HEAD>
<BODY>
<!--#include file="_xtop.asp"-->
                  <!--////////////////////////////////////////Start主体-->
                  <table border="1" align="center" cellpadding="0" cellspacing="5">
                    <form action="?" method="post" name="fm01">
                      <tr>
                        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
                          <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> ( <a href="vprize.asp?ID=<%=ID%>" target="_blank">查看贺奖名单</a>) &nbsp; </LEGEND>
                          <%=InfRem1%>
                          </FIELDSET></td>
                      </tr>
                      <tr>
                        <td width="650" align="left" valign="bottom" style="padding:5px;"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1">
                            <%
sql = " SELECT * FROM [VoteItem] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY SetVote DESC,KeyID " 
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
i=i+1
If i<=10 Then 
 Tops = "<img src='vimg/"&i&".jpg' width=13 height=11>"
Else
 Tops = i
End If
iKeyID = rs("KeyID")
iInfSubj = rs("InfSubj")
iSetVote = rs("SetVote") :If i=1 Then MaxVote=iSetVote
iImage = rs("ImgName")&""
If iImage<>"" Then
 sImg = "../upfile/vote/"&iImage&""
Else
 sImg = "../img/logo/no_pic211.jpg"
End If
iRate = 100*iSetVote/MaxVote
iVot1 = Int(iRate)
iVot2 = 100-iVot1
%>
                            <tr>
                              <td width="5%" align="right" nowrap style="padding-right:3px;"><%=Tops%></td>
                              <td width="100" align="center" nowrap><a 
                                href="#" onClick="javascript:window.open('vitem.asp?ID=<%=iKeyID%>','Win3','scrollbars=yes,width=600,height=480')"><img 
                                src="<%=sImg%>" alt="Image" width="90" height="60" vspace="1" border="0" 
                                onload="javascript:setImgSize(this);"></a> </td>
                              <td width="80%" align="left" nowrap><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
                                  <tr>
                                    <td><a 
                                href="#" onClick="javascript:window.open('vitem.asp?ID=<%=iKeyID%>','Win3','scrollbars=yes,width=600,height=480')"><%=iInfSubj%></a></td>
                                  </tr>
                                  <tr>
                                    <td>现有：<%=iSetVote%>票 (<%=FormatNumber(iRate,2)%>%<font color="#CCCCCC">[以最多得票为基准]</font>)</td>
                                  </tr>
                                  <tr>
                                    <td><table width="450" border="0" align="center" cellpadding="0" cellspacing="0">
                                        <tr height="12">
                                          <td width="<%=iVot1%>%" style="background-image:url(../img/vote/vote52.gif); line-height:12px;"></td>
                                          <td width="<%=iVot2%>%" style="background-image:url(../img/vote/vote62.gif); line-height:12px;"></td>
                                        </tr>
                                      </table></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <%
rs.MoveNext()
Loop
rs.Close()
%>
                          </table></td>
                      </tr>
                    </form>
                  </table>
                  <!--////////////////////////////////////////End主体-->
                  <!--#include file="_xbot.asp"-->
</BODY>
</HTML>
