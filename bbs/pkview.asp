<!--#include file="binc/_config.asp"-->
<%
ID = RequestS("ID",3,48)
Act = RequestS("Act",3,48)

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT TOP 1 * FROM [BPKTitle] WHERE KeyID='"&ID&"' ",conn,1,3 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
'InfCont = Replace(InfCont,"<br>"," &nbsp; ")
InfOrg = rs("InfOrg")&""
InfOUrl = rs("InfOUrl")
InfMemb = rs("InfMemb")
InfStart = rs("InfStart")
InfEnd = rs("InfEnd")
InfView1 = Show_Text(rs("InfView1"))
InfView2 = Show_Text(rs("InfView2"))
InfVote1 = rs("InfVote1")
InfVote2 = rs("InfVote2")
SetRead = rs("SetRead")
SetRead = SetRead + 1
SetHot = rs("SetHot")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
  StrDetail = ""
If InfOrg<>"" Then
  StrDetail = "<br>详情:<A href='"&InfOUrl&"' target=_blank><font color=blue>"&InfOrg&"</font></A>"
End If
Call rs_DoSql(conn,"UPDATE BPKTitle SET SetRead=SetRead+1 WHERE KeyID='"&KeyID&"'")
else
Response.Redirect "/pklist.asp"
end if 
rs.Close()
SET rs=Nothing

  StrDetail = ""
If InfOrg<>"" Then
  StrDetail = "<br>详情:<A href='"&InfOUrl&"' target=_blank><font color=blue>"&InfOrg&"</font></A>"
End If

Vote1 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Aegis'") 
Vote2 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Oppose'")
Vote3 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Neutral'")
View1 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Aegis'") 
View2 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Oppose'")
View3 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Neutral'")
Pnt1 = View1*5 + Vote1*2
Pnt2 = View2*5 + Vote2*2
Pnt3 = View3*5 + Vote3*2
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=bbsName%></title>
<link rel="stylesheet" type="text/css" href="pkimg/style.css">
</head>
<body>

<TABLE cellSpacing=0 cellPadding=0 width=800 align=center border=0>
  <TBODY>
    <TR>
      <TD width=1 bgColor=#a9a9a9></TD>
      <TD vAlign=top bgColor=#fcfcfc><TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD height=10></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=745 align=center border=0>
          <TBODY>
            <TR>
              <TD class=style1 width=656><A class=style12 
            href="/">论坛首页</A>&nbsp;&gt;&gt;&nbsp;<A class=style12 
            href="/pklist.asp">辩论首页</A> &gt;&gt; <%=InfSubj%></TD>
              <TD width=94 align="center"><DIV align=right><span class="style1"><a href="/pklist.asp"><IMG height=19 src="pkimg/s_100.gif" 
            width=64 border=0></a></span></DIV></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD height=10></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_49.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE height=88 cellSpacing=0 cellPadding=0 width=746 align=center 
      border=0>
          <TBODY>
            <TR>
              <TD width=1 bgColor=#bdbdbd></TD>
              <TD vAlign=top bgColor=#f5f5f5><TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD width=66><IMG height=23 src="pkimg/s_51.gif" 
                width=53></TD>
                      <TD class=style8 vAlign=bottom width=502><%=InfSubj%></TD>
                      <TD width=152><span class="style1"><a href="pkjoin.asp?ID=<%=KeyID%>&Act=PubView&TP="><img src="pkimg/s_162.gif" width="59" height="18" border="0"></a></span></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD bgColor=#dadada height=1></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE height=51 cellSpacing=0 cellPadding=0 width=720 align=center 
            border=0>
                  <TBODY>
                    <TR>
                      <TD class=style1 width=420><A 
                              class=style1 
                              href="/pkview.asp?ID=<%=KeyID%>" 
                              target=_blank><%=InfCont%></A> &nbsp; <%=StrDetail%></TD>
                      <TD width=1 bgColor=#dadada></TD>
                      <TD width=12>&nbsp;</TD>
                      <TD class=style1 width=287><TABLE width="100%">
                          <TBODY>
                            <TR class="style17">
                              <TD width="50%">·发起人：<%=InfMemb%></TD>
                              <TD>·开始时间：<%=InfStart%></TD>
                            </TR>
                            <TR class="style17">
                              <TD>·浏览量：<%=SetRead%></TD>
                              <TD>·结束时间：<%=InfEnd%></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width=1 bgColor=#bdbdbd></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_50.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD height=10></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD align="left" valign="top"><table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" height="54" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="1" bgcolor="#BDBDBD"></td>
                    <td valign="top" bgcolor="#FEFAE7"><table width="350" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="83" rowspan="2"><img src="pkimg/s_175.gif" width="80" height="93"></td>
                          <td height="76" colspan="3" class="style10"><%=InfView1%></td>
                        </tr>
                        <tr height="24">
                          <td width="24" height="24" align="center" valign="middle"><img src="pkimg/s_55.gif" width="17" height="13"></td>
                          <td align="left" valign="middle" class="style1">观点发表数：<%=View1%> 正方得票：<%=Vote1%></td>
                        </tr>
                        <tr>
                          <td height="24">&nbsp;</td>
                          <td width="24" class="style1">&nbsp;</td>
                          <td class="style1"><a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=PubView&TP=Aegis"><img 
						  src="pkimg/s_162.gif" width="59" height="18" border="0"></a>&nbsp;<a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=MemList&TP=Aegis"><img 
						  src="pkimg/s_161.gif" width="72" height="18" border="0"></a>&nbsp;<a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=PutVote&TP=Aegis"><img 
						  src="pkimg/s_118.gif" width="72" height="18" border="0"></a>&nbsp;</td>
                        </tr>
                      </table></td>
                    <td width="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="10"></td>
                  </tr>
                </table>
                <!-- Span /////////////////////////////////////// -->
                <%
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [BPKView] WHERE KeyMod='"&ID&"' AND KeyFlag='Aegis' ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
'InfCont = Replace(InfCont,"<br>"," &nbsp; ")
InfMemb = rs("InfMemb")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
%>
                <TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                  <TBODY>
                    <TR>
                      <TD width=18><TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                          <TBODY>
                            <TR>
                              <TD vAlign=top><TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD bgColor=#bdbdbd height=1></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD width=1 bgColor=#bdbdbd></TD>
                                      <TD vAlign=top bgColor=#FEFAE7><TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE 
                        style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                        cellSpacing=1 cellPadding=0 width=345 align=center 
                        border=0>
                                          <TBODY>
                                            <TR>
                                              <TD width=21><IMG height=16 src="pkimg/s_63.gif" width=12></TD>
                                              <TD class=style11 vAlign=bottom width=324><%=InfSubj%></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD><IMG height=1 src="pkimg/s_65.gif" width=345></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE 
                        style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                        cellSpacing=0 cellPadding=0 width=345 align=center 
                        border=0>
                                          <TBODY>
                                            <TR>
                                              <TD 
                              class=style1>&nbsp;&nbsp;&nbsp;<span class=style1><%=InfCont%></span></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD><IMG height=1 src="pkimg/s_65.gif" 
                              width=345></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD width=21><IMG height=12 
                              src="pkimg/s_62.gif" width=12></TD>
                                              <TD class=style1 vAlign=bottom width=130>作者：<A 
                              class=style1 
                              href="#"><%=InfMemb%></A></TD>
                                              <TD align="right" vAlign=bottom class=style1><span class="unnamed1">发表时间：<%=LogATime%></span></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE></TD>
                                      <TD width=1 bgColor=#bdbdbd></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD bgColor=#bdbdbd 
              height=1></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="10"></td>
                  </tr>
                </table>
                <%
rs.MoveNext
Loop
else
%>
                <table width="364" border="0" cellspacing="2" cellpadding="2">
                  <tr>
                    <td bgcolor="#FEFAE7">正方暂无辩论资料</td>
                  </tr>
                </table>
                <%
end if 
rs.Close()
SET rs=Nothing
%>
                <!-- Span /////////////////////////////////////// -->
              </TD>
              <TD align="right" valign="top"><table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" height="54" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="1" bgcolor="#BDBDBD"></td>
                    <td valign="top" bgcolor="#F0F7E7"><table width="350" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="83" rowspan="2"><img src="pkimg/s_176.gif" width="80" height="93"></td>
                          <td height="76" colspan="3" class="style10"><%=InfView2%></td>
                        </tr>
                        <tr>
                          <td width="24" height="24" align="center" valign="middle"><img src="pkimg/s_55.gif" width="17" height="13"></td>
                          <td align="left" valign="middle" class="style1">观点发表数：<%=View2%> 反方得票：<%=Vote2%></td>
                        </tr>
                        <tr>
                          <td height="24">&nbsp;</td>
                          <td width="24" class="style1">&nbsp;</td>
                          <td class="style1"><a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=PubView&TP=Oppose"><img 
						  src="pkimg/s_162.gif" width="59" height="18" border="0"></a>&nbsp;<a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=MemList&TP=Oppose"><img 
						  src="pkimg/s_161.gif" width="72" height="18" border="0"></a>&nbsp;<a 
						  href="pkjoin.asp?ID=<%=ID%>&Act=PutVote&TP=Oppose"><img 
						  src="pkimg/s_118.gif" width="72" height="18" border="0"></a>&nbsp;</td>
                        </tr>
                      </table></td>
                    <td width="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="1" bgcolor="#BDBDBD"></td>
                  </tr>
                </table>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="10"></td>
                  </tr>
                </table>
                <!-- Span /////////////////////////////////////// -->
                <%
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [BPKView] WHERE KeyMod='"&ID&"' AND KeyFlag='Oppose' ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
'InfCont = Replace(InfCont,"<br>"," &nbsp; ")
InfMemb = rs("InfMemb")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
%>
                <TABLE width=364 border=0 cellPadding=0 cellSpacing=0>
                  <TBODY>
                    <TR>
                      <TD width=18><TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                          <TBODY>
                            <TR>
                              <TD vAlign=top><TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD bgColor=#bdbdbd height=1></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD width=1 bgColor=#bdbdbd></TD>
                                      <TD vAlign=top bgColor=#F0F7E7><TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE 
                        style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                        cellSpacing=1 cellPadding=0 width=345 align=center 
                        border=0>
                                          <TBODY>
                                            <TR>
                                              <TD width=21><IMG height=16 src="pkimg/s_63.gif" width=12></TD>
                                              <TD class=style11 vAlign=bottom width=324><%=InfSubj%></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD><IMG height=1 src="pkimg/s_65.gif" width=345></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE 
                        style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                        cellSpacing=0 cellPadding=0 width=345 align=center 
                        border=0>
                                          <TBODY>
                                            <TR>
                                              <TD 
                              class=style1>&nbsp;&nbsp;&nbsp;<span class=style1><%=InfCont%></span></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD><IMG height=1 src="pkimg/s_65.gif" 
                              width=345></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD width=21><IMG height=12 
                              src="pkimg/s_62.gif" width=12></TD>
                                              <TD class=style1 vAlign=bottom width=130>作者：<A 
                              class=style1 
                              href="#"><%=InfMemb%></A></TD>
                                              <TD align="right" vAlign=bottom class=style1><span class="unnamed1">发表时间：<%=LogATime%></span></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE>
                                        <TABLE cellSpacing=0 cellPadding=0 width=345 
                        align=center border=0>
                                          <TBODY>
                                            <TR>
                                              <TD height=8></TD>
                                            </TR>
                                          </TBODY>
                                        </TABLE></TD>
                                      <TD width=1 bgColor=#bdbdbd></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=364 border=0>
                                  <TBODY>
                                    <TR>
                                      <TD bgColor=#bdbdbd 
              height=1></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="10"></td>
                  </tr>
                </table>
                <%
rs.MoveNext
Loop
else
%>
                <table width="364" border="0" cellspacing="2" cellpadding="2">
                  <tr>
                    <td bgcolor="#F0F7E7">反方暂无辩论资料</td>
                  </tr>
                </table>
                <%
end if 
rs.Close()
SET rs=Nothing
%>
                <!-- Span /////////////////////////////////////// --></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_49.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD width=1 bgColor=#bdbdbd></TD>
              <TD vAlign=top width=744 bgColor=#f5f5f5><TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD width=192><IMG height=31 src="pkimg/s_109.gif" 
                  width=125></TD>
                      <TD class=style8 vAlign=bottom width=391>&nbsp;</TD>
                      <TD width=112><DIV align=right><A 
                  href="pkjoin.asp?ID=<%=ID%>&Act=PubView&TP="><IMG 
                  height=21 src="pkimg/s_119.gif" width=112 
                  border=0></A></DIV></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD bgColor=#dadada height=1></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <!-- Span /////////////////////////////////////// -->
                <%
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [BPKView] WHERE KeyMod='"&ID&"' AND KeyFlag='Neutral' ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
'InfCont = Replace(InfCont,"<br>"," &nbsp; ")
InfMemb = rs("InfMemb")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
%>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD vAlign=top><TABLE style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                  cellSpacing=0 cellPadding=0 width=720 align=center border=0>
                          <TBODY>
                            <TR>
                              <TD width=21><IMG height=16 src="pkimg/s_63.gif" 
                        width=12></TD>
                              <TD class=style11 vAlign=bottom width=351><%=InfSubj%></TD>
                              <TD width=21><IMG height=12 src="pkimg/s_62.gif" 
                        width=12></TD>
                              <TD class=style1 vAlign=bottom width=130>作者：<%=InfMemb%></TD>
                              <TD class=style1 vAlign=bottom width=64>&nbsp;</TD>
                              <TD class=style1 vAlign=bottom width=67>&nbsp;</TD>
                              <TD class=style1 vAlign=bottom width=76>&nbsp;</TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
                  border=0>
                          <TBODY>
                            <TR>
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=720 border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=1 src="pkimg/s_117.gif" 
                      width=719></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
                  border=0>
                          <TBODY>
                            <TR>
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE style="TABLE-LAYOUT: fixed; WORD-WRAP: break-word" 
                  cellSpacing=0 cellPadding=0 width=720 border=0>
                          <TBODY>
                            <TR>
                              <TD><SPAN class=style1><%=InfCont%></SPAN></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD align="right" vAlign=bottom class=style17>发表时间：<%=LogATime%></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD bgColor=#dadada height=1></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <%
rs.MoveNext
Loop
else
%>
                <table width="364" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>暂无辩论资料</td>
                  </tr>
                </table>
                <%
end if 
rs.Close()
SET rs=Nothing
%>
                <!-- Span /////////////////////////////////////// -->
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD height=6></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width=1 bgColor=#bdbdbd></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_50.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD height=10></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_49.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE height=46 cellSpacing=0 cellPadding=0 width=746 align=center 
      border=0>
          <TBODY>
            <TR>
              <TD width=1 bgColor=#bdbdbd></TD>
              <TD vAlign=center bgColor=#f5f5f5><TABLE cellSpacing=0 cellPadding=0 width=720 align=center 
              border=0>
                  <TBODY>
                    <TR>
                      <TD width=66><IMG height=23 src="pkimg/s_66.gif" 
                width=54></TD>
                      <TD class=style8 width=271>
					  <%
					  If DateDiff("d",Date(),InfEnd)>0 Then
					    s = "辩论进行中..."
					  Else
					    s = "辩论结束: "
						If Pnt1=Pnt Then
						  s = s & " [无] 获胜方??? "
						ElseIf Pnt1>Pnt2 Then
						  s = s & " [正方] 胜利! <a href='../img/tool/flower.gif' target='_blank'><img src='../img/tool/flower.gif' width='16' height='16' border='0' align='absmiddle'></a>"
						ElseIf Pnt1<>Pnt2 Then
						  s = s & " [反方] 胜利! <a href='../img/tool/flower.gif' target='_blank'><img src='../img/tool/flower.gif' width='16' height='16' border='0' align='absmiddle'></a>"
						End If
					  End If
					  Response.Write s
					  %>
                       </TD>
                      <TD class=style8 width=271><table width="120" border="0" align="center" cellpadding="1" cellspacing="1">
                          <tr align="center" class="unnamed1">
                            <td>正方得分</td>
                            <td><%=Pnt1%></td>
                          </tr>
                          <tr align="center" class="unnamed1">
                            <td>反方得分</td>
                            <td><%=Pnt2%></td>
                          </tr>
                          <tr align="center" class="unnamed1">
                            <td>中立方得分</td>
                            <td><%=Pnt3%></td>
                          </tr>
                        </table></TD>
                      <TD width=112><span class="style1"><a href="pkjoin.asp?ID=<%=ID%>&Act=PubView&TP="><img src="pkimg/s_162.gif" width="59" height="18" border="0"></a></span></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width=1 bgColor=#bdbdbd></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=746 align=center border=0>
          <TBODY>
            <TR>
              <TD><IMG height=9 src="pkimg/s_50.gif" 
      width=746></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD height=18></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
      <TD width=1 bgColor=#a9a9a9></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="1" bgcolor="#BDBDBD"></td>
  </tr>
</table>

</BODY>
</HTML>
