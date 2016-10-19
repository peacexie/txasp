<!--#include file="binc/_config.asp"-->
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
      <TD vAlign=top><TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD height=10></TD>
            </TR>
          </TBODY>
        </TABLE>
        <TABLE cellSpacing=0 cellPadding=0 width=750 align=center border=0>
          <TBODY>
            <TR>
              <TD vAlign=top width=172>
                  <TABLE cellSpacing=0 cellPadding=0 border=0>
                    <TBODY>
                      <TR>
                        <TD><IMG height=34 src="pkimg/login_top.gif" 
              width=172></TD>
                      </TR>
                      <TR>
                        <TD align=middle background=pkimg/login_bg.gif>
						
						
<%
  if Session("chkCode")&"" = "" then
    Randomize
    chkCode = 12345678 + Int(87654321*Rnd()) 
    Session("chkCode") = Right(chkCode,4)&Left(chkCode,4)
  end if
	if Session("MBID")&"" = "" then
	%>
      <TABLE border=0 align=center cellPadding=1 cellSpacing=1>
        <FORM name=ff action="../sys/member/login.asp" method=post>
          <TBODY>
            <TR align="center" class="style1">
              <TD colspan="2" nowrap><a href="/sys/member/service/mu_<%=AppRand12%>.asp" style="z-index:9986; "><font color="#0000FF">注册用户</font></a>|<a href="/sys/member/service/get_pw.asp" style="z-index:9987; "><font color="#FF0000">忘记密码</font></a> </TD>
              </TR>
            <TR class="style1">
              <TD align="right" nowrap>帐号</TD>
              <TD nowrap><INPUT name=MBID class=bu1 id="MBID" value="" 
                              size=15  maxLength=24 style="z-index:9981; ">
              <input name="chkCode" type="hidden" id="chkCode" value="(iTop)"></TD>
              </TR>
            <TR class="style1">
              <TD align="right" nowrap>密码</TD>
              <TD nowrap><INPUT name=MBPW 
                               type=password class=bu1 id="MBPW" size=15 
                              maxLength=24 style="z-index:9982; ">
              <input name="DirPag" type="hidden" id="DirPag" value="/pklist.asp"></TD>
              </TR>
            <TR class="style1">
              <TD align="right" nowrap><input name="send" type="hidden" id="send" value="send"></TD>
              <TD nowrap><input type="submit" name="Submit" value="登陆" style="z-index:9983; ">
&nbsp;     
                <input id=Submit1 type=Reset value=重设 name=Submit1 >     
              </TD>     
              </TR>     
          </TBODY>     
        </FORM>     
      </TABLE>     
      <%else%>   
      <TABLE border=0      
                        align=center cellPadding=0 cellSpacing=1>     
        <FORM name=ff action=login.asp method=post>     
          <TBODY>     
            <TR bgcolor="ffffff">     
              <TD nowrap class="style1"><font color="#0000FF"><%=Session("MBID")%> </font>您好</TD>     
            </TR>     
            <TR>     
              <TD nowrap class="style1"><a      
			href="../sys/member/login.asp?send=out&USType=mem&DirPag=/pklist.asp" target="_parent">安全登出</a> - <a      
			href="/bbsuser/bu_sign.asp" target="_blank">会员后台</a></TD>     
            </TR>     
            <TR>     
              <TD nowrap class="style1"><a href="/sys/member/service/mu_appread.asp">重新注册</a> - <a      
			  href="/sys/member/service/get_pw.asp">忘记密码</a></TD>     
            </TR>     
          </TBODY>     
        </FORM>     
      </TABLE>     
      <%end if%>
						</TD>
                      </TR>
                      <TR>
                        <TD><IMG height=16 src="pkimg/login_down.gif" 
                width=172></TD>
                      </TR>
                      <TR>
                        <TD height=8></TD>
                      </TR>
                    </TBODY>
                  </TABLE>

                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD><IMG height=56 src="pkimg/s_122.gif" 
              width=172></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD width=1 bgColor=#dbdbdb></TD>
                      <TD vAlign=top><TABLE cellSpacing=0 cellPadding=0 width=160 align=center 
                  border=0>
                          <TBODY>
                            <TR>
                              <TD width=21 height=25><IMG height=9 src="pkimg/s_124.gif" width=9></TD>
                              <TD class=style10 width=139 height=25>无</TD>
                            </TR>
                            <TR bgColor=#dbdbdb>
                              <TD colSpan=2 height=1></TD>
                            </TR>
                            <TR bgColor=#dbdbdb>
                              <TD colSpan=2 height=1></TD>
                            </TR>
                          </TBODY>
                      </TABLE></TD>
                      <TD width=1 bgColor=#dbdbdb></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD><IMG height=14 src="pkimg/s_123.gif" 
              width=172></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD height=7></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD><IMG height=40 src="pkimg/s_139.gif" 
              width=172></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD vAlign=top background=pkimg/s_140.gif><TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                          <TBODY>
                            <TR>
                              <TD height=4></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=156 align=center 
                  border=0>
                          <TBODY>
                            <TR>
                              <TD width=32 height=23><DIV align=center><IMG height=11 
                        src="../img/inum/6001.jpg" width=11></DIV></TD>
                              <TD width=124 height=23><DIV class=style1 align=left>
                                <p>无</p>
                                </DIV></TD>
                            </TR>
                            <TR>
                              <TD colSpan=3><DIV align=center><IMG height=1 
                        src="pkimg/s_142.gif" 
                  width=150></DIV></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                      </TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=172 border=0>
                  <TBODY>
                    <TR>
                      <TD><IMG height=9 src="pkimg/s_153.gif" 
              width=172></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD vAlign=top width=10></TD>
              <TD vAlign=top width=568>
			  
<%

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT TOP 1 * FROM [BPKTitle] ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Left(rs("InfCont"),72)
InfCont = Show_Text(InfCont)
InfCont = Replace(InfCont,"<br>"," &nbsp; ")&"..."
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
SetHot = rs("SetHot")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
end if 
rs.Close()
SET rs=Nothing

  StrDetail = ""
If InfOrg<>"" Then
  StrDetail = "<br>详情:<A href='"&InfOUrl&"' target=_blank><font color=blue>"&InfOrg&"</font></A>"
End If

%>

			  <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD class=style1 width=496><A class="style1 style12" 
                  href="/">东莞论坛</A>&nbsp;&gt;&nbsp;东莞辩论</TD>
                      <TD nowrap><span class="style1"><a href="pkjoin.asp?ID=<%=KeyID%>"><img src="pkimg/s_162.gif" width="59" height="18" border="0"></a></span></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD class=style1 height=18>当前辩题数：<SPAN class=style21>(0)</SPAN> 当前辩论评论数：<SPAN class=style21>(0)</SPAN> 当前辩论参与人数：<SPAN 
                  class=style21>(0)</SPAN></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD><IMG height=71 src="pkimg/s_156.gif" 
              width=568></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD vAlign=top width=276><TABLE cellSpacing=0 cellPadding=0 width=276 border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=9 src="pkimg/s_125.gif" 
                      width=276></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=276 border=0>
                          <TBODY>
                            <TR>
                              <TD vAlign=top background=pkimg/s_126.gif><TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD height=5></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD width=25><IMG height=17 
                              src="pkimg/s_128.gif" width=19></TD>
                                      <TD class=style11 width=225><p><A class=style11 
                              href="pkview.asp?ID=<%=KeyID%>" 
                              target=_blank><%=InfSubj%></A>
                                      </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD height=5></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD><IMG height=1 src="pkimg/s_129.gif" 
                              width=250></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD height=5></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=240 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD class=style1 vAlign=top height=100><A 
                              class=style1 
                              href="pkview.asp?ID=<%=KeyID%>" 
                              target=_blank><%=InfCont%></A> &nbsp; <%=StrDetail%></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                                <TABLE cellSpacing=0 cellPadding=0 width=250 
                        align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD height=10></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=276 border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=9 src="pkimg/s_127.gif" 
                      width=276></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                      <TD width=16></TD>
                      <TD vAlign=top width=276><TABLE cellSpacing=0 cellPadding=0 width=242 
                        align=center border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=7 src="pkimg/s_130.gif" 
                              width=242></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=242 
                        align=center border=0>
                          <TBODY>
                            <TR>
                              <TD bgColor=#d2ebef><TABLE height=134 cellSpacing=0 cellPadding=0 
                              width=242 align=center border=0>
                                  <TBODY>
                                    <TR>
                                      <TD bgColor=#d2ebef><TABLE width="100%" border="0" cellpadding="1" cellspacing="1">
                                        <TR class="style17">
                                          <TD align="right">发起人：</TD>
                                          <TD><%=InfMemb%></TD>
                                        </TR>
                                        <TBODY>
                                          <TR class="style17">
                                            <TD width="50%" align="right">浏览量：</TD>
                                            <TD><%=SetRead%></TD>
                                          </TR>
                                          <TR class="style17">
                                            <TD align="right">正反得票：</TD>
                                            <TD><%=InfVote1%>/<%=InfVote2%></TD>
                                          </TR>
                                          <TR class="style17">
                                            <TD align="right">开始时间：</TD>
                                            <TD><%=InfStart%></TD>
                                          </TR>
                                          <TR class="style17">
                                            <TD align="right">结束时间：</TD>
                                            <TD><%=InfEnd%></TD>
                                          </TR>
                                          <TR align="center" class="style17">
                                            <TD colspan="2"><A 
                                href="pkjoin.asp?ID=<%=KeyID%>&Act=PubView&TP="><IMG 
                                height=27 src="pkimg/s_138.gif" width=97 
                                border=0></A></TD>
                                            </TR>
                                        </TBODY>
                                      </TABLE>
                                      </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=242 
                        align=center border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=7 src="pkimg/s_131.gif" 
                              width=242></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=5></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD background=pkimg/s_132.gif height=25><TABLE cellSpacing=0 cellPadding=0 width=560 border=0>
                          <TBODY>
                            <TR>
                              <TD width=28><DIV align=center><IMG height=15 
                        src="pkimg/s_133.gif" width=15></DIV></TD>
                              <TD class=style10 vAlign=center 
width=481>其它当前进行中的辩论主题</TD>
                              <TD class=style1 vAlign=bottom width=51><A class=style1 
                        href="#">更多&gt;&gt;</A></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=5></TD>
                    </TR>
                  </TBODY>
                </TABLE>
				
				<!----------------->
				
<%

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [BPKTitle] WHERE KeyID <>'"&KeyID&"' ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))
InfCont = Replace(InfCont,"<br>"," &nbsp; ")
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
SetHot = rs("SetHot")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")

Vote1 = rs_Count(conn,"BPKVote WHERE KeyMod='"&KeyID&"' AND KeyFlag='Aegis'") 
Vote2 = rs_Count(conn,"BPKVote WHERE KeyMod='"&KeyID&"' AND KeyFlag='Oppose'")

View1 = rs_Count(conn,"BPKView WHERE KeyMod='"&KeyID&"' AND KeyFlag='Aegis'") 
View2 = rs_Count(conn,"BPKView WHERE KeyMod='"&KeyID&"' AND KeyFlag='Oppose'")

%>
				
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=8><TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=12 src="pkimg/s_134.gif" 
              width=568></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE height=80 cellSpacing=0 cellPadding=0 width=568 border=0>
                          <TBODY>
                            <TR>
                              <TD width=1 bgColor=#d8d8d8></TD>
                              <TD bgColor=#f5f5f5><TABLE cellSpacing=0 cellPadding=0 width=540 align=center 
                  border=0>
                                  <TBODY>
                                    <TR>
                                      <TD class=style11 colSpan=2 height=25><A class=style11 
                              href="pkview.asp?ID=<%=KeyID%>" ><%=InfSubj%></A></TD>
                                      <TD class=style1 width=154 height=25><A 
                        href="pkjoin.asp?ID=<%=KeyID%>&Act=PubView&TP="><IMG 
                        height=21 src="pkimg/s_155.gif" width=80 
                        border=0></A></TD>
                                    </TR>
                                    <TR>
                                      <TD colSpan=3><IMG height=1 src="pkimg/s_136.gif" 
                        width=544></TD>
                                    </TR>
                                    <TR>
                                      <TD class=style1 width=186>发起人：<SPAN 
                        class=style12><%=InfMemb%></SPAN></TD>
                                      <TD class=style1 width=205><STRONG 
                        class=style1>正反方得票：<SPAN 
                        class=style12><%=Vote1%>/<%=Vote2%></SPAN></STRONG></TD>
                                      <TD class=style1><SPAN class=style22><STRONG 
                        class=style1><STRONG 
                        class=style1>发起时间</STRONG>：<STRONG class=style1><SPAN 
                        class=unnamed1><STRONG class=style1><STRONG 
                        class=style1><SPAN 
                        class=style21><%=InfStart%></SPAN></STRONG></STRONG></SPAN></STRONG></STRONG></SPAN></TD>
                                    </TR>
                                    <TR>
                                      <TD colSpan=3><IMG height=1 src="pkimg/s_136.gif" 
                        width=544></TD>
                                    </TR>
                                    <TR>
                                      <TD class=style1><STRONG 
                        class=style1><SPAN 
                        class=style22>访问量：</SPAN><SPAN 
                        class=style12><%=SetRead%></SPAN><STRONG 
                        class=style1></STRONG></STRONG></TD>
                                      <TD class=style1><STRONG 
                        class=style1>正反方评论：</STRONG><STRONG 
                        class=style1><STRONG class=style1><SPAN 
                        class=style21><%=View1%>/<%=View2%></SPAN></STRONG></STRONG></TD>
                                      <TD><SPAN class=style1><STRONG 
                        class=style1><SPAN class=style22><STRONG 
                        class=style1>结束时间</STRONG></SPAN>：<SPAN 
                        class=style21><%=InfEnd%></SPAN></STRONG></SPAN></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                              <TD width=1 bgColor=#d8d8d8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                          <TBODY>
                            <TR>
                              <TD><IMG height=12 src="pkimg/s_135.gif" 
              width=568></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                          <TBODY>
                            <TR>
                              <TD height=5></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
				
<%
rs.MoveNext
Loop
end if 
rs.Close()
SET rs=Nothing

%>
				
				<!----------------->
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD background=pkimg/s_137.gif height=25><TABLE cellSpacing=0 cellPadding=0 width=560 border=0>
                          <TBODY>
                            <TR>
                              <TD width=28><DIV align=center><IMG height=15 
                        src="pkimg/s_133.gif" width=15></DIV></TD>
                              <TD class=style10 vAlign=center width=481><STRONG 
                        class=style10 style="FONT-SIZE: 12px"><FONT 
                        color=#333333>热门辩论</FONT></STRONG></TD>
                              <TD class=style1 vAlign=bottom width=51><A class=style1 
                        href="#">更多&gt;&gt;</A></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=5></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                无
                <TABLE cellSpacing=0 cellPadding=0 width=568 border=0>
                  <TBODY>
                    <TR>
                      <TD height=5></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
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
