<!--#include file="binc/_config.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%
ID = RequestS("ID",3,48)
TP = RequestS("TP",3,48)
Act = RequestS("Act",3,48)

  If Session("MBID")&""="" Then 
    'DirPag = "/pkview.asp?ID="&ID&"" 
	'Response.Write js_Alert("请先登陆!","Redir","/mem_login.asp?goMod=ComPK&DirPag="&DirPag) 
	'Response.End()
  End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT TOP 1 * FROM [BPKTitle] WHERE KeyID='"&ID&"' AND InfEnd>=#"&Date()&"# ",conn,1,3 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj2 = Show_Text(rs("InfSubj"))
InfCont2 = Show_Text(rs("InfCont"))
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
Vote1 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Aegis'") 
Vote2 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Oppose'")
Vote3 = rs_Count(conn,"BPKVote WHERE KeyMod='"&ID&"' AND KeyFlag='Neutral'")
View1 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Aegis'") 
View2 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Oppose'")
View3 = rs_Count(conn,"BPKView WHERE KeyMod='"&ID&"' AND KeyFlag='Neutral'")
Pnt1 = View1*5 + Vote1*2
Pnt2 = View2*5 + Vote2*2
Pnt3 = View3*5 + Vote3*2
else
Response.Write js_Alert("错误,数据不存在或已经停止投票!","Redir","/pklist.asp")
end if 
rs.Close()
SET rs=Nothing
  
If Act="PutVote" Then
  If rs_Exist(conn,"SELECT KeyID FROM BPKVote WHERE KeyMod='"&ID&"' AND LogAUser='"&Session("MBID")&"' AND LogATime>'"&Date()&"' ")="YES" Then
    Response.Write js_Alert("您今天已经过投票!","Redir","pkview.asp?ID="&ID&"")
  Else
sql = " INSERT INTO [BPKVote] (" 
sql = sql& "  KeyID, KeyMod, KeyFlag" 
sql = sql& ", LogAddIP, LogAUser, LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & get_AutoID(24) &"'" 
sql = sql& ", '" & ID &"'" 
sql = sql& ", '" & TP &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("MBID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
Call rs_DoSql(conn,sql)  
  End If
  Response.Write js_Alert("感谢您投票!","Redir","pkview.asp?ID="&ID&"")
End If

send = RequestS("send",3,48)

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

If send="send" Then
KeyID = get_AutoID(24)
InfSubj = RequestS("InfSubj"&sys27_Rnd(1),3,255)
InfCont = RequestS("InfCont"&sys27_Rnd(2),3,4000)
InfSubj = Chr_Fil2(InfSubj) 
InfCont = Chr_Fil2(InfCont) 
sql = " INSERT INTO [BPKView] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyFlag" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfMemb" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ID &"'" 
sql = sql& ", '" & RequestS("KeyFlag",3,24) &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfCont &"'" 
sql = sql& ", '" & Session("MBID") &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("MBID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  If InfSubj<>"" And InfCont<>"" Then
Call rs_DoSql(conn,sql) 
  End If
Response.Write js_Alert("感谢您评论!","Redir","/pkview.asp?ID="&ID&"") '.Redirect "info_list.asp"
End If

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
            href="/pklist.asp">辩论首页</A> &gt;&gt; <%=InfSubj2%></TD>
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
                      <TD class=style8 width=502><%=InfSubj2%></TD>
                      <TD width=152><span class="style1">                        <a href="/pkview.asp?ID=<%=KeyID%>">返回</a> </span></TD>
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
                      <TD class=style1 
                  width=420><A 
                              class=style1 
                              href="/pkview.asp?ID=<%=KeyID%>" 
                              target=_blank><%=InfCont2%></A> &nbsp; <%=StrDetail%></TD>
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
                    <td valign="top" bgcolor="#FEFAE7"><table width="364" height="54" border="0" cellpadding="0" cellspacing="0">
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
                        </table></td>
                        <td width="1" bgcolor="#BDBDBD"></td>
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
                </table></TD>
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
                </table></TD>
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
              <TD vAlign=top width=744 bgColor=#f5f5f5>
			  <%
			  If Act="PubView" Then
  'If rs_Exist(conn,"SELECT KeyID FROM BPKView WHERE KeyMod='"&ID&"' AND LogAUser='"&Session("MBID")&"' ")="YES" Then
    'Response.Write js_Alert("您已发表过评论!","Redir","pkview.asp?ID="&ID&"")
  'End If
			    If TP="Aegis" Then
				  InfSubj = InfView1
				ElseIf TP="Oppose" Then
				  InfSubj = InfView2
				Else
				  InfSubj = "RE:"&InfSubj
				End If
			  %>
                <table width="720" border="0" align="center" cellpadding="1" cellspacing="1">
                  <form name="fm01" method="post" action="?">
                    <tr class="style1">
                      <td width="20%" align="right">观点标题</td>
                      <td><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="48" value="<%=InfSubj%>">
                      </td>
                    </tr>
                    <tr class="style1">
                      <td align="right">观点立场</td>
                      <td><select name="KeyFlag" id="KeyFlag">
					  <option value="Neutral" selected>[中立方]</option>
					  <option value="Aegis"  <%If TP="Aegis"  Then Response.Write("selected")%>>**正方**</option>
					  <option value="Oppose" <%If TP="Oppose" Then Response.Write("selected")%>>**反方**</option>
                      </select></td>
                    </tr>
                    <tr class="style1">
                      <td align="right">观点内容</td>
                      <td><textarea name="InfCont<%=sys27_Rnd(2)%>" cols="48" rows="4" id="InfCont<%=sys27_Rnd(2)%>"></textarea></td>
                    </tr>
                    <tr class="style1">
                      <td align="right"><input name="send" type="hidden" id="send" value="send">
                      <input name="ID" type="hidden" id="ID" value="<%=ID%>">
                      <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
                      </td>
                      <td><input type="button" name="Button" value="提交" onClick="chkData()"></td>
                    </tr>
                  </form>
                </table>
				<%End If%>
				
			  <%
			  If Act="MemList" Then
			  %>
                <table width="720" border="0" align="center" cellpadding="1" cellspacing="1">
                 
                    <tr class="style1">
                      <td width="20%" align="right">&nbsp;</td>
                      <td colspan="4">评论支持人列表</td>
                    </tr>

                <%
SET rs=Server.CreateObject("Adodb.Recordset") 
If TP<>"" Then
  sqlK = " AND KeyFlag='"&TP&"'"
End If
rs.Open "SELECT * FROM [BPKView] WHERE KeyMod='"&ID&"' "&sqlK&" ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
RecCount = rs.RecordCount
Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfMemb = rs("InfMemb")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogAddIP = rs("LogAddIP")
If KeyFlag="Aegis" Then
  KeyFlag = "正方"
ElseIf KeyFlag="Oppose" Then 
  KeyFlag = "反方"
Else
  KeyFlag = "中立方"
End If
%>
                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td><%=KeyFlag%></td>
                      <td><%=InfMemb%></td>
                      <td><%=LogAddIP%></td>
                      <td><%=LogATime%></td>
                    </tr>
                <%
rs.MoveNext
Loop
%>
                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td colspan="4">共计: <%=RecCount%>人</td>
                    </tr>
<%
else
%>

                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td colspan="4">暂无辩论资料</td>
                    </tr>
                <%
end if 
rs.Close()
SET rs=Nothing
%>
                    <tr class="style1">
                      <td colspan="5" align="center"><hr></td>
                    </tr>
					
					
                    <tr class="style1">
                      <td width="20%" align="right">&nbsp;</td>
                      <td colspan="4">投票支持人列表</td>
                    </tr>

                <%
SET rs=Server.CreateObject("Adodb.Recordset") 
If TP<>"" Then
  sqlK = " AND KeyFlag='"&TP&"'"
End If
rs.Open "SELECT * FROM [BPKVote] WHERE KeyMod='"&ID&"' "&sqlK&" ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
RecCount = rs.RecordCount
Do While Not rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
'InfMemb = rs("InfMemb")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogAddIP = rs("LogAddIP")
If KeyFlag="Aegis" Then
  KeyFlag = "正方"
ElseIf KeyFlag="Oppose" Then 
  KeyFlag = "反方"
Else
  KeyFlag = "中立方"
End If
%>
                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td><%=KeyFlag%></td>
                      <td><%=LogAUser%></td>
                      <td><%=LogAddIP%></td>
                      <td><%=LogATime%></td>
                    </tr>
                <%
rs.MoveNext
Loop
%>
                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td colspan="4">共计: <%=RecCount%>人</td>
                    </tr>
<%
else
%>

                    <tr class="style1">
                      <td align="right">&nbsp;</td>
                      <td colspan="4">暂无辩论资料</td>
                    </tr>
                <%
end if 
rs.Close()
SET rs=Nothing
%>
					
					
                    <tr class="style1">
                      <td align="right">&nbsp;                      </td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
              
                </table>
				<%End If%>

				
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

<script type="text/javascript">

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length>1200) 
   {   
     alert(" 说明 不能超过1200字!"); 
     document.fm01.InfCont<%=sys27_Rnd(2)%>.value = document.fm01.InfCont<%=sys27_Rnd(2)%>.value.substr(0,1200-3)+"...";
     document.fm01.InfCont<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }
 //tmv = chkF_Date(document.fm01.InfMemb,"日期 不规范！");
 //if (tmv=='ER') 
   {         //alert(""); //document.fm01.VSStart.focus();
     //eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</BODY>
</HTML>
