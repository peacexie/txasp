

<table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6595D6">
  <tr bgcolor="#6595D6">
    <td height="24" colspan="2" nowrap background="/bbsimg/bbs_mut15.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><font color="#FFFFFF"><b>&nbsp;推荐PK(辩论)</b></font></td>
            <td align="right">&nbsp;</td>
          </tr>
    </table></td>
    <td width="35%" nowrap background="/bbsimg/bbs_mut15.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><font color="#FFFFFF"><b>&nbsp;最新PK(辩论)</b></font></td>
        <td align="right">&nbsp;<font color="#FFFFFF"> <a href="/pklist.asp" target="_blank">更多...</a>&nbsp;</font></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="45%" valign="top" bgcolor="#FFFFFF"><table width='100%' align='center' cellspacing='1'>
<%

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT TOP 1 * FROM [BPKTitle] ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
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
SetUBB = rs("SetUBB")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
ImgAlign = rs("ImgAlign")
ImgTitle = rs("ImgTitle")
ImgWidth = rs("ImgWidth")
ImgHeight = rs("ImgHeight")
ImgScale = rs("ImgScale")
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
      <tr>
        <th height="21" bgcolor="#FFFFFF"><A class=style11 
                              href="/pkview.asp?ID=<%=KeyID%>" 
                              target=_blank><%=InfSubj%></A></th>
      </tr>
      <tr>
        <td height="21"><span class="style1"><%=InfCont%></span></td>
      </tr>
      <tr>
        <td height="21" bgcolor="#FFFFFF"><span class="style10">&nbsp;<img src="pkimg/s_175.gif" width="18" align="absmiddle">正方:<%=InfView1%></span></td>
      </tr>
      <tr>
        <td height="21" bgcolor="#FFFFFF"><span class="style10">&nbsp;<img src="pkimg/s_176.gif" width="18" align="absmiddle">反方:<%=InfView2%></span></td>
      </tr>

    </table></td>
    <td bgcolor="#FFFFFF"><table width='100%' align='center' cellspacing='1'>

      <tr bgcolor='<%=col%>'>
        <td height="21" nowrap bgcolor='<%=col%>'><TABLE width="100%" border="0" cellpadding="1" cellspacing="1">
          <TR class="style17">
            <TD width="50%" align="right" nowrap>发起人:</TD>
            <TD nowrap><%=InfMemb%></TD>
          </TR>
          <TBODY>
            <TR class="style17">
              <TD align="right" nowrap>开始:</TD>
              <TD nowrap><%=InfStart%></TD>
            </TR>
            <TR class="style17">
              <TD align="right" nowrap>结束:</TD>
              <TD nowrap><%=InfEnd%></TD>
            </TR>
            <TR align="center" class="style17">
              <TD colspan="2"><A 
                                href="pkjoin.asp?ID=<%=KeyID%>&Act=PubView&TP="><IMG 
                                height=27 src="pkimg/s_138.gif" width=97 
                                border=0></A></TD>
            </TR>
          </TBODY>
        </TABLE></td>
      </tr>

    </table></td>
    <td width="35%" valign="top" bgcolor="#FFFFFF"><table width='100%' align='center' cellspacing='1'>

      <tr>
        <td height="21" nowrap>
<%

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT TOP 6 * FROM [BPKTitle] WHERE KeyID <>'"&KeyID&"' ORDER BY KeyID DESC ",conn,1,1 
if NOT rs.eof then 
Do While NOT rs.EOF
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
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
SetUBB = rs("SetUBB")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
ImgAlign = rs("ImgAlign")
ImgTitle = rs("ImgTitle")
ImgWidth = rs("ImgWidth")
ImgHeight = rs("ImgHeight")
ImgScale = rs("ImgScale")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")


%>
 &middot;<A class=style11 href="/pkview.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></A><br>
<%
rs.MoveNext
Loop
end if 
rs.Close()
SET rs=Nothing

%>
		</td>
      </tr>

    </table></td>
  </tr>
</table>
