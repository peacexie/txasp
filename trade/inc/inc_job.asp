<%
    sql = " SELECT * FROM "&ModTab&" "&sqlK
	sql =sql& " ORDER BY SetTop,LogATime DESC "
	'Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 1200
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>

<table width="100%"  border="0" align="center" cellpadding="2" cellspacing="1">
  <tr>
    <td><%
  If not rs.eof then
  i=0
  Do While NOT rs.EOF
  i=i+1
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
SetSubj = rs("SetSubj")
InfSubj = rs("InfSubj") : BakSubj = InfSubj
InfSubj = Left(rs("InfSubj")&"",128)
InfSubj = Show_sTitle(InfSubj,SetSubj)
xxxCont = Show_Html(rs("InfCont"))
SetRead = rs("SetRead")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogETime = rs("LogETime")
If inStr(tmsListTab(12),MD)>0 Then
  MDJoin = MD
ElseIf inStr(tmsListTab(12),TP)>0 Then
  MDJoin = TP
Else
  MDJoin = "InfJ124"
End If
 
	  %>
      <a name="<%=KeyID%>"></a>
      <table width="100%" height="27"  border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
        <tr>
          <td align="left" bgcolor="#F0F0F0" style="cursor:hand;" onclick="ShowDiv('<%=Show_jsKey(KeyID)%>')">&nbsp;<B> (<%=i%>) <%=InfSubj%></B></td>
          <td width="20%" align="center" bgcolor="#F0F0F0">&nbsp;<a href="gbook.asp?ModID=<%="TraA124&ID="&KeyID&"&UsrID="&US&"&ObjSubj="&Server.URLEncode(BakSubj)%>" target="_blank" class="cRed"><%=vJob_Apply%></a>&nbsp;</td>
        </tr>
      </table>
      <div style="line-height:1px;">&nbsp;</div>
      <div id="<%=Show_jsKey(KeyID)%>" style="display:none; ">
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#F0F0F0">
          <tr>
            <td align="left" bgcolor="#FFFFFF" class="SysCont" style="padding:5px 13px; "><%=Show_sfData(KeyID,"fcont.htm")%></td>
          </tr>
        </table>
        <div style="line-height:2px;">&nbsp;</div>
      </div>
      <%
  rs.Movenext
  Loop
%>
      <%  
  Else
  %>
      无信息
      <%
  End If
	  rs.Close()
	  
	  %>
    </td>
  </tr>
</table>
