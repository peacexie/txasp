<%
	sql = " SELECT * FROM [GboInfo] "&sqlK&" AND SetShow='Y' "
	sql =sql& " ORDER BY KeyID DESC " ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 8 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If
%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="8">
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize

KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfSubj = Left(rs("InfSubj"),30)
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
xxxCont = Show_Html(rs("InfCont"))
xxxReply = Show_Html(rs("InfReply"))
InfReply = Show_sfRead(KeyID,".rep.htm") 'Show_sfGbook(KeyID,".rep.htm")
'Response.Write InfReply
If InfReply&""="" Then
  InfReply = "<font color=gray>"&vGbo_RepNO&"</font>"
End If
LnkName = Show_Text(rs("LnkName"))
LnkFace = rs("LnkFace")
LnkEmail = Show_Text(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
If LnkEmail<>"" Then
  LnkEmail2 = "email.gif" 
Else
  LnkEmail2 = "email2.gif"   
End If
If LnkName="" Then
  LnkName = "<font color='CCCCCC'>"&vGbo_NoName&"</font>"   
End If
	  %>
  <tr>
    <td align="center"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="tdbg">
        <tr>
          <td width="15%" align="center" valign="top"><%=vGbo_Subj%></td>
          <td width="40%" align="left" valign="top"><font color="#0000FF"><%=InfSubj%></font></td>
          <td width="15%" align="center" valign="top"><%=vGbo_Pub%></td>
          <td align="center" valign="top"><%=LogATime%></td>
        </tr>
        <tr>
          <td align="center" valign="top"><%=vGbo_Name%></td>
          <td align="left" valign="top"><%=LnkName%></td>
          <td align="center" valign="top"><%=vGbo_Type%></td>
          <td align="center" valign="top"><%=TypName%></td>
        </tr>
        <tr>
          <td align="center" valign="top"><%=vGbo_Email%></td>
          <td align="left" valign="top"><%=LnkEmail%></td>
          <td align="center" valign="top">IP</td>
          <td align="center" valign="top"><%=LogAddIP%></td>
        </tr>
        <tr>
          <td colspan="4" align="left" valign="top" style="padding:8px;"><fieldset style="paddingt: 5px; border:#D0D0D0 1px solid">
              <legend class="fnt666"><%=vGbo_Cont%></legend>
              <table width="100%" border="0" cellpadding="2" cellspacing="1">
                <tr>
                  <td class="SysCont" style="padding:8px;">
				  <%Call Show_sfGbook(KeyID,".org.htm")%>
				  </td>
                </tr>
              </table>
            </fieldset></td>
        </tr>
        <tr>
          <td colspan="4" align="left" valign="top" style="padding:8px;">
            <fieldset style="paddingt: 5px; border:#D0D0D0 1px solid">
              <legend class="fnt666"><%=vGbo_Reply%></legend>
              <table width="100%" border="0" cellpadding="2" cellspacing="1">
                <tr>
                  <td class="SysCont" style="padding:8px;">
				  <%=InfReply%>
                  </td>
                </tr>
              </table>
            </fieldset>
            </td>
        </tr>
      </table></td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <tr align="center">
    <td height="27" align="center"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <%  
  Else
  %>
  <tr align="center">
    <td colspan="3"><%=vPMsg_InfNull%></td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  
	  %>
</table>
