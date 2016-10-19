
<%
    sql = " SELECT * FROM "&ModTab&" "&sqlK
	sql =sql& " ORDER BY SetTop,LogATime DESC "
	'Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 12
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>

<table width="100%"  border="0" align="center" cellpadding="2" cellspacing="1">

  <tr class="fnt666">
    <th height="27" nowrap>&nbsp;</th>
    <th align="left" nowrap><%=vInf_Subj%></th>
    <th width="5%" align="center" nowrap><%=vInf_Type%></th>
    <th width="5%" align="center" nowrap><%=vInf_Read%></th>
    <th width="15%" align="center" nowrap><%=vInf_Pub%></th>
  </tr>
  <tr>
    <td colspan="5" nowrap bgcolor="#999999" class='td1px'>&nbsp;</td>
  </tr>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
TypName = GetTName(KeyMod,InfType,"")
SetSubj = rs("SetSubj")
InfSubj = rs("InfSubj")
InfSubj = Show_sTitle(InfSubj,SetSubj)
InfPara = rs("InfPara")
SetRead = rs("SetRead") 
ImgName = rs("ImgName") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser") 
InfLink = "" 
Call get_TmpLink() 
	  %>
  <tr>
    <td width="3%" height="24" align="center">&middot;</td>
    <td> <a <%=InfLink%> target="_blank"><%=InfSubj%></a> </td>
    <td width="5%" align="center" nowrap><%=TypName%></td>
    <td width="5%" align="center" nowrap>&nbsp;<%=SetRead%>&nbsp;</td>
    <td width="15%" align="center" nowrap><%=LogATime%></td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <tr>
    <td colspan="5" nowrap bgcolor="#999999" class='td1px'>&nbsp;</td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="5" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"&DepID="&DP&"",1)%></td>
  </tr>
  <%  
  Else
  %>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="5"><%=vPMsg_NoRec%></td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  
	  %>
  <%If SwhGuestPub="Y" Then%>
  <tr>
    <form name="fmGuestLogin" id="fmGuestLogin" action="../smod/info/guest_login.asp" method="post" target="_blank">
    <td height="30" colspan="5" align="right" style="padding-right:20px">
     <input name="ModID" type="hidden" value="<%=MD%>" />
     <input name="TypID" type="hidden" value="<%=InfType%>" />
     <input name="act" type="hidden" value="in" />
     <input name="ChkSes" type="hidden" value="<%=Hex(Session.SessionID)%>" />
     <input name="ChkUrl" type="hidden" value="<%=Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")%>" />
     <input name="<%=App29Code%>" id="<%=App29Code%>" type="hidden" value="<%=Now()%>" />
     <div class="guest_up"><a href="#" target="_self" onclick="GstLogin('<%=App29Code%>')">我要上传</a></div>
    </td>
    </form>
  </tr>
  <%End If%>
</table>


