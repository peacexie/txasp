﻿
<%
    sql = " SELECT * FROM "&ModTab&" "&sqlK
	sql =sql& " ORDER BY SetTop,LogATime DESC "
   rs.Open Sql,conn,1,1
   rs.PageSize = 1200
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
    <td colspan="5" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
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
InfPara = rs("InfPara") : aPara = Split(InfPara,"^")
SetRead = rs("SetRead") 
ImgName = rs("ImgName") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser") 
	  %>
  <tr>
    <td width="3%" height="24" align="center">&middot;</td>
    <td> <a href="?ModID=<%=MD%>&TypID=<%=TP%>&UsrID=<%=US%>&Flag=Cont&ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a> </td>
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
    <td colspan="5" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
  </tr>
  <!--
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="5" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"",1)%></td>
  </tr>
  -->
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
</table>


