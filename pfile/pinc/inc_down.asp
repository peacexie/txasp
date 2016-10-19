
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
    <th width="5%" align="center" nowrap><%=vPic_VDown%></th>
    <th width="5%" align="center" nowrap><%=vPic_VType%></th>
    <th width="5%" align="center" nowrap><%=vPic_VSize%></th>
    <th width="15%" align="center" nowrap><%=vInf_Pub%></th>
  </tr>
  <tr>
    <td colspan="6" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
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
InfLink = "" 
Call get_TmpLink() 

fName = Replace(Config_Path&aPara(7)&"","//","/") 
If Len(fName)>12 Then
  fExt = lCase(Mid(fName,InStrRev(fName,"."),8)) 
Else
  fExt = ""
End If
aPara(4) = Get_BSize(aPara(4))

	  %>
  <tr>
    <td width="3%" height="24" align="center">&middot;</td>
    <td> <a href="<%=fName%>" target="_blank"><%=InfSubj%></a> </td>
    <td width="5%" align="center" nowrap><a href="<%=fName%>" target="_blank"><%=vPic_VDown%></a></td>
    <td width="5%" align="center" nowrap><%=fExt%></td>
    <td width="5%" align="center" nowrap>&nbsp;<%=aPara(4)%>&nbsp;</td>
    <td width="15%" align="center" nowrap><%=LogATime%></td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <tr>
    <td colspan="6" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="6" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"&DepID="&DP&"",1)%></td>
  </tr>
  <%  
  Else
  %>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="6"><%=vPMsg_NoRec%></td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  
	  %>
</table>


