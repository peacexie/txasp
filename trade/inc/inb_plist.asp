<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <%
sql = " SELECT * FROM "&ModTab&" "&sqlK
sql =sql& " ORDER BY SetTop,LogATime DESC "
rs.Open Sql,conn,1,1
rs.PageSize = 1200
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If
if rs.eof then 
%>
  <tr>
    <td height="30" align="center"><%=vPMsg_NoRec%></td>
  </tr>
  <% else %>
  <tr>
    <td align="center"><%    
rs.AbsolutePage = Page
for ix = 1 to rs.PageSize 'Do While Not rs.eof 
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
aImgs = Show_sfImgs(ImgName,KeyID)
%>
      <div class="ItemPic">
        <table width="168" border="0" align="center" cellpadding="2" cellspacing="1" style="border:1px #F0F0F0 solid;">
          <tr align="center" valign="middle">
            <td width="160" height="120"><a href="?ModID=<%=MD%>&TypID=<%=TP%>&UsrID=<%=US%>&Flag=Cont&ID=<%=KeyID%>" target="_blank"><img src="<%=aImgs(1)%>" width="150" height="120" border="0" onload="javascript:setImgSize(this);" /></a></td>
          </tr>
          <tr align="left">
            <td align="center"><a href="?ModID=<%=MD%>&TypID=<%=TP%>&UsrID=<%=US%>&Flag=Cont&ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a></td>
          </tr>
        </table>
      </div>
      <%
rs.movenext
  if rs.eof then 
   exit for
  end if     
next 'loop	  
%>
    </td>
  </tr>
  <!--
  <tr>
    <td height="30"><%=RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"",1)%></td>
  </tr>
  -->
  <%
end if
rs.close()
%>
</table>
