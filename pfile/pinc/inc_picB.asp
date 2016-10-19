<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <%
sql = " SELECT * FROM "&ModTab&" "&sqlK
sql =sql& " ORDER BY SetTop,LogATime DESC "
'Response.Write sql
rs.Open Sql,conn,1,1
rs.PageSize = 6
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
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead") 
ImgName = rs("ImgName") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser") 
InfLink = "" 
aImgs = Show_sfImgs(ImgName,KeyID)
Call get_TmpLink()
%>
      <div class="ItemPicB">
        <table border="0" align="center" cellpadding="2" cellspacing="1" style="border:1px #F0F0F0 solid;">
          <tr align="center" valign="middle">
            <td width="190" height="250"><a <%=InfLink%> target="_blank"><img src="<%=aImgs(1)%>" width="180" height="240" border="0" onload="javascript:setImgSize(this);" /></a></td>
          </tr>
          <tr align="left">
            <td align="center"><a <%=InfLink%> target="_blank"><%=InfSubj%></a></td>
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
  <tr>
    <td height="30"><%=RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"&DepID="&DP&"",1)%></td>
  </tr>
  <%
end if
rs.close()
%>
</table>
