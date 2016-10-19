<link href="../tools/mzoom/lbox-style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
  lbox_BasePath = "../tools/mzoom/";
</script>
<script src="../tools/mzoom/lbox-script.js" type="text/javascript"></script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <%
sql = " SELECT * FROM "&ModTab&" "&sqlK
sql =sql& " ORDER BY SetTop,LogATime DESC "
'Response.Write sql
rs.Open Sql,conn,1,1
rs.PageSize = 5
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
InfSubj = rs("InfSubj") : bakSubj = InfSubj
InfSubj = Show_sTitle(InfSubj,SetSubj)
xxxCont = rs("InfCont") 
InfCont = Show_sfRead(KeyID,"/fcont.htm")
InfCont = Show_HText(Replace(InfCont,"&nbsp;"," "),120)&"…"
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead") 
ImgName = rs("ImgName") 
LogATime = rs("LogATime")
LogAUser = rs("LogAUser") 
InfLink = "" 
aImgs = Show_sfImgs(ImgName,KeyID)
Call get_TmpLink()
%>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" style="border:1px #F0F0F0 solid;">
          <tr align="center" valign="middle">
            <td width="210" height="160" rowspan="4">
            <!--a <%=InfLink%> target="_blank"-->
            <a title="<%=akSubj%>" href="<%=get_1Img(KeyID,ImgName)%>" rel="lightbox">
            <img src="<%=aImgs(1)%>" width="200" height="150" border="0" onload="javascript:setImgSize(this);" /></a></td>
            <td align="left" valign="top"><strong><%=vPic_Name%></strong>: <a <%=InfLink%> target="_blank"><%=InfSubj%></a></td>
          </tr>
          <tr align="left">
            <td align="left" valign="top"><strong><%=vPic_Code%></strong>: <%=KeyCode%></td>
          </tr>
          <tr align="left">
            <td height="80" align="left" valign="top"><strong><%=vPMsg_RDetail%>:</strong> <%=InfCont%> </td>
          </tr>
          <tr align="left">
            <td align="right" valign="top"><%=vPMsg_RDetail%>&gt;&gt;&gt;&nbsp;</td>
          </tr>
        </table>
        <div class="line10">&nbsp;</div>
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
