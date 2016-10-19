<!--#include file="../../upfile/sys/config/MTypList.asp"-->
<%
 aCode = Split(sDocD124Code,"|")
 aName = Split(sDocD124Name,"|")
 nSpan = rs_Count(conn," AdmSyst WHERE SysType='Inner' ") 'uBound(aCode)-1
 nSpan = nSpan + 2
%>
<table width="100%" border="0" cellpadding="0" cellspacing="1" id="docType" style="border-style:outset; border:1px solid #094C8D;">
  <tr valign="middle" bgcolor="#FFFFFF">
    <td id='id_All__User'>类别</td>
    <td align="center" bgcolor="#999999">&nbsp;</td>
    <!--#include file="../../upfile/sys/doc/list_depart.asp"-->
  </tr>
  <tr align="left" valign="top">
    <td colspan="<%=nSpan%>" bgcolor="#FFFFFF"><!--Groups-->
      <div id='div_All__User' style='visibility:xhidden;'>
        <li class='mInnID'><a href='?TG=All__User' class="cDRed">所有公文</a></li> 
        <%	
 for i = 0 to uBound(aCode)-1	
   Response.Write "<li class='mInnID'><a href='?TG=All__User&TP="&aCode(i)&"'>"&aName(i)&"</a></li>"
 next
		%>  
      </div>
      <!--#include file="../../upfile/sys/doc/list_user.asp"--></td>
  </tr>
</table>
<script type="text/javascript">
  setEvent("onmouseover","mTypOver","docType","td");
</script>
