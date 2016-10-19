<%If Len(ID)>15 Or Flag<>"" Then%>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td height="500" align="left" valign="top" style="padding:12px;">
	  <%If Flag="Cont" Then%>
	  <!-- #include file="ina_cont.asp" -->
      <%ElseIf Flag="FAQ" Then %>
      <!-- #include file="inb_faq.asp" -->
      <%ElseIf Flag="NTab" Then %>
      <!-- #include file="inb_nlist.asp" -->
      <%ElseIf Flag="PTab" Then %>
      <!-- #include file="inb_plist.asp" -->
      <%ElseIf MD="TraT124" Then %>
      <!-- #include file="inc_pics.asp" -->
      <%End If%>
    </td>
  </tr>
</table>
<%Else%>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td height="500" align="left" valign="top" style="padding:12px;">
	   <!--表格列表-->
	   <%If MD="TraJ124" AND US<>"" Then%>
	   <!-- #include file="inc_job.asp" -->
       <%Else%>
       <!-- #include file="inc_news.asp" -->
	   <%End If%>
	</td>
  </tr>
</table>
<%End If%>