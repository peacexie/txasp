<!--KeyID; tmpList=Cont,Next,FAQ,NTab,PTab,News,Pics,Jobs,Vdos-->
<%If Len(ID)>15 Or tmpList="Cont" Or tmpList="Next" Then%>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td height="500" align="left" valign="top" style="padding:12px;">
	<!--file:ina_cont.asp-->
	<!-- #include file="ina_cont.asp" --></td>
  </tr>
</table>
<%ElseIf inStr(";Pics;PicA;PicB;PicC;PTab;Vdos;",tmpList)>0 Then %>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td height="500" align="left" valign="top" style="padding:5px 0px 5px 8px;">
	  <!--网格显示--><!--.Peacd.A.Start.-->
	  <%If tmpList="Pics" Then%>
      <!--file:inc_pics.asp-->
	  <!-- #include file="inc_pics.asp" -->
      
      <%ElseIf tmpList="PicA" Then%>
      <!--file:inc_picA.asp-->
	  <!-- #include file="inc_picA.asp" -->
      <%ElseIf tmpList="PicB" Then%>
      <!--file:inc_picB.asp-->
	  <!-- #include file="inc_picB.asp" -->
      <%ElseIf tmpList="PicC" Then%>
      <!--file:inc_picC.asp-->
	  <!-- #include file="inc_picC.asp" -->
      
      <%ElseIf tmpList="PTab" Then%>
      <!--file:inb_plist.asp-->
	  <!-- #include file="inb_plist.asp" -->
      <%ElseIf tmpList="Vdos" Then%>
      <!--file:inc_vdo.asp-->
	  <!-- #include file="inc_vdo.asp" -->
      <%End If%>
      <!--.Peacd.A.End.-->
      </td>
  </tr>
</table>
<%Else%>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
    <td height="500" align="left" valign="top" style="padding:12px;">
	   <!--表格列表--><!--.Peacd.B.Start.-->
	  <%If tmpList="FAQ" Then%>
      <!--file:inb_faq.asp-->
	  <!-- #include file="inb_faq.asp" -->
      <%ElseIf tmpList="NTab" Then%>
      <!--file:inb_nlist.asp-->
	  <!-- #include file="inb_nlist.asp" -->
      <%ElseIf tmpList="Jobs" Then%>
      <!--file:inb_job.asp-->
	  <!-- #include file="inb_job.asp" -->
      <%ElseIf tmpList="News" Then%>
      <!--file:inc_news.asp-->
	  <!-- #include file="inc_news.asp" -->
      <%ElseIf tmpList="Down" Then%>
      <!--file:inc_down.asp-->
	  <!-- #include file="inc_down.asp" -->
      <%End If%>
      <!--.Peacd.B.End.-->
      </td>
  </tr>
</table>
<%End If%>
