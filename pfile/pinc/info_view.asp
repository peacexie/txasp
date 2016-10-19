<a name="_InfSubj"></a>
<div style="padding:0px 12px 0px 12px;">
  <h1><%=InfSubj%></h1>
  <div class="SysSubj"> <%=vInf_Pub%>:<%=LogATime%> &nbsp; &nbsp; <%=vInf_Read%>:<span id='tms_SetRead'><%=SetRead%></span> </div>
</div>
<!--.Peacd.1.Start.-->
<%

If inStr(aImgs(1),"tool/no_pic")>0 Then
  fNone = " display:none; "
Else
  fNone = " "
End If

%>

<%If tmpShow = "1Pics" Then%>
<!--file:ind_1pics.asp-->
<!-- #include file="ind_1pics.asp" -->
<%ElseIf tmpShow = "1Pic3" Then%>
<!--file:ind_1pic3.asp-->
<!-- #include file="ind_1pic3.asp" -->
<%ElseIf tmpShow = "1Vdos" Then%>
<!--file:ind_1vdo.asp-->
<!-- #include file="ind_1vdo.asp" -->
<%ElseIf tmpShow = "1News" Then %>
<div id="pageData" class="SysCPad SysCont">
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>

<%ElseIf tmpShow = "2Pics" Then%>
<!--file:ind_2pics.asp-->
<!-- #include file="ind_2pics.asp" -->

<%ElseIf tmpShow = "6UD" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p align="center" style='<%=fNone%>padding:8px 0px 8px 0px'> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="480" height="360" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%ElseIf tmpShow = "6Left" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p style="<%=fNone%>float:left; padding:8px 18px 8px 0px"> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="240" height="180" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%ElseIf tmpShow = "6Right" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p style="<%=fNone%>float:right; padding:8px 3px 8px 18px"> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="240" height="180" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%End If%>
<!--.Peacd.1.End.-->
<%
'Response.Write tmpShow
If inStr(";1News;6UD;6Left;6Right;",tmpShow)>0 Then
%>
<!--分页插件Start-->
<div style="clear:both;"></div>
<div id="pageOne" style="display: none" class="SysCPad SysCont">[pageOne]</div>
<div style="clear:both;"></div>
<TABLE id="pageBox" border="0" align=center cellSpacing=0 style="MARGIN: 12px auto 12px auto; display: none">
  <TR>
    <TD id="pageBar" class="pageCell">
    [pageBar copy@peace.xie.ys 2010-07-08]</TD>
  </TR>
</TABLE>
<script src="../inc/home/jsPager.js" type="text/javascript"></script>
<!--分页插件End-->
<%End If%>
<div class="SysCBot">&nbsp;</div>
<%If get_TmpID("Vote",MD&";"&TP)="Y" Then%>
<!--心情投票插件Start-->
<!-- #include file="../../tools/nmood/form.asp" -->
<!--心情投票插件End-->
<%End If%>
<div class="SysELeft"> 
  <% ' tmsNextTab(6) ' 上下篇模版
  tmpNext = get_TmpID("Next",MD&";"&TP)
  If tmpNext="X" Then 
  Else
    If tmpNext="T" Then 
	  nType=TP
	Else
      nType=""
	End IF
  %>
  <%=vPMsg_InfPrev%> <%=ListPNext(ModTab,MD,nType,LogATime,"<")%><br />
  <%=vPMsg_InfNext%> <%=ListPNext(ModTab,MD,nType,LogATime,">")%> 
  <%End If%>
</div>
<div class="SysERight"> 
  <%
  If get_TmpID("Rem",MD&";"&TP)="Y" Then
    Response.Write rmUrl&"<br>"
  End If
  %>
  <a href="javascript:window.close()">[<%=vInf_Close%>]</a> </div>
<div style="clear:both"></div>
