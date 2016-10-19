<link href="../../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<a name="_InfSubj"></a>
<div style="padding:0px 12px 0px 12px;">
  <h1><%=InfSubj%></h1>
  <div class="SysSubj"> <%=vInf_Pub%>:<%=LogATime%> &nbsp; &nbsp; <%=vInf_Read%>:<span id='tms_SetRead'><%=SetRead%></span> </div>
</div>
<%
TmpID = get_TmpID("Show",MD&";"&TP)
'TmpID = "6UD" '用于测试

If inStr(aImgs(1),"tool/no_pic")>0 Then
  fNone = " display:none; "
Else
  fNone = " "
End If

%>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%If TmpID = "1Pics" Then%>
<!-- #include file="../../pfile/pinc/ind_1pics.asp" -->
<%ElseIf TmpID = "1Pic3" Then%>
<!-- #include file="../../pfile/pinc/ind_1pic3.asp" -->
<%ElseIf TmpID = "1Vdos" Then%>
<!-- #include file="../../pfile/pinc/ind_1vdo.asp" -->
<%ElseIf TmpID = "1News" Then %>
<div class="SysCPad SysCont">
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<%ElseIf TmpID = "6UD" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p align="center" style='<%=fNone%>'> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="480" height="360" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%ElseIf TmpID = "6Left" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p style="<%=fNone%>float:left; padding:8px 18px 8px 0px"> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="240" height="180" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%ElseIf TmpID = "6Right" Then %>
<div id="pageData" class="SysCPad SysCont">
  <p style="<%=fNone%>float:right; padding:8px 3px 8px 18px"> <a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="240" height="180" border="0" onload="javascript:setImgSize(this);" /></a> </p>
  <%Call Show_sfData(ID,"fcont.htm")%>
</div>
<div style="clear:both"></div>
<%End If%>
<%
'Response.Write TmpID
If inStr(";1News;6UD;6Left;6Right;",TmpID)>0 Then
%>
<div style="clear:both;"></div>
<div id="pageOne" style="display: none" class="SysCPad SysCont">[pageOne]</div>
<div style="clear:both;"></div>
<TABLE id="pageBox" border="0" align=center cellSpacing=0 style="MARGIN: 12px auto 12px auto; display: none">
  <TR>
    <TD id="pageBar" class="pageCell"> [pageBar copy@peace.xie.ys 2010-07-08]</TD>
  </TR>
</TABLE>
<script src="../inc/home/jsPager.js" type="text/javascript"></script>
<%End If%>
<div class="SysCBot">&nbsp;</div>
<div style="padding:0px 12px 0px 12px;">
  <table width="100%" border="0">
    <tr>
      <td align="center"><a href="index.asp" class="cDRed">[供求首页]</a></td>
      <td align="center"><a href="corp.asp?UsrID=<%=US%>">[会员首页]</a></td>
      <td align="center"><a href="gbook.asp?ModID=<%="TraR124&ID="&KeyID&"&UsrID="&LogAUser&"&ObjSubj="&Server.URLEncode(InfSubj)%>">[信息评论(<%=rs_Count(conn,"TradeGbook WHERE InfReply='"&KeyID&"'")%>)]</a></td>
      <%If MD="TraT124" Then%>
      <td align="center"><a href="gbook.asp?ModID=<%="TraO124&ID="&KeyID&"&UsrID="&LogAUser&"&ObjSubj="&Server.URLEncode(InfSubj)%>">[订购产品]</a></td>
      <%End If%>
      <%If MD="TraJ124" Then%>
      <td align="center"><a href="gbook.asp?ModID=<%="TraA124&ID="&KeyID&"&UsrID="&LogAUser&"&ObjSubj="&Server.URLEncode(InfSubj)%>">[应聘职位]</a></td>
      <%End If%>
      <td align="center"><a href="javascript:window.close()" class="cRed">[关闭本页]</a></td>
    </tr>
  </table>
</div>
<div style="clear:both"></div>
