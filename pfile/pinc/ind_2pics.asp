<SCRIPT src="../tools/mzoom/magzoom.js" type=text/javascript></SCRIPT>
<LINK media=screen href="../tools/mzoom/magzoom.css" type=text/css rel=stylesheet>
<div class="SysCPad SysCont">
<%
If aPara(4)="0" Then aPara(4)=vPic_Pric2
If aPara(5)="0" Then aPara(5)=vPic_Pric2
%>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr>
      <td width="50%" height="200" align="center" style="padding:8px">
      <a class="MagicZoom MagicThumb" id=MagZzoom1 rel="thumb-change: mouseover" href='<%=aImgs(1)%>' target="_blank">
      <img src='<%=aImgs(1)%>' width="360" height="480" border="0" onload="javascript:setImgSize(this);" /></a></td>
      <td align="left" valign="top"><table with="100%"  border="0" cellpadding="1" cellspacing="2">
            <tr>
              <td width="20%" align="right"><%=vPic_Name%>:</td>
              <td><%=InfSubj%></td>
            </tr>
            <tr>
              <td align="right"><%=vPic_Code%>:</td>
              <td><%=KeyCode%></td>
            </tr>
            <tr>
              <td align="right"><%=vPic_Speci%>:</td>
              <td><%=aPara(3)%></td>
            </tr>
            <tr>
              <td align="right"><%=vPic_From%>:</td>
              <td><%=aPara(2)%></td>
            </tr>
            <tr>
              <td align="right"><%=vPic_Pric3%>:</td>
              <td><%=aPara(5)%></td>
            </tr>
            <tr>
              <td align="right"><%=vPic_Price%>:</td>
              <td><%=aPara(4)%></td>
            </tr>
            <tr>
              <td align="right"><%=vInf_Pub%>:</td>
              <td><%=LogATime%></td>
            </tr>
        </table></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan="2"><%Call Show_sfData(ID,"fcont.htm")%></td>
    </tr>
  </table>
</div>
<div style="clear:both"></div>