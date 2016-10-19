<div class="SysCPad SysCont">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr>
      <td align="left" valign="top" style="padding:8px"><%

'MedPath="/player/media/"
'MedExt=".FLV|.RM|.WMV|.AVI|.MPG|.jpg|"
'flv,swf,
'asf,avi,mpg,wna,wmv,
'mp3,mp4,rm,mid,

MedW=400
MedH=300
If Left(aPara(7),7)<>"http://" Then
  fName = Replace(Config_Path&aPara(7)&"","//","/") 'RequestS("fName",3,240) 'fName
Else
  fName = aPara(7) '// http://www.tudou.com/v/50Gco9ct0IM/v.swf
End If
fNBak = fName&""
If Len(fName)>12 Then
  fExt = lCase(Mid(fName,InStrRev(fName,"."),8)) 'lCase(RequestS("fExt",3,24))
Else
  fExt = ""
End If

If fExt="" or fExt=".swf" Or fExt=".flv" Then
 ' Default
ElseIf fExt=".jpg"  Or fExt=".gif"  Or fExt=".jpeg"_
    Or fExt=".png"  Or fExt=".bmp"  Or fExt=".tif"   Or fExt=".tiff"_
   Then
 fExt="[Pic]"
ElseIf fExt=".doc"  Or fExt=".xls"  Or fExt=".ppt"   Or fExt=".pps"_
    Or fExt=".wps"  Or fExt=".et"   Or fExt=".dpt"_ 
   Then
 fExt="[Office]"
ElseIf fExt=".midi" Or fExt=".wav"  Or fExt=".mp3"_
    Or fExt=".rm"   Or fExt=".ram"  Or fExt=".snd"_
    Or fExt=".aif"  Or fExt=".aiff" Or fExt=".au"_
    Or fExt=".avi"  Or fExt=".wmv"  Or fExt=".mpg"   Or fExt=".mpeg" Or fExt=".mov"_
   Then
 fExt="[Media]" 
Else 'fExt=".rar" Or fExt=".txt" Or fExt=".htm" Or fExt=".html"
 fExt="[Link]"
End If

aPara(4) = Get_BSize(aPara(4))

%>
        <%If fExt="" Then%>
        <table width="360" height="240" 
  border="0" align="center" cellpadding="5" cellspacing="8" bgcolor="#666666">
          <tr>
            <td align="center" bgcolor="#FFFFFF"><%=vVdo_NoFile%>...</td>
          </tr>
        </table>
        <%ElseIf fExt=".swf" Then%>
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=MedW%>" height="<%=MedH%>">
          <param name="movie" value="<%=fName%>" />
          <param name="quality" value="high" />
          <embed src="<%=fName%>" quality="high" type="application/x-shockwave-flash" width="<%=MedW%>" height="<%=MedH%>"></embed>
        </object>
        <%ElseIf fExt=".flv" Then%>
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="<%=MedW%>" height="<%=MedH%>">
          <param name="movie" value="<%=Config_Path%>ext/api/play/playflv.swf" />
          <param name="quality" value="high" />
          <param name="FlashVars" value="vcastr_file=<%=fName%>" />
          <embed src="<%=Config_Path%>ext/api/play/playflv.swf" FlashVars="vcastr_file=<%=fName%>" type="application/x-shockwave-flash" width="<%=MedW%>" height="<%=MedH%>"></embed>
        </object>
        <%ElseIf fExt="[Pic]" Then%>
        <img src="<%=fName%>" width="<%=MedW%>" height="<%=MedH%>" border="0" onload="javascript:setImgSize(this);" />
        <%ElseIf fExt="[Office]" Then%>
        <table width="360" height="240" 
  border="0" align="center" cellpadding="5" cellspacing="8" bgcolor="#CCCCCC">
          <tr>
            <td align="center" bgcolor="#FFFFFF"><%=vVdo_OfficeFile%>...<a href="<%=fName%>" target="_blank"><%=fNBak%></a></td>
          </tr>
        </table>
        <%ElseIf fExt="[Media]" Then  %>
        <object id='pVideo' width='<%=MedW%>' height='<%=MedH%>' codebase='http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701'
  type='application/x-oleobject' standby='加载 Microsoft Windows Media Player 组件...' classid='clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95'>
          <param name="FileName" value="<%=fName%>" />
          <param name="AutoSize" value="0" />
          <param name="BufferingTime" value="5" />
          <param name="DisplaySize" value="2" />
          <param name="ShowDisplay" value="0" />
          <param name="ShowControls" value="1" />
          <param name="ShowStatusBar" value="1" />
        </object>
        <%Else%>
        <table width="360" height="240" 
  border="0" align="center" cellpadding="5" cellspacing="8" bgcolor="#CCCCCC">
          <tr>
            <td align="center" bgcolor="#FFFFFF"><%=vVdo_OpenFile%>...<a href="<%=fName%>" target="_blank"><%=fNBak%></a></td>
          </tr>
        </table>
      <%End If%></td>
      <td width="40%" height="200" align="center" valign="top">
      
<table width="100%"  border="0" cellpadding="1" cellspacing="2">
    <tr>
      <td width="30%" align="right">格式：</td>
      <td align="left"><%=aPara(3)%></td>
    </tr>
    <tr>
      <td align="right">大小：</td>
      <td align="left"><%=aPara(4)%></td>
    </tr>
    <tr>
      <td align="right">片长：</td>
      <td align="left"><%=aPara(5)%> (min)</td>
    </tr>
    <tr>
      <td align="right">主演：</td>
      <td align="left"><%=aPara(8)%></td>
    </tr>
    <tr>
      <td align="right">导演：</td>
      <td align="left"><%=aPara(1)%></td>
    </tr>
    <tr>
      <td align="right">版权：</td>
      <td align="left"><%=aPara(2)%></td>
    </tr>
    <%If Session("MemID")&""="xxx" Then%>
    <%Else%>
    <%End If%>

</table></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan="2"><%Call Show_sfData(ID,"fcont.htm")%></td>
    </tr>
  </table>
</div>
<div style="clear:both"></div>