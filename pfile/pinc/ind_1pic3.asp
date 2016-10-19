<div class="SysCPad SysCont">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr>
      <td align="left" valign="top"><table width="100%"  border="0" cellpadding="1" cellspacing="2">
          <form action="ocar.asp" method="post" name="fm01">
            <tr>
              <td width="20%" align="right"><%=vPic_Name%>:</td>
              <td><%=InfSubj%>
                <input name="Act" type="hidden" id="Act" value="PAdd" />
                <input name="ID" type="hidden" id="ID" value="<%=ID%>" />
                <input name="MD" type="hidden" id="MD" value="<%=MD%>" />
                <input name="iSubj" type="hidden" id="iSubj" value="<%=iSubj%>" /></td>
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
          </form>
        </table></td>
      <td width="50%" height="200" align="center" style="padding:8px"><table border="0" cellpadding="0" cellspacing="0" xbgcolor="#F0F0F0">
          <tr>
            <td colspan="3" align="center" id="ImgDiv" style="padding-bottom:5px"><a href='<%=aImgs(1)%>' target="_blank"><img src='<%=aImgs(1)%>' width="300" height="240" border="0" xonload="javascript:setImgSize(this);" /></a></td>
          </tr>
          <tr>
            <td align="center" valign="top"><a style="cursor:pointer;"><img src="<%=aImgs(1)%>" width="90" border="0" onclick="ImgSet('<%=aImgs(1)%>','1')" /></a></td>
            <td align="center" valign="top"><a style="cursor:pointer;"><img src="<%=aImgs(2)%>" width="90" border="0" onclick="ImgSet('<%=aImgs(2)%>','2')" /></a></td>
            <td align="center" valign="top"><a style="cursor:pointer;"><img src="<%=aImgs(3)%>" width="90" border="0" onclick="ImgSet('<%=aImgs(3)%>','3')" /></a></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan="2"><%Call Show_sfData(ID,"fcont.htm")%></td>
    </tr>
  </table>
</div>
<div style="padding:8px">
  <fieldset style="padding:1px 1px 10px 15px">
    <legend><%=vPic_RPic%></legend>
    <%
sql3 = "SELECT * FROM [InfoPhoto] WHERE KeyRe='"&ID&"' ORDER BY SetTop,LogATime DESC " '  TOP 3 AND SetHot='Y'
'Set rs=Server.CreateObject("Adodb.Recordset")
rs.Open sql3,conn,1,1
If NOT rs.EOF Then
Do While NOT rs.EOF
 KeyID3 = rs("KeyID")
 InfSubj3 = Show_SLen(rs("InfSubj"),12)
 InfSubj3 = Replace(InfSubj3,"'","")
  ImgNamR = rs("ImgName")&"" 
  If ImgNamR="" Then
    ImgNamR = Config_Path&"img/tool/no_pic_160120.jpg"
  Else
    ImgNamR = Replace(Config_Path&"upfile/"&ImgNamR,"//","/")
  End If
%>
    <div class="ItemRPic"> <a href="<%=ImgNamR%>" target="_blank"><img src="<%=ImgNamR%>" width="150" height="120" border="0" onload="javascript:setImgSize(this);"></a>
      <div style="padding:3px 1px"><%=InfSubj3%></div>
    </div>
    <%
rs.MoveNext
Loop
Else
%>
    <div class="ItemRPic"> <span class="fntF00"><%=vPMsg_NoRec%></span> </div>
    <%
End If
rs.Close()

%>
  </fieldset>
</div>
<script type="text/javascript">

var fm = document.fm01;

function AddItem(xID) 
{   
  if(window.showModalDialog) 
  {   
    var sRtn; var i,resLen;
    resRtn = showModalDialog("pitems.asp?ID="+xID+"&1=1","","center=yes;dialogWidth=240pt;dialogHeight=150pt");
  }else{ 
    alert("Internet Explorer 4.0 or Later Is Required."); }
}  

function ImgSet(xImg,xID)
{   
   
   var img = "<img src='"+xImg+"' width='300' height='240' border='0' />";
   var a = "<a href='"+xImg+"' target='_blank'>"+img+"</a>";
   var f = xImg.indexOf("no_pic_160120.jpg"); //alert(f);
   if( f>0 ){
	   alert("Error!");
   }else{
	   ImgDiv.innerHTML = a;  
   }
} 
</script>
<div style="clear:both"></div>