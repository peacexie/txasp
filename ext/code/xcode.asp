
<%If TP="" Then%>
<tr>
  <td style="padding:18px 0px 8px 100px;">
  <!-- #include file="../pfile/pinc/inc_ptyp.asp" -->
  </td>
</tr>
<%Else%>
<tr>
  <td style="padding:8px 0px 8px 8px;">
  <!-- #include file="../pfile/pinc/inc_pics.asp" -->
  </td>
</tr>
<%End IF%>

<%


'set conn= Server.CreateObject("adodb.connection")
'connstr="provider=microsoft.jet.oledb.4.0;data source="&server.MapPath("\script\#dbf#\PeaceWeb.mdb")
'connstr="provider=microsoft.jet.oledb.4.0;data source="&server.MapPath("\u\demo\upfile\#dbf#\ysWeb_0BA8F85C-DAF6-00CA-2275-PEACE0ASP03E.Peace!DB")
'conn.open connstr


r = Date()&"@"&Timer()
Response.Write r
' Sub-Domain ====================================================== 


' Player Pic4 ======================================================
'"iview.asp?KeyID="&sID&"
sql = "SELECT TOP 4 * FROM [InfoPics] WHERE KeyMod IN('PicS124') AND SetShow='Y' AND LEN(ImgName)>15 ORDER BY LogATime DESC " ' AND SetHot='Y'
rp1Arr = get_Player(sql,"iview.asp?KeyID=($KeyID)")			
rp1Img=rp1Arr(0)
rp1Url=rp1Arr(1)
rp1Txt=rp1Arr(2)
'-------------------------
rp1Width = 230
rp1Height = 175
rp1TxtH = 24
rp1SwfH = rp1TxtH+rp1Height
%>

<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="<%=rp1Width%>" height="<%=rp1SwfH%>">
  <param name="movie" value="../api/play/playpic.swf">
  <param name=wmode value=transparent>
  <param name="quality" value="high">
  <param name="menu" value="false">
  <param name=wmode value="opaque">
  <param name="FlashVars" value="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>">
  <embed src="../api/play/playpic.swf" FlashVars="pics=<%=rp1Img%>&links=<%=rp1Url%>&texts=<%=rp1Txt%>&borderwidth=<%=rp1Width%>&borderheight=<%=rp1Height%>&textheight=<%=rp1TxtH%>" wmode="opaque" menu="false" bgcolor="#DADADA" quality="high" width="<%=rp1Width%>" height="<%=rp1SwfH%>" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
</object>
<script type="text/javascript">
var nplayer_focus_width=232;
var nplayer_focus_height=145;
var nplayer_text_height=0;
var nplayer_swf_height=nplayer_focus_height+nplayer_text_height ;
var nplayer_pics_list='<%=rp1Img%>';
var nplayer_links_list='<%=rp1Url%>';
var nplayer_texts_list='<%=rp1Txt%>';
var pics='<%=rp1Img%>'; 
var links='<%=rp1Url%>';
var texts='<%=rp1Txt%>';
  document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ nplayer_focus_width +'" height="'+ nplayer_swf_height +'">');
  document.write('<param name="movie" value="../inc/home/playpic.swf">');
  document.write('<param name="quality" value="high">');
  document.write('<param name="bgcolor" value="#ffffff">');
  document.write('<param name="menu" value="false">');
  document.write('<param name=wmode value="TRANSPARENT">'); 
  document.write('<param name="FlashVars" value="pics='+nplayer_pics_list+'&links='+nplayer_links_list+'&texts='+nplayer_texts_list+'&borderwidth='+nplayer_focus_width+'&borderheight='+nplayer_focus_height+'&textheight='+nplayer_text_height+'">');
  document.write('<embed src="../inc/home/playpic.swf" wmode="opaque" FlashVars="pics='+nplayer_pics_list+'&links='+nplayer_links_list+'&texts='+nplayer_texts_list+'&borderwidth='+nplayer_focus_width+'&borderheight='+nplayer_focus_height+'&textheight='+nplayer_text_height+'" menu="false" bgcolor="#ffffff" quality="high" width="'+ nplayer_focus_width +'" height="'+ nplayer_swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
  document.write('</object>');  
</script>
-----------------
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" id=scriptmain name=scriptmain codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="319" height="248">
  <param name="movie" value="../_bcastr.swf?bcastr_xml_url=_bcastr.asp?a=<%=Timer()%>">
  <param name="quality" value="high">
  <param name=scale value=noscale>
  <param name="LOOP" value="false">
  <param name="menu" value="false">
  <param name="wmode" value="transparent">
  <embed src="../_bcastr.swf?bcastr_xml_url=_bcastr.asp?a=<%=Timer()%>" width="319" height="248" loop="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" salign="T" name="scriptmain" menu="false" wmode="transparent"></embed>
</object>
<br />
_bcastr.asp:
<?xml version="1.0" encoding="utf-8"?>
<bcaster autoPlayTime="3">
<%
sql = "SELECT TOP 8 * FROM [JK_Picture] "
sql = sql& " ORDER BY LEN(PTTop),PTTop,PTID DESC "
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
  PTID = rs("PTID")
  PTSubj = rs("PTSubj")
  ImgName = rs("ImgName")
  imgStr = "/script/modjk/pic/"&ImgName ''show_Media("","/script/modjk/pic/"&ImgName,150,80,ImgScale,"L")
  lnkStr = "/joke/pic_detail.asp?PTID="&PTID
  subStr = Replace(PTSubj,"""","")
  Response.Write vbcrlf&"<item item_url="""&imgStr&""" link="""&lnkStr&""" itemtitle="""&subStr&"""></item>"
rs.MoveNext
Loop
rs.Close
%>  
</bcaster>
-----------------
<%

' Show Play02 ====================================================== 

sql = "SELECT TOP 10 * FROM [InfoPics] WHERE KeyMod='PicT124' AND SetShow='Y' AND InfType LIKE 'T110016%' AND LEN(ImgName)>15 ORDER BY SetTop,LogATime DESC " 
sRool02 = get_Play02(sql,"iview.asp?KeyID=($KeyID)")

%>
<script src="../api/play/jsPlay02.js" language="JavaScript1.2"></script>
<div id="imgADPlayer"></div>
<script type="text/javascript">
 <%=sRool02%>
 setTimeout("jsPlay02.init('imgADPlayer',1003,290)",120); 
 //jsPlay02.init( "imgADPlayer", 1003, 290 );
</script>
<%

' Show Cont1 ====================================================== 
tCont = rs_Val("","SELECT InfCont FROM InfoNews WHERE InfType like '%N110012;%'")			  


' Show Pic1 ======================================================

sql = "SELECT TOP 1 KeyID+ImgName AS sPic FROM [InfoNews] WHERE KeyMod='InfN124' AND SetShow='Y' AND InfType LIKE 'N110032%' AND SetHot='Y' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
sPic = rs_Val("",sql) : aPic = Split(sPic&"^","^") 
sKey = aPic(0) : sPic = get_1Img(sKey,sPic)
'aRow = rs_Row("",sql)

sPic = "<a href='iview.asp?KeyID="&sKey&"'><img src='"&sPic&"' width='203' height='200'></a>" 

' Show ListA ====================================================== 

sTmp1 = vbcrlf& "<ul class='ItmPub Itm01'>"
sTmp1 = sTmp1& "<li class='PubLeft'><a href='pview.asp?KeyID=($KeyID)'>($InfSubj)</a></li>"
sTmp1 = sTmp1& "<li class='ItmTime PubRight'>($LogATime)</li></ul>"
sData1 = ListPub(sTmp1,28,"SELECT TOP 7 KeyID,InfSubj,SetSubj,LogATime FROM [InfoPics] WHERE KeyMod='PicS124' ORDER BY LogATime DESC")

' Show ListB ====================================================== 

sql = "SELECT TOP 6 * FROM [InfoNews] WHERE KeyMod='InfN124' AND SetShow='Y' AND InfType LIKE 'N110012%' ORDER BY SetTop,LogATime DESC " 
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 sID = rs("KeyID")
 sSubj = Show_SLen(rs("InfSubj"),120)
 sSet = rs("SetSubj")
 sSubj = Show_sTitle(sSubj,sSet)
 sTime = FormatDateTime(rs("LogATime"),2)
%>
<div> <img src="../images/index_19.jpg" width="12" height="12"> <a href="iview.asp?KeyID=<%=sID%>"><%=sSubj%></a> [<%=sTime%>] </div>
<%
rs.MoveNext
Loop
rs.Close()

' Scrool Pic1 ====================================================== 

sql = "SELECT TOP 8 * FROM [InfoPics] WHERE KeyMod='PicT124' AND SetShow='Y' AND InfType LIKE 'T110016%' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 sID = rs("KeyID")
 sBak = rs("InfSubj")
 sSubj = Show_SLen(sBak,10)
 sSubj = Replace(sSubj,"'","")
 sSet = rs("SetSubj")
 sSubj = Show_sTitle(sSubj,sSet)
 sImg = get_1Img(sID,rs("ImgName"))
 'LogATime = FormatDateTime(rs("LogATime"),2)
%>
<div> <a href="iview.asp?KeyID=<%=sID%>" target="_blank"> <img src="<%=sImg%>" height="80" hspace="1" border="0" alt="<%=sBak%>" />
  <div class="line05">&nbsp;</div>
  <%=sSubj%></a> </div>
<%
rs.MoveNext
Loop
rs.Close()

' Show TypList ====================================================== 

   'sMD = "InfA124"
   'sTmp=GetItemPart(sMD,"A110048")
   'aTmp=Split(sTmp,"^")
   'aCode = Split(aTmp(0),"|")
   'aName = Split(aTmp(1),"|")
   'aNam2 = Split(aTmp(2),"|")

   sMD = "PicS124"
   aCode = Split(Eval("s"&sMD&"Code"),"|")
   aName = Split(Eval("s"&sMD&"Name"),"|")
   aNam2 = Split(Eval("s"&sMD&"Nam2"),"|")
  For i=0 to uBound(aCode)
    If aCode(i)<>"" Then
      iFlag = aNam2(i)
%>
<ul class="ItmPub Itm02">
  <li class="PubLeft"> <a href="iview.asp?ModID=<%=sMD%>&amp;TypID=<%=aCode(i)%>" class="as01" ><%=aName(i)%></a> </li>
</ul>
<%
    End If
  Next

' Data-Cont ====================================================== 

sql = "SELECT TOP 8 * FROM [InfoPics] WHERE KeyMod='PicT124' AND SetShow='Y' AND InfType LIKE 'T110016%' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
rs.Open sql,conn,1,1 
If NOT rs.EOF Then
 sID = rs("KeyID")
 sSubj = rs("InfSubj")
 sCont = rs("InfCont") 
 sCont = Replace(sCont,"&nbsp;"," ")
 sCont = Show_HText(sCont,30)&"…"
 sImg = get_1Img(sID,rs("ImgName"))
End If
rs.Close() 


' Link Pic1 ======================================================
          
sql = "SELECT TOP 3 * FROM [InfoPics] WHERE KeyMod='PicTBak' AND InfType LIKE 'TB10012%' AND LEN(ImgName)>15 ORDER BY SetTop,LogATime DESC " 
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 sID = rs("KeyID")
 sBak = rs("InfSubj")
 sSubj = Show_SLen(sBak,10)
 sSubj = Replace(sSubj,"'","")
 sSet = rs("SetSubj")
 sSubj = Show_sTitle(sSubj,sSet)
 sImg = get_1Img(sID,rs("ImgName"))
%>
<tr>
  <td height="50" align="center" valign="bottom"><a href="iview.asp?KeyID=<%=sID%>" target="_blank"> <img src="<%=sImg%>" width="273" height="45" hspace="0" border="0" /></a></td>
</tr>
<%
rs.MoveNext
Loop
rs.Close()


' Side Cont ====================================================== 

sfName = Get_fName() ': Response.Write sfName

If Left(sfName,5)="about" Or sfName="contact.asp" Then

ElseIf sfName="pic2.asp" Then

ElseIf sfName="pic.asp" Or sfName="pview.asp" Then

ElseIf sfName="news.asp" Or sfName="nview.asp" Then

ElseIf sfName="gbook.asp" Then

ElseIf sfName="service.asp" Then

Else

End If


' 导入资料 ====================================================== 


 con6 = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="&Server.MapPath(Config_Path&"upfile/#dbf#/OLD_Product.Peace!DB" ) 
 sql6 = "SELECT * FROM [xxxTable] " :i=0
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open sql6,con6,1,1 
 Do While NOT rs.EOF
   i=i+1
   iID = rs("ID") 
   iSubj = rs("Subj") 
   '...
   iType = rs_Val(con6,"SELECT smallclass FROM [smallclass] WHERE ID="&iClass&"")
   iType = rs_Val(conn,"SELECT TypLayer FROM [WebTyps] WHERE TypName='"&iType&"'")
	If rs_Exist(conn,"SELECT KeyID FROM [uuuTable] WHERE LogAddIP='"&iID&"'")="EOF" Then
      sql = " INSERT INTO [uuuTable] (" 
      sql = sql& "  KeyID" 
      sql = sql& ", InfSubj" 
	  '... 
      sql = sql& ", LogAddIP, LogAUser, LogATime" 
      sql = sql& ")VALUES(" 
      sql = sql& "  '" & iID &"'" 
      sql = sql& ", '" & iSubj & "'" 
	  '...
      sql = sql& ", '" & iID & "', '" & Session("UsrID") &"', '" & Now() &"'" 
      sql = sql& ")" ':Response.Write slq
	  Call rs_DoSql(conn,sql)
	  Response.Write vbcrlf&"<br>----"&i&" - "&iSubj&" - "&"<br>"
	End If
 rs.MoveNext
 Loop
 rs.Close()
 Set rs = Nothing
%>
