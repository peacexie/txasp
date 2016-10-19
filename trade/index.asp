<!--#include file="inc/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>商务与供求-<%=Config_Name%></title>
<link rel="stylesheet" type="text/css" href="inc/style.css"/>
<link rel="stylesheet" type="text/css" href="../img/rnd_nid/box_nid.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../ext/api/play/jsPlayer.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="_head.asp"-->
<%	

ctCorp = rs_Count(conn,"TradeCorp")
ctTrade = rs_Count(conn,"TradeInfo WHERE KeyMod='TraT124'")
ctJobs = rs_Count(conn,"TradeInfo WHERE KeyMod='TraJ124'")
	
sql = "SELECT TOP 4 * FROM [TradeInfo] WHERE KeyMod IN('TraT124') AND LEN(ImgName)>15 ORDER BY LogATime DESC " ' AND SetHot='Y'
rs.Open sql,conn,1,1 
rpImg="":rpUrl="":rpTxt=""
Do While NOT rs.EOF
 KeyID = rs("KeyID")
 InfSubj = Show_SLen(rs("InfSubj"),18)
 InfSubj = Replace(InfSubj,"'","")
 ImgName = get_1Img(KeyID,rs("ImgName")) 
 rpImg = rpImg&ImgName&"|"
 rpUrl = rpUrl&"iview.asp?KeyID="&KeyID&"|"
 rpTxt = rpTxt&InfSubj&"|"
rs.MoveNext
Loop
rs.Close()
If rpUrl<>"" Then
  rpImg=rpImg&"$" : rpImg=Replace(rpImg,"|$","")
  rpUrl=rpUrl&"$" : rpUrl=Replace(rpUrl,"|$","")
  rpTxt=rpTxt&"$" : rpTxt=Replace(rpTxt,"|$","")
End If
rpWidth = 230
rpHeight = 175
rpTxtH = 24
rpSwfH = rpTxtH+rpHeight

sTmp1 = vbcrlf& "<ul class='ItmPub Itm01'>"
sTmp1 = sTmp1& "<li class='PubLeft'><a href='iview.asp?KeyID=($KeyID)'>($InfSubj)</a></li>"
sTmp1 = sTmp1& "<li class='ItmTime PubRight'>($LogATime)</li></ul>"
sTmp2 = vbcrlf& "<ul class='ItmPub Itm02'>"
sTmp2 = sTmp2& "<li class='PubLeft'><a href='iview.asp?KeyID=($KeyID)'>($InfSubj)</a></li>"
sTmp2 = sTmp2& "</ul>" '<li class='ItmTime PubRight'>($LogATime)</li>

%>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="bgLine1">
  <tr>
    <td width="12" align="center">&nbsp;</td>
    <td height="33" align="left"><table border="0" align="left" cellpadding="1" cellspacing="1">
        <form id="fms01" name="fms01" method="post" onsubmit="InnSearch(this)">
          <tr>
            <td>站内搜索：</td>
            <td><input name="KW" type="text" class="fm" id="KW" /></td>
            <td><select name="TP" id="TP" style="width:120px;">
                <option value="ypage.asp?m=s">[会员黄页]</option>
                <option value="supply.asp?m=s">[会员供求]</option>
                <option value="field.asp?m=s">[行业新闻]</option>
                <option value="position.asp?m=s">[企业招聘]</option>
              </select></td>
            <td><input type="submit" name="button" id="button" value="搜索" /></td>
          </tr>
        </form>
      </table></td>
    <td width="12" align="center">&nbsp;</td>
    <td align="right" width="490">
    
    会员:<%=ctCorp%>个； 供求:<%=ctTrade%>个；招聘:<%=ctJobs%>个；
    
    </td>
  </tr>
</table>
<div class="line08">&nbsp;</div>
<div class="pgArea" style="text-align:center">
  <div class="pgBorder PubLeft pgW240">
    <div class="line05">&nbsp;</div>
    <object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="<%=rpWidth%>" height="<%=rpSwfH%>">
      <param name="movie" value="../ext/api/play/playpic.swf">
      <param name=wmode value=transparent>
      <param name="quality" value="high">
      <param name="menu" value="false">
      <param name=wmode value="opaque">
      <param name="FlashVars" value="pics=<%=rpImg%>&links=<%=rpUrl%>&texts=<%=rpTxt%>&borderwidth=<%=rpWidth%>&borderheight=<%=rpHeight%>&textheight=<%=rpTxtH%>">
      <embed src="../ext/api/play/playpic.swf" wmode="opaque" FlashVars="pics=<%=rpImg%>&links=<%=rpUrl%>&texts=<%=rpTxt%>&borderwidth=<%=rpWidth%>&borderheight=<%=rpHeight%>&textheight=<%=rpTxtH%>" menu="false" bgcolor="#DADADA" quality="high" width="<%=rpWidth%>" height="<%=rpSwfH%>" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />
    </object>
  </div>
  <div class="pgBorder PubLeft pgW450" style="margin-left:8px;">
    <div class="Inf1Tbg">
      <div class="PubLeft Inf1Tit">公司新闻</div>
      <div class="PubRight Inf1TMore"><a href="news.asp?ModID=InfN124">More</a></div>
    </div>
    <div style=" height:175px;">
      <div class="PubClear"></div>
      <%
		iSql = "SELECT TOP 7 KeyID,InfSubj,SetSubj,LogAUser,LogATime FROM [TradeInfo] WHERE KeyMod='TraN124' ORDER BY LogATime DESC"
		Response.Write ListInf1(sTmp1,24,"TraN124",iSql)
	  %>
    </div>
  </div>
  <div class="pgBorder PubRight pgW240">
    <div class="Inf1Tbg" style="text-align:center"> <span id="SwhItem01" class="SwhItmA" onmousemove="setTimeout('SwhChang(1)',300)">供求</span> <span id="SwhItem02" class="SwhItmA" onmousemove="setTimeout('SwhChang(2)',300)">新闻</span> <span id="SwhItem03" class="SwhItmA" onmousemove="setTimeout('SwhChang(3)',300)">招聘</span> </div>
    <div id="SwhIDiv01" style="height:175px;visibility:visible">
      <%
		iSql = "SELECT TOP 7 KeyID,InfSubj,SetSubj,LogAUser,LogATime FROM [TradeInfo] WHERE KeyMod='TraT124' ORDER BY LogATime DESC"
		Response.Write ListInf1(sTmp2,16,"TraT124",iSql)
	  %>
    </div>
    <div id="SwhIDiv02" style="height:175px;visibility:hidden;display:none;">
      <%
		iSql = "SELECT TOP 7 KeyID,InfSubj,SetSubj,LogAUser,LogATime FROM [TradeInfo] WHERE KeyMod='TraN124' ORDER BY LogATime DESC"
		Response.Write ListInf1(sTmp2,16,"TraN124",iSql)
	  %>    
    </div>
    <div id="SwhIDiv03" style="height:175px;visibility:hidden;display:none;">
      <%
		iSql = "SELECT TOP 7 KeyID,InfSubj,SetSubj,LogAUser,LogATime FROM [TradeInfo] WHERE KeyMod='TraJ124' ORDER BY LogATime DESC"
		Response.Write ListInf1(sTmp2,16,"TraJ124",iSql)
	  %>    
    </div>
  </div>
</div>
<div class="PubClear"></div>
<div class="line08">&nbsp;</div>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="120" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle">产品供求</td>
          <td width="580" align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif"><div class="PubRight Inf1TMore"><a href="news.asp?ModID=InfN224">More</a></div></td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:3px; padding-bottom:8px;"><div style="height:95px;">
              <div id="mRightDiv0" style="OVERFLOW: hidden; WIDTH: 685px; COLOR: #ffffff; text-align:center">
                <div class="line02">&nbsp;</div>
                <table cellspacing="0" cellpadding="0" align="left" border="0" cellspace="0">
                  <tr>
                    <td id="mRightDiv1" valign="top"><!--///////可编辑区 Srart////////-->
                      <table Xwidth="685" height="90"  border="0" cellpadding="3" cellspacing="1" bgcolor="#E0E0E0">
                        <tr bgcolor="#FFFFFF">
                          <%
sql = "SELECT TOP 8 * FROM [TradeInfo] WHERE KeyMod='TraT124' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
rs.Open sql,conn,1,1 
rpImg="":rpUrl="":rpTxt=""
Do While NOT rs.EOF
 KeyID = rs("KeyID")
 InfSubj = Show_SLen(rs("InfSubj"),8)
 InfSubj = Replace(InfSubj,"'","")
 ImgName = get_1Img(KeyID,rs("ImgName")) 
 LogAUser = rs("LogAUser")
%>
                          <td nowrap="nowrap" valign="bottom"><a href="tview.asp?ID=<%=LogAUser%>&KeyID=<%=KeyID%>" target="_blank"> <img src="<%=ImgName%>" height="90" border="0" /><br />
                            <div class="line05">&nbsp;</div>
                            <%=InfSubj%></a></td>
                          <%
rs.MoveNext
Loop
rs.Close()
%>
                        </tr>
                      </table>
                      <!--///////可编辑区 End////////--></td>
                    <td id="mRightDiv2" valign="top">　</td>
                  </tr>
                </table>
                
              </div>
            </div>

            </td>
        </tr>
      </table></td>
    <td width="240" align="left" valign="top" style="padding:0px 0px 0px 8px;"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="100" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle">文字连接</td>
          <td align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif"><img height="27" src="../pfile/pimg/qqimg21_13.gif" width="13" align="absmiddle" /></td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:3px;"> --- </td>
        </tr>
      </table></td>
  </tr>
</table>
<div class="PubClear line08">&nbsp;</div>
<center>

      <div id='ImgPlayer_AdSBot' style='width:950px; height:60px; overflow:hidden; border:1px solid #CCC'>
      </div><script src='<%=Config_Path%>upfile/sys/xadv/pic_AdSBot.js'></script>
  
</center>
<!--#include file="_foot.asp"-->
<div class="PubClear"></div>
<script src="inc/jshome.js" type="text/javascript"></script>
</body>
</html>
