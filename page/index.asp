<!--#include file="_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=vPMsg_WHome%>-<%=vPMsg_WName%><%=subjseo_Home%></title>
<meta name="keywords" content="<%=keysseo_Home%>" />
<meta name="description" content="<%=contseo_Home%>" />
<link rel="Shortcut" href="/favicon.ico" />
<link rel="Bookmark" href="/favicon.ico" />
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../pfile/pub/jsload.js" type="text/javascript"></script>
<script src="../ext/api/play/jsPlay02.js" type="text/javascript"></script>
<script src="../inc/home/jquery.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="_head.asp"-->
<%	
'//echo Request.QueryString()
sql = "SELECT TOP 4 * FROM [InfoPics] WHERE KeyMod IN('PicS124') AND SetShow='Y' AND LEN(ImgName)>15 ORDER BY LogATime DESC " ' AND SetHot='Y'
rp1Arr = get_Player(sql,"iview.asp?KeyID=($KeyID)")			
rp1Img=rp1Arr(0)
rp1Url=rp1Arr(1)
rp1Txt=rp1Arr(2)
rp1Width = 230
rp1Height = 175
rp1TxtH = 24
rp1SwfH = rp1TxtH+rp1Height
rpPlayer = "null" ':bcastr,playp02,playp03,null
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
                <option value="">[选类别]</option>
                <option value="info.asp?ModID=PicS124">产品图片</option>
                <option value="info.asp?ModID=InfN124">新闻中心</option>
              </select></td>
            <td><input type="submit" name="button" id="button" value="搜索"/></td>
          </tr>
        </form>
      </table></td>
    <td width="12" align="center">&nbsp;</td>
    <td align="center" width="490">&nbsp;</td>
  </tr>
</table>
<div class="line08">&nbsp;</div>
<div class="pgArea" style="text-align:center">
  <div class="pgBorder PubLeft pgW240">
    <div class="line05">&nbsp;</div>
    <!--#include file="../ext/api/play/playpic.asp"-->
  </div>
  <div class="pgBorder PubLeft pgW450" style="margin-left:8px;">
    <div class="Inf1Tbg">
      <div class="PubLeft Inf1Tit">新闻中心</div>
      <div class="PubRight Inf1TMore"><a href="info.asp?ModID=InfN124">More</a></div>
    </div>
    <div style=" height:175px;">
      <div class="PubClear"></div>
      <%
	  sTmp = vbcrlf& "<ul class='ItmPub Itm01'>"
	  sTmp = sTmp& "<li class='PubLeft'><a href='iview.asp?KeyID=($KeyID)'>($InfSubj)</a></li>"
	  sTmp = sTmp& "<li class='ItmTime PubRight'>($LogATime)</li></ul>"
	  Response.Write ListPub(sTmp,20,"SELECT TOP 7 KeyID,InfSubj,SetSubj,LogATime FROM [InfoNews] WHERE KeyMod='InfN124' AND SetShow='Y' ORDER BY SetTop,LogATime DESC")
	  'Response.Write ListTemp(sTmp,2)
	  %>
    </div>
  </div>
  <div class="pgBorder PubRight pgW240">
    <div class="Inf1Tbg" style="text-align:center"> <span id="SwhItem01" class="SwhItmA" onmousemove="setTimeout('SwhChang(1)',300)">衣(01)</span> <span id="SwhItem02" class="SwhItmA" onmousemove="setTimeout('SwhChang(2)',300)">食(02)</span> <span id="SwhItem03" class="SwhItmA" onmousemove="setTimeout('SwhChang(3)',300)">住(03)</span> </div>
    <div id="SwhIDiv01" style="height:175px;visibility:visible">
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv01 SwhIDiv01</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">Test Subject</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
    </div>
    <div id="SwhIDiv02" style="height:175px;visibility:hidden;display:none;">
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv02 A</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv02 B</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
    </div>
    <div id="SwhIDiv03" style="height:175px;visibility:hidden;display:none;">
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv03 A</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv03 B</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
      <ul class="ItmPub Itm02">
        <li class="PubLeft">SwhIDiv03 C</li>
        <li class="ItmTime PubRight">2009-12-31</li>
      </ul>
    </div>
  </div>
</div>

<div class="PubClear"></div>
<div class="line08">&nbsp;</div>
<div class="pgArea">
  <div class="pgBorder PubLeft pgW240">
    <div class="Inf1Tbg">
      <div class="PubLeft Inf1Tit">上滚动</div>
      <div class="PubRight Inf1TMore">More</div>
    </div>
    <div class="PubClear"></div>
    <div id="mTop01" style="overflow:hidden; width:240px; height:175px;">
      <div id="mTop02">
        <!--///////可编辑区 Start////////-->
        <%

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
          <li class="PubLeft"><a href="info.asp?ModID=<%=sMD%>&TypID=<%=aCode(i)%>" class="as01" ><%=aName(i)%></a></li>
        </ul>
        <%

    End If
  Next
%>
        <!--///////可编辑区 End////////-->
      </div>
    </div>
  </div>
  <div class="pgBorder PubLeft pgW240" style="margin-left:8px;">
    <div class="Inf1Tbg">
      <div class="PubLeft Inf1Tit">下滚动</div>
      <div class="PubRight Inf1TMore">More</div>
    </div>
    <div class="PubClear"></div>
    <div id="mBot00" style='overflow:hidden;height:175px;width:240px; position:relative;'>
      <div id="mBot01" style="position:relative;overflow:hidden;">
        <!--///////可编辑区 Start////////-->
        <%
 sql = "SELECT TOP 12 KeyID,InfSubj,SetSubj FROM [InfoNews] WHERE KeyMod='InfN124' AND SetShow='Y' ORDER BY SetTop,LogATime DESC"
 rs.Open sql,conn,1,1 
 If Not rs.EOF Then 
 Do While NOT rs.EOF
   KeyID = rs("KeyID")
   InfSubj = Show_SLen(rs("InfSubj"),24)
   SetSubj = rs("SetSubj")
   InfSubj = Show_sTitle(InfSubj,SetSubj)
%>
        <ul class="ItmPub Itm02">
          <li class="PubLeft"><a href='iview.asp?KeyID=<%=KeyID%>'><%=InfSubj%></a></li>
          <li class="ItmTime PubRight"><%=LogATime%></li>
        </ul>
        <%
 rs.MoveNext
 Loop
 End If
 rs.Close()
%>
        <!--///////可编辑区 End////////-->
      </div>
      <div id="mBot02" style="position:relative;"></div>
    </div>
  </div>
  <div class="pgBorder PubRight pgW450">
    <div class="Inf1Tbg">
      <div class="Inf1Tbg">
        <div class="PubLeft Inf1Tit">左滚动</div>
        <div class="PubRight Inf1TMore">More</div>
      </div>
    </div>
    <div style="height:75px; border-bottom:1px solid #b6d9f7; padding-left:9px">
      <div id="mLeft0Div0" style="OVERFLOW: hidden; WIDTH: 430px; COLOR: #ffffff;">
        <div class="line10">&nbsp;</div>
        <table cellspacing="0" cellpadding="0" align="left" border="0" cellspace="0">
          <tr>
            <td id="mLeft0Div1" valign="top"><!--///////可编辑区 Srart////////-->
              <table height="55" border="0" cellpadding="1" cellspacing="1" bgcolor="#E0E0E0" id="mLeft0JQ1">
                <tr bgcolor="#FFFFFF">
                  <%
sql = "SELECT TOP 8 * FROM [InfoPics] WHERE KeyMod='PicS124' AND SetShow='Y' AND LEN(ImgName)>15 ORDER BY LogATime DESC " 
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 sID = rs("KeyID")
 sSBak = rs("InfSubj")
 sSubj = Show_SLen(sSBak,10)
 sSubj = Replace(sSubj,"'","")
 sImg = get_1Img(sID,rs("ImgName")) 
 sTime = FormatDateTime(rs("LogATime"),2)
%>
                  <td nowrap="nowrap" align="center"><a href="iview.asp?KeyID=<%=sID%>" target="_blank"> <img src="<%=sImg%>" height="50" hspace="5" border="0" alt="<%=InfSBak%>" /></a></td>
                  <td nowrap="nowrap" align="center"><a href="iview.asp?KeyID=<%=sID%>" target="_blank"> <img src="<%=sImg%>" height="50" hspace="5" border="0" alt="<%=InfSBak%>" /></a></td>
                  <%
rs.MoveNext
Loop
rs.Close()
%>
                </tr>
              </table>
              <!--///////可编辑区 End////////--></td>
            <td id="mLeft0Div2" valign="top">　</td>
          </tr>
        </table>
      </div>
    </div>
    <div class="Inf1Tbg">
      <div class="Inf1Tbg">
        <div class="PubLeft Inf1Tit">右滚动</div>
        <div class="PubRight Inf1TMore">More</div>
      </div>
    </div>
    <div style="height:70px; padding-left:8px">
      <div id="mRight0Div0" style="OVERFLOW: hidden; WIDTH: 430px; COLOR: #ffffff">
        <div class="line10">&nbsp;</div>
        <table cellspacing="0" cellpadding="0" align="left" border="0" cellspace="0">
          <tr>
            <td id="mRight0Div1" valign="top"><!--///////可编辑区 Srart////////-->
              <table width="430" height="50"  border="0" cellpadding="1" cellspacing="1" bgcolor="#E0E0E0">
                <tr bgcolor="#FFFFFF">
                  <td width='80' nowrap>信息1 信息1 信息1 信息1 </td>
                  <td width='80' nowrap>信息2 信息2 信息2 信息2 </td>
                  <td width='80' nowrap>信息3 信息3 信息3 信息3 </td>
                  <td width='80' nowrap>信息4 信息4 信息4 信息4 </td>
                  <td width='80' nowrap>信息5 信息5 信息5 信息5 </td>
                  <td width='80' nowrap>信息6 信息6 信息6 信息6 </td>
                  <td width='80' nowrap>信息7 信息7 信息7 信息7 </td>
                  <td width='80' nowrap>信息8 信息8 信息8 信息8 </td>
                </tr>
              </table>
              <!--///////可编辑区 End////////--></td>
            <td id="mRight0Div2" valign="top">　</td>
          </tr>
        </table>
      </div>
    </div>
  </div>
</div>
<div class="line08">&nbsp;</div>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="100" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle">Demo ... </td>
          <td align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif"><img height="27" src="../pfile/pimg/qqimg21_13.gif" width="13" align="absmiddle" /></td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:3px;"><TABLE width=620 border=0 align=center cellPadding=2 cellSpacing=1>
              <TR>
                <TD align=center><table border="0" align="left" cellpadding="2" cellspacing="1">
                  <form id="fms" name="fms02" method="post" action="../page/osearch.asp" target="_blank">
                    <tr>
                      <td>OrderNO</td>
                      <td><textarea name="KeyWD" rows="4" class="fm" id="KeyWD" style="font-size:12px;width:150px; "></textarea></td>
                      <td width="180" rowspan="2" align="left">注意：可一次填入多个定单号，<br />
                        如输入多个定单号，则一行一个；<br />
                        一次最多可查询10个定单；<br />
                        eg: <a href="ord_view.asp?ID=0BA8A839F7D008237A187B3T" target="_blank">108RB-JK82K-984</a>, <a href="ord_view.asp?ID=0BA89D325A3615FFKH8X25HV" target="_blank">8DA24S-GDW1K6</a>, <a href="ord_view.asp?ID=0BA73B5EBA162DSVN71DFRRF" target="_blank">Ord-0998H-1B8FR</a> 。</td>
                    </tr>
                    <tr>
                      <td>Search</td>
                      <td><input type="submit" name="button2" id="button2" style="width:120px;" value="Order Search"/></td>
                    </tr>
                  </form>
                </table></TD>
                <TD align=center><%

' Show Play02 ====================================================== 
sql = "SELECT TOP 10 * FROM [InfoPics] WHERE KeyMod='PicS124' AND SetShow='Y' AND LEN(ImgName)>15 ORDER BY SetTop,LogATime DESC " 
%>
                  <div id="imgADPlayer"></div>
                  <script type="text/javascript">
 <%=get_Play02(sql,"iview.asp?KeyID=($KeyID)")%>
 setTimeout("jsPlay02.init('imgADPlayer',120,100)",120); 
 //jsPlay02.init( "imgADPlayer", 1003, 290 );
</script></TD>
              </TR>
            </TABLE></td>
        </tr>
      </table></td>
    <td width="300" align="left" valign="top" style="padding:0px 0px 0px 8px;"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="100" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle">文字连接</td>
          <td align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif"><img height="27" src="../pfile/pimg/qqimg21_13.gif" width="13" align="absmiddle" /></td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:3px;"><table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr>
                <td><select name="selhm01" size="1" class="hmOptLink" id="selhm01" onchange="goOptUrl(this,'_self')">
                    <option value="">请选择</option>
                    <option value="http://www.dgchr.com/">dgchr.com</option>
                  </select></td>
                <td><select name="selhm03" size="1" class="hmOptLink" id="selhm03" onchange="goOptUrl(this,'')">
                    <option value="">请选择</option>
                    <option value="http://www.dg.gd.cn/">dg.gd.cn</option>
                  </select></td>
                <td><select name="selhm02" size="1" class="hmOptLink" id="selhm02" onchange="goOptUrl(this,'')">
                  <option value="">请选择</option>
                  <option value="http://www.txjia.com/">txjia.com</option>
                </select></td>
              </tr>
            </table>
            <table width="200" border="0" align="center" cellpadding="0" cellspacing="1">
              <TR>
                <TD align=center id="DateTim11">&nbsp;
                  <script src="../inc/home/jsdate.asp?verNow=1&divObj=DateTim11" type="text/javascript"></script></TD>
              </TR>
              <TR>
                <TD align=center id="DateTim12">&nbsp;
                  <script src="../inc/home/jsdate.asp?verNow=2&divObj=DateTim12" type="text/javascript"></script></TD>
              </TR>
              <TR>
                <TD align=center id="DateTim13">&nbsp;
                  <script src="../inc/home/jsdate.asp?verNow=3&divObj=DateTim13" type="text/javascript"></script></TD>
              </TR>
              <TR>
                <TD align=center id="DateTim14">&nbsp;
                  <script src="../inc/home/jsdate.asp?verNow=4&divObj=DateTim14" type="text/javascript"></script></TD>
              </TR>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="_foot.asp"-->
<div class="PubClear"></div>

<script type="text/javascript">
$(document).ready(function(){ 
  //alert("hello"); 
  //$("#picShowBig").hide("fast"); 
  //$("#picShowBig").show("slow"); 
  //$("#picShowBig").fadeOut(); 
  //$("#picShowBig").fadeIn(); 
  //$("#mLeft0JQ1").fadeIn(); 
  //$("#mTop02").slideUp(3000); //300,callback
});
</script>

<script 'xxx-' src="../pfile/pub/jshome.js" type="text/javascript"></script>
<!-- Adv Banner Start -->
<div class="PubClear line08">&nbsp;</div>
<!-- Adv Banner End -->
</body>
</html>
