<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<%

'Call Remote__Test()
'Response.Write Request.ServerVariables("HTTP_HOST")
'ModName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")

ID = Request("ID")
yAct = Request("yAct")
nMin = 0 
nMax = 0

'/////////////////////////////////////////////////////////////
'// 打开配置信息

 sql = " SELECT * FROM [InfoOuter] WHERE KeyID='"&ID&"' "
 sql =sql& " ORDER BY KeyID DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 If NOT rs.EOF Then
KeyMod = rs("KeyMod")
InfSubj = rs("InfSubj")
InfType = rs("InfType") 
InfTyp2 = rs("InfTyp2")
InfUrl = rs("InfUrl")
InfCSet = rs("InfCSet")
InfP1Str = rs("InfP1Str")  
InfP1 = rs("InfP1")
InfP2Str = rs("InfP2Str")  
InfP2 = rs("InfP2")
InfDel1 = rs("InfDel1")
InfDel2 = rs("InfDel2")
InfQ1Str = rs("InfQ1Str") 
InfQ1 = rs("InfQ1")
InfQ2Str = rs("InfQ2Str") 
InfQ2 = rs("InfQ2")
InfRep1 = rs("InfRep1")
InfRep2 = rs("InfRep2")
InfKill = rs("InfKill")

 End If
 Set rs = Nothing
If InfP1Str="" Then InfP1Str="<body"
If InfP2Str="" Then InfP2Str="</body"
If InfQ1Str="" Then InfQ1Str="<body"
If InfQ2Str="" Then InfQ2Str="</body"
doFile = "" : If InfKill="xx01-imp.asp" Then doFile=InfKill


'/////////////////////////////////////////////////////////////
'// 得到内容
If yAct="Insert" Then

  no = Request("no")
  Url = Request("Url")
  id1 = Request("id1")
  id2 = Request("id2")
  Frm = Request("Frm")
  Img = Request("Img")
  KeyID = id1&id2
  
  sHtml=OutPage(Url,InfCSet)&"" : sHbak = sHtml 
  p1=OutSPos(sHtml,InfQ1Str,InfQ1) 
  p2=OutSPos(sHtml,InfQ2Str,InfQ2)
  If p1>0 And p2-p1>0 Then
    sHtml=Mid(sHtml,p1,p2-p1)
	f1 = "Y"
  Else
    sHtml=""
	f1 = "N"
  End If
  
  sTitle = Get_Title(sHbak)
  
  If InfRep1<>"" Then
    sHtml = Get_Rep(sHtml,InfRep1,InfRep2)
  End If
  If InfKill<>"" Then
    sHtml = Get_Rep(sHtml,InfKill,"")
	sTitle = Get_Rep(sTitle,InfKill,"")
  End If
  
  If SwhRemSave="Y" Then
    sHtml = RemoteReplaceUrl(sHtml, upRoot, KeyID)
  End If
  sHtml = Show_Html(xStr)
  If Config_Cont="DB" Then
    xxxCont = sHtml
	xxxCont = Replace(xxxCont,"'","''")
  Else
    xxxCont = ""
  End If
  
  KeyCode = RequestS("KeyCode",3,24)
  SetTop = "888"
  SetHot = "N"
  SetShow = "Y"
  SetSubj = "000000"
  IP = Get_CIP()
  LogATime = DateAdd("s",no,Request("Now"))
  InfPara = PrmFlag
  For i=1 To 96 
   If i=2 Then 
    InfPara = InfPara&"^"&Frm
   Else
    InfPara = InfPara&"^"&RequestS("InfPara"&i,3,120) 
   End If
  Next
  ImgName = ""
  If Img<>"" Then
    Call fold_add9(upRoot, KeyID, 0)
	SaveFileType = Mid(Img, InstrRev(Img, ".") + 1) 
	SaveFileName = Get_FmtID("mdhnsx","")&"_"&Rnd_ID("KEY",4)&"(0."&SaveFileType
	locPath = upRoot&Replace(KeyID,"-","/")&"/"&SaveFileName
	If RemoteSaveFile(locPath, Img, 9600, "(.jpg.gif.jpeg.png.tif.swf.flv)") Then
	  'sHtml = Replace(sHtml,Img,locPath)
	  ImgName = "^"&SaveFileName&"^^^^^^^^"
	  f2 = "Y"
	Else
	  f2 = "N"
	End If
  Else
      f2 = "-"
  End If
  
  sql = " INSERT INTO "&ModTab&" (" 
  sql = sql& "  KeyID" 
  sql = sql& ", KeyMod" 
  sql = sql& ", KeyCode" 
  sql = sql& ", InfType" 
  sql = sql& ", InfTyp2" 
  sql = sql& ", InfSubj"  
  sql = sql& ", InfCont" 
  sql = sql& ", InfPara"
  sql = sql& ", SetSubj" 
  sql = sql& ", SetRead" 
  sql = sql& ", SetHot" 
  sql = sql& ", SetTop" 
  sql = sql& ", SetShow" 
  sql = sql& ", ImgName" 
  sql = sql& ", LogAddIP" 
  sql = sql& ", LogAUser" 
  sql = sql& ", LogATime" 
  sql = sql& ", LogEditIP"
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & KeyID &"'" 
  sql = sql& ", '" & ModID &"'" 
  sql = sql& ", '" & KeyCode &"'" 
  sql = sql& ", '" & InfType &"'" 
  sql = sql& ", '" & InfTyp2 &"'" 
  sql = sql& ", '" & sTitle &"'" 
  sql = sql& ", '" & xxxCont &"'" 
  sql = sql& ", '" & RequestSafe(InfPara,3,8000) &"'" 
  sql = sql& ", '" & SetSubj &"'" 
  sql = sql& ", 0" 
  sql = sql& ", '" & SetHot &"'" 
  sql = sql& ", " & SetTop &"" 
  sql = sql& ", '" & SetShow &"'" 
  sql = sql& ", '" & ImgName & "'" 
  sql = sql& ", '" & IP &"'" 
  sql = sql& ", '" & Get_PUser(PrmFlag) &"'" 
  sql = sql& ", '" & LogATime &"'" 
  sql = sql& ", '(imp_do.asp)'"
  sql = sql& ")"
  If f1 = "Y" Then
    Call rs_DoSql(conn,sql)
	Call add_sfFile()
  End If

  Response.Write "("&id2&")OK("&f1&f2&")"
  Response.End()
  
ElseIf yAct="Detail" Then
  Url = Request("Url") : Url = Replace(Url,"&amp;","&")
  Response.Write "<a href='"&Url&"'>"&Url&"</a>"
  sHtml=OutPage(Url,InfCSet)&"" : sHbak = sHtml 
  p1=OutSPos(sHtml,InfQ1Str,InfQ1) 
  p2=OutSPos(sHtml,InfQ2Str,InfQ2)

ElseIf InfTyp2="1Link" Then
  sHtml=OutPage(InfUrl,InfCSet)&"" : sHbak = sHtml
  p1=OutSPos(sHtml,InfQ1Str,InfQ1) 
  p2=OutSPos(sHtml,InfQ2Str,InfQ2)
Else
  sHtml=OutPage(InfUrl,InfCSet)&"" : sHbak = sHtml
  p1=OutSPos(sHtml,InfP1Str,InfP1) 
  p2=OutSPos(sHtml,InfP2Str,InfP2)
End If 
'Response.Write Show_Text(sHtml)
'Response.Write p1&p2&":"&inStr(sHtml,InfPStr1)



'/////////////////////////////////////////////////////////////
'// 处理 显示内容

If p1>0 And p2-p1>0 Then
  sHtml=Mid(sHtml,p1,p2-p1)
Else
  sHtml="(Error)"
End If


If InfRep1<>"" Then
  sHtml = Get_Rep(sHtml,InfRep1,InfRep2)
End If
If InfKill<>"" Then
  sHtml = Get_Rep(sHtml,InfKill,"")
End If

 
%>


<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" src="../../inc/home/jsInfo.js"></script>
<style type="text/css">
.iBox {
	border:1px solid #EEEEEE;
	padding:0px 0px 0px 0px;
	margin:0px 2px 5px 2px;
}
.PLink {
	width:130px;
	height:110px;
	float:left;
	overflow:hidden;
	text-align:center;
	padding:3px;
	margin:1px;
}
.hand { cursor: pointer }
.PLink { }
.scomm {}
.sOK { color:#030; }
.sNG { color:#FF0; }
</style>
</head>
<body>
<%

If yAct="CText" Or yAct="CHtml" Then
  
  aPart = Split(InfUrl,"/")
  If yAct="CText" Then sHtml = Show_Text(sHtml)

ElseIf yAct="Detail" Then

  sTitle = Get_Title(sHbak)
  If sTitle<>"" Then
    'sTitle = Replace(sTitle,"<","[")
	'Response.Write InfKill
	sTitle = Get_Rep(sTitle,InfKill,"")
    sHtml = "<h1>"&sTitle&"</h1>"&sHtml
  End If
  

ElseIf yAct="PHtml" Then 

  sHtml = Show_Text(sHbak)
  
ElseIf yAct="PText" Then

  Set rExp = New RegExp                ' 建立正则表达式。
    rExp.Global = True                 ' 设置为全局
    rExp.Pattern = "<.*?>|<[^<]*>|<!--[^<]*-->"              ' 设置模式。
    rExp.IgnoreCase = true             ' 设置是否区分大小写。
    sHtml = rExp.Replace(sHbak,"") ' 作替换。
  Set regEx = Nothing

ElseIf yAct="LAll" Then

  sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>([^<]*|<img[^<]*>)</a>")
  sHtml = Replace(sHtml,"||",vbcrlf&"<br>")

ElseIf yAct="LText" Then

 If InfTyp2="PLink" Then
  sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>(<img[^<]*>)</a>")
  sHtml = Show_Text(sHtml)
 Else
  sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>")
  sHtml = Show_Text(sHtml)
 End If
 sHtml = Replace(sHtml,"||",vbcrlf&"<br>")
 
ElseIf InfTyp2="1Link" Then

 'sHtml = Replace(sHtml,"||",vbcrlf&"<br>")

Else ' 连接列表 

prtID = Left(rs_AutID(conn,ModTab,"KeyID",upPart,"1",""),22)
codID = Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",4)
Dim Si,Ai(255),Ui :Ui="02" :Si=""
For i=0 To 240 ',96,300
  Ui = Next_ID(Ui,"02",3)
  Ai(i) = Ui
  Si = Si&Ui&"|"
Next
aPart = Split(InfUrl,"/")

 If InfTyp2="PLink" Then
  sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>(<img[^<]*>)</a>")
  sHtml = Replace(sHtml,"<img ","<img width='120' ")
 Else
  sHtml = Get_HLinks(sHtml,"<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>")
 End If
 a = Split(sHtml,"||") : sHtml=""
 nMin = 0+InfDel1
 nMax = uBound(a)-1-InfDel2
 For i=nMin To nMax
   iUrl = Server.URLEncode(Get_1Url(a(i),"href="))
   iDo = ""&doFile&"?yAct=Insert"
   iDo = iDo&"&no="&i&"&ID="&ID&"&ModID="&ModID&"&KeyCode="&codID&""&Ai(i)&""
   iDo = iDo&"&Url="&iUrl&"&id1="&prtID&"&id2="&Ai(i)&""
   iDo = iDo&"&Now="&Now()&"&Frm="&aPart(2)&""
   iDo = iDo&"&Img="&iImg&""
   iDo = iDo&"&Typ="&InfType&""	
   iDo = "<a href='"&iDo&"'>iDo</a>"
   If InfTyp2="PLink" Then
     iImg = Get_1Url(a(i),"src=")
     iLink = vbcrlf&"<div id='Box"&Ai(i)&"' class='PLink'><input id='Url"&Ai(i)&"' type='radio' checked value='"&iUrl&"'>"
     iLink = iLink&"<a href='?ID="&ID&"&yAct=Detail&Url="&iUrl&"'>View</a>-"
	 iLink = iLink&"<span id='Del"&Ai(i)&"' class='hand' onClick=""reItem('"&Ai(i)&"')"">Del</span>."
	 iLink = iLink&"<span id='No"&Ai(i)&"' class='scomm'>"&(i+1)&"-"&iDo&"</span><br />"
	 iLink = iLink&"<input id='Img"&Ai(i)&"' type='hidden' value='"&iImg&"'>"
	 iLink = iLink&""&a(i)&""
	 iLink = iLink&"</div>"
	 sHtml = sHtml&vbcrlf&iLink
   Else
     iLink = vbcrlf&"<div id='Box"&Ai(i)&"' class='TLink'><input id='Url"&Ai(i)&"' type='radio' checked value='"&iUrl&"'>"
     iLink = iLink&"<a href='?ID="&ID&"&yAct=Detail&Url="&iUrl&"'>View</a>-"
	 iLink = iLink&"<span id='Del"&Ai(i)&"' class='hand' onClick=""reItem('"&Ai(i)&"')"">Del</span>."
	 iLink = iLink&"<span id='No"&Ai(i)&"' class='scomm'>"&(i+1)&"-"&iDo&" </span>"
	 iLink = iLink&"<input id='Img"&Ai(i)&"' type='hidden' value=''>"
	 iLink = iLink&""&a(i)&""
	 iLink = iLink&"</div>"
	 sHtml = sHtml&vbcrlf&iLink
   End If
   'sHtml = sHtml&vbcrlf&"<li><input id='' name='' type='radio' value=''>"&"<a href='?ID="&ID&"&yAct=Detail&Url="&iUrl&"' Xtarget='_blank'>View-"&(i+1)&"</a> : "&a(i)&"</li>" ' &Replace(sHtml,"||",vbcrlf&"<br>")
 Next
 
End If


%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:1px solid #CCC; margin-top:2px;">
  <tr>
    <td height="20" nowrap style="padding-left:30px;"><div style="float:right"> 
    <a href="?ID=<%=ID%>&yAct=CText">截取代码</a>
    <a href="?ID=<%=ID%>&yAct=CHtml">截取资料</a>
    <a href="?ID=<%=ID%>&yAct=PText">页面文本</a> 
    <a href="?ID=<%=ID%>&yAct=PHtml">页面代码</a>
    <a href="?ID=<%=ID%>&yAct=LText">连接代码</a> 
    <a href="?ID=<%=ID%>&yAct=">连接列表</a>
    <a href="?ID=<%=ID%>&yAct=LAll">所有连接</a>
    </div>
      <strong> <%=InfSubj%> 采集方案 </strong> </td>
    <td width="25%" rowspan="4" align="left" valign="top">
    <div style="width:260px; padding:5px; background-color:#F0F0F0"> 说明： <br>
      <div id="msgBox" style="padding:1px; background-color:#FFFFCC; clear:both;">
	    
        <div id="msgSend" style="float:left;"></div>
        <div id="msgEnd" style="float:left;"></div>
        &nbsp;
      </div>
      <!--#include file="imp_read.asp"-->
    </div></td>
  </tr>
  <tr>
    <td height="100%" valign="top" style="padding:8px;">
<div style="text-align:right">
        <input type="submit" name="btmSend" id="btmSend" value="确认提交" onClick="sdForms()" <%If KeyMod="Module" Then Response.Write("disabled")%>>
      </div>
<!-- Start -->
<%=sHtml%>
<!-- End -->
    </td>
  </tr>
</table>
<script type="text/javascript">

var Si = "<%=Si%>"; 
var Ai = Si.split("|");
var Ni=<%=nMin%>; 
var sendNO = 0;
var sendOK = 0;
var sendNull = 0;

getElmID("msgSend").innerHTML = " &nbsp; 提示：资料待确认提交... ";
function sdForms()
{ 
  getElmID("btmSend").disabled = true;
  if(Ni<=<%=nMax%>) { 
    try{ 
	  iBox = getElmID("Box"+Ai[Ni]);
	  getElmID("Del"+Ai[Ni]).innerHTML = " ";
	  getElmID("No"+Ai[Ni]).innerHTML = " ... ";
	  getElmID("msgSend").innerHTML = " &nbsp; 正在提交第["+[Ni+1]+"]个... ";
      var url = "<%=doFile%>?yAct=Insert";
	    url += "&no="+Ni;
		url += "&ID=<%=ID%>";
		url += "&ModID=<%=ModID%>";
		url += "&KeyCode=<%=codID%>"+Ai[Ni];
		url += "&Url="+getElmID("Url"+Ai[Ni]).value;
		url += "&id1=<%=prtID%>";
		url += "&id2="+Ai[Ni];
		url += "&Now=<%=Now()%>";
		url += "&Frm=<%=aPart(2)%>";
		url += "&Img="+getElmID("Img"+Ai[Ni]).value;
		//window.open(url);
      var iHttp = getXmlHttp();
	  iHttp .open("GET", url, true); 
      //xmlHttp .onreadystatechange = ChkAjUpd(iHttp);
	  iHttp.onreadystatechange = function(){
	    if (iHttp.readyState == 4) {
	      var rData = iHttp.responseText; 
	      var rMsg = "";
	      if(rData=="Y"){ rMsg = "<%=vMem_AP1B1%>";}
	      if(rData=="N"){ rMsg = "<%=vMem_AP1B2%>";}
	      var id6 = rData.substring(1,3); 
		  getElmID("Del"+id6).innerHTML = id6;
	      getElmID("No"+id6).innerHTML = rData.replace("("+id6+")","");
		  //getElmID("msgEnd").innerHTML = " 已处理完交第["+[Ni+1]+"]个;  ";
	    }
	  }
      iHttp .send(null);
	  sendOK++;
	  de = 500;
    }catch(err){  
      sendNull++;
	  de = 19;
    }
	Ni++;
	setTimeout("sdForms()",de); // setInterval
  }else{
	//alert(' OK ');
	getElmID("msgSend").innerHTML = " &nbsp; ["+sendOK+"] 个项目提交完毕: ["+sendNull+"] 个空资料忽略！";
  }
}
function reItem(id)
{ 
  e = getElmID("Box"+id);
  e.parentNode.removeChild(e); 
  //e.style.display='none';
  //e.style.visibility='hidden'; 
}
</script>
</body>
</html>
