<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<%

IR = RequestS("IR",3,48) 
yAct = Request("yAct") ': Response.Write yAct
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)
ModSubj = rs_Val("","SELECT InfSubj FROM "&ModTab&" WHERE KeyID='"&IR&"'")


cID = 0
sID = ""

If yAct="SetTop" Then

  yVal = RequestS("yVal","N",888)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sql = " UPDATE InfoPhoto SET "&yAct&"="&yVal&" WHERE KeyID='"&iID&"'" 
	  Call rs_DoSql(conn,sql)
	End If
  Next
  Msg = " 设置成功!"

Elseif yAct="Edit" then

   i = RequestS("i","N",0)
   t = RequestS("i"&i&"Top","N",888)
   c = RequestS("i"&i&"Cont","C",512)
   s = RequestS("i"&i&"Subj","C",512)
   k = RequestS("i"&i&"ID","C",48)
   sql = "UPDATE InfoPhoto SET InfSubj='"&s&"',InfCont='"&c&"',SetTop="&t&" WHERE KeyID='"&k&"' "
   'Response.Write sql
   Call rs_DoSql(conn,sql)
   Msg = " 修改成功!"

Elseif yAct="del_sel" then

  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  iImg = rs_Val("","SELECT ImgName FROM InfoPhoto WHERE KeyID='"&iID&"'")
	  If iImg&""<>"" Then
	    'Response.Write iImg
		Call fil_del(Config_Path&"upfile/"&iImg)
	  End If
	  sql = " DELETE FROM InfoPhoto WHERE KeyID='"&iID&"'" 
	  Call rs_DoSql(conn,sql)
	End If
  Next

	'Call rs_DFile("InfoPhoto",sID,"")
    Msg = " 删除成功!"
	
Elseif yAct="delClear" then
	'Call rs_DFile("InfoPhoto",sID,"")
    'Msg = cID&" 条记录删除成功!"
	
End If
    sql = " SELECT * FROM [InfoPhoto] "
	sql =sql& " WHERE KeyRe='"&IR&"' "
	sql =sql& " ORDER BY SetTop,LogATime DESC " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center">
    <td height="27" colspan="7" align="center" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center">
          <td width="30%" align="center"><strong>相关图片 | <a href="../info/info_list.asp?PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&TP2=<%=TP2%>&PrmFlag=<%=PrmFlag%>">返回</a> | <a href="?IR=<%=IR%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>">刷新</a></strong></td>
          <td width="30%" align="center">&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <td align="right" nowrap>原文:<a href="../info/info_view.asp?ID=<%=IR%>" target="_blank"><%=ModSubj%></a></td>
        </tr>
      </table></td>
  </tr>

  <tr align="center" bgcolor="#F0F0F0">
    <td width="5%" align="center" nowrap>NO</td>
    <td align="center">主题</td>
    <td width="15%" align="center" nowrap>图片</td>
    <td width="12%" align="center" nowrap>内容</td>
    <td width="12%" align="center" nowrap>顺序</td>
    <td width="8%" align="center" nowrap>发布</td>
    <td width="8%" align="center" nowrap>修改</td>
  </tr>

  <form name="flist" id="flist" method="post" action="?">
    <%
  If not rs.eof then
  i = 0
  Do While NOT rs.EOF
  i = i + 1
  col = "#FFFFFF"
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = rs("InfSubj")
InfCont = Show_Form(rs("InfCont"))
SetTop = rs("SetTop")
ImgName = rs("ImgName") ':Response.Write ImgName
LogATime = rs("LogATime")
If ImgName<>"" Then
  ImgName = "<img src='"&Replace(Config_Path&ImgName,"//","/")&"' border=0 align='absmiddle'width='80' height='20' border=0 onload='javascript:setImgSize(this);'>" 
Else
  ImgName = "<font color=red>[无]</font>"
End If
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" valign="top" nowrap bgcolor="#FFFFFF"><%=i%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td>
        <input name="i<%=i%>ID" type="hidden" id="i<%=i%>ID" value="<%=KeyID%>" />
        <input name="i<%=i%>Subj" type="text" id="i<%=i%>Subj" style="width:150px" value="<%=InfSubj%>" size="12" maxlength="120"></td>
      <td align="center" nowrap><IFRAME marginWidth=0 marginHeight=0
          src="../file/img_set.asp?TabID=InfoPhoto&upPath=<%=IR%>&ID=<%=KeyID%>&NO=&WW=200"
          width="240" height='23' frameBorder=0 scrolling=no> </IFRAME></td>
      <td align="center" nowrap><input name="i<%=i%>Cont" type="text" id="i<%=i%>Cont" style="width:80px" value="<%=InfCont%>" size="12" maxlength="120"></td>
      <td align="center" nowrap><input name="i<%=i%>Top" type="text" id="i<%=i%>Top" value="<%=SetTop%>" size="4" maxlength="4" style="width:40px"></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap>
        <input type="button" name="Submit2" value="修改" onClick="ESend(<%=i%>)">
      </td>
    </tr>
    <%
  rs.Movenext
  Loop
  
%>

    <tr bgcolor="#F0F0F0">
      <td height="21" align="right" nowrap><span id="yFlag" style="visibility:hidden ">N</span>
        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td align="left">全选
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="i" type="hidden" id="i" value="0" />
        <input name="IR" type="hidden" id="IR" value="<%=IR%>"></td>
      <td align="right" nowrap>&nbsp;</td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
        <option value="del_sel">删除.所选</option>
        <option value="Edit">设置/修改</option>
        </select>        <select name="yVal" id="yVal" style="width:40px">
          <option value="Y">Y</option>
          <option value="N">N</option>
          <option value="X">X</option>
          <%For j=0 To 9%>
          <option value="<%=j%>" ><%=j%></option>
          <%Next%>
          <option value="888">888</option>
      </select></td>
      <td align="center" nowrap><input type="submit" name="Submit" value="执行"></td>
    </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="7" align="center"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TP2="&TP2&"",1)%></td>
  </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="9">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
  </form>

</table>
<script type="text/javascript">

var fmObj = document.flist;

function ESend(n)
{
  fmObj.yAct.value = "Edit";
  fmObj.i.value = n;
  fmObj.submit();
  //alert('a');
}

function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>

<%
prtID = Left(rs_AutID(conn,ModTab,"KeyID",upPart,"1",""),22)
codID = Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",4)
Dim Si,Ui :Ui="02" 
For i=1 To 240 ',96 
  Ui = Next_ID(Ui,"02",3)
  Si = Si&Ui&"|" 
Next
%>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:1px solid #CCC; margin-top:2px;">

  <tr>
    <td id="batPics"></td>
    <td width="32%" align="left" valign="top" style="padding:5px; background-color:#F0F0F0">说明：<br>
      <span class="col00F">***1.</span> 建议<span class="colF0F">用标题作为图片名</span>。
      <div id="msgBox" style="padding:1px; background-color:#FFFFCC"></div></td>
  </tr>
  <tr>
    <td><div style="float:right">
        <input type="submit" name="btmSend" id="btmSend" value="确认增加" onClick="sendForms()">
      </div>
      <input type="button" name="btnA5" id="btnA5" value="+16" onClick="insBox(16)">
      <input type="button" name="btnA4" id="btnA4" value="+8" onClick="insBox(8)">
      <input type="button" name="btnA3" id="btnA3" value="+4" onClick="insBox(4)">
      <input type="button" name="btnA2" id="btnA2" value="+2" onClick="insBox(2)">
      <input type="button" name="btnA1" id="btnA1" value="+1" onClick="insBox(1)">
      <a href="?IR=<%=IR%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>">重载本页</a></td>
    <td width="32%" align="left" valign="top" style="padding:5px; background-color:#F0F0F0">&nbsp;</td>
  </tr>
</table>
<script type="text/javascript">

var Si = "<%=Si%>"; 
var Ai = Si.split("|"); 
var Ni=0; 

var sendNO = 0;
var sendOK = 0;
var sendNull = 0;
function sendForms()
{ 
  getElmID("btmSend").disabled = true;
  for(i=1;i<=5;i++){ getElmID("btnA"+i).disabled = true; }
  sendNO++; i = sendNO; 
  if(sendNO<=Ni) { 
	de = 13; if(chkBox(Ai[i])) { de=300; }
	setTimeout("sendForms()",de); // setInterval
  }else{
	sendNG = Ni-sendOK-sendNull;
	sendMsg = " 共 ["+sendOK+"] 个图片上传完毕！";
	if(sendNull>0) {sendMsg +=" <br>["+sendNull+"] 个项目被忽略(移除); ";}
	if(sendNG>0) {sendMsg +=" <br>["+sendNG+"] 个空项目未提交! ";}
	getElmID("msgBox").innerHTML = "<span class='col00F'>***5.</span>"+sendMsg;
	//alert(sendMsg);
  }
}
function chkBox(id)
{ 
  try{ 
    var sDoc = window.document.getElementById("iFrame"+id).contentWindow.document;
	var sFrm = sDoc.getElementById("iForm<%=Show_jsKey(prtID)%>"); 
	var sImg = sDoc.getElementById("ImgName1"); 
	//alert(sImg.value);
	if(sImg.value.length==0) {  getElmID("iBox"+id).innerHTML = " <div style='padding:3px 0px 0px 24px'><span class='colF0F'>空项目，未提交！</span></div>"; }
	else { sFrm.submit(); sendOK++; }
	return true;
  }catch(err){  
    sendNull++;
	return false;
  }
  
}


tmp = "";
tmp += "<div id='iBox_ID' class='iBox'>";
tmp += "<IFRAME id='iFrame_ID' src='rpic_form.asp(Paras)' frameBorder=0 width='560' scrolling='no' height='50'></IFRAME>";
tmp += "</div>";
function insBox(n)
{
  var fmsObj = getElmID("batPics");
  for(i=1;i<=n;i++)	
  {
	j = i+Ni; if(j>24){ alert("最多24个"); return; }
	var p = "?no="+j;
	p += "&id1="+'<%=prtID%>';
	p += "&id2="+Ai[j];
	p += "&ModID=<%=ModID%>";
	p += "&codID=<%=codID%>";
	p += "&IR=<%=IR%>";
	p += "&Now=<%=Now()%>";
	ti = tmp.replace("(Paras)",p);
	ti = ti.replace(/_ID/g,Ai[j],"gi");
	var ep = getElmID("batPics");
	eBox = document.createElement("div");
	eBox.id = Ai[j];
	eBox.innerHTML = ti;
	fmsObj.appendChild(eBox); 
  }
  Ni=Ni+n; //alert(sn);
}
insBox(1);


</script>

</body>
</html>
