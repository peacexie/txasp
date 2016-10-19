<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" src="../../inc/home/jsInfo.js"></script>
<mce:style type="text/css">
<!--  
#preview_wrapper{        
    display:inline-block;        
    width:100px;        
    height:48px;        
    background-color:#F0F0F0;        
}        
#preview_fake{ /* 该对象用户在IE下显示预览图片 */        
    filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);        
}        
#preview_size_fake{ /* 该对象只用来在IE下获得图片的原始尺寸，无其它用途 */        
    filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=image);          
    visibility:hidden;        
}        
#preview{ /* 该对象用户在FF下显示预览图片 */        
    width:100px;        
    height:48px;        
}        

-->
</mce:style>
<style type="text/css" mce_bogus="1">
#preview_wrapper {
	display:inline-block;
	width:100px;
	height:48px;
	background-color:#F0F0F0;
}
#preview_fake { /* 该对象用户在IE下显示预览图片 */
 filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);
}
#preview_size_fake { /* 该对象只用来在IE下获得图片的原始尺寸，无其它用途 */
 filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=image);
	visibility:hidden;
}
#preview { /* 该对象用户在FF下显示预览图片 */
	width:100px;
	height:48px;
}
</style>
<script type="text/javascript" src="bat_funcs.js"></script>
</head>
<body>
<%

no = Request("no")
id1 = Request("id1")
id2 = Request("id2")
ModName = Request("ModName") 
id = id1&id2

  Dim sys27_Rnd(10)
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next


tmpCont = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp"&ModID&"'") 
'Response.Write tmpCont
If tmpCont<>"" Then 
  tmpCont = Replace(tmpCont,vbcrlf," ")
  tmpCont = Replace(tmpCont,vbcr," ")
  tmpCont = Replace(tmpCont,vblf," ")
  
  tmpCont = Replace(tmpCont, chr(34), "\"&chr(34))
  tmpCont = Replace(tmpCont, "'", "\'") 
  tmpCont = Replace(tmpCont, "<", "\<")
  tmpCont = Replace(tmpCont, ">", "\>")

End If

%>
<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1'>
  <form name="iForm<%=Show_jsKey(id1)%>" id="iForm<%=Show_jsKey(id1)%>" action="../info/info_add2.asp?FrmFlag=<%=FrmFlag%>" enctype="multipart/form-data" method="post" target="_self">
    <tr>
      <td width='110' rowspan='2' align='center' xid='ImgShow'>
        <div id="preview_wrapper">
          <div id="preview_fake"> 
            <img id="preview" alt="图片预览" onload="onPreviewLoad(this)"/> 
          </div>
        </div></td>
      <td align='right'> 
      <div style="width:1px; height:1px; display:inline-block; float:left; overflow:hidden">
      <img id="preview_size_fake" />
      </div>
      (<%=no%>)</td>
      <td align='left'><input type='file' name='ImgName1' id='ImgName1' onChange="chkPic(this)" style='width:360px'>
      </td>
    </tr>
    <tr>
      <td width='30' align='right'>标题</td>
      <td align='left' nowrap><input type='text' name='InfSubj<%=sys27_Rnd(1)%>' onChange="setCont(this)" id='InfSubj<%=sys27_Rnd(1)%>' style='width:360px'>
        <a href='#' target='_self' onClick='remPic("_ID")'>移除</a></td>
    </tr>
    <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
    <input name="ModImgCount" type="hidden" id="ModImgCount" value="<%=ModImgCount%>">
    <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
    <input name="bakID" type="hidden" id="bakID" value="<%=id1%>">
    <input name="bakExt" type="hidden" id="bakExt" value="<%=id2%>">
    <input name="KeyCode" type="hidden" id="KeyCode" value="<%=Request("codID")&id2%>">
    <input name="InfCont<%=sys27_Rnd(4)%>" type="hidden" id="InfCont<%=sys27_Rnd(4)%>">
    <input name="InfType" type="hidden" id="InfType" value="">
    <input name="SetTop" type="hidden" id="SetTop" value="888">
    <input name="SetHot" type="hidden" id="SetHot" value="N">
    <input name="SetShow" type="hidden" id="SetShow" value="Y">
    <input name="SetSubj" type="hidden" id="SetSubj" value="000000">
    <input name="LogATime" type="hidden" id="LogATime" value="<%=DateAdd("s",no,Request("Now"))%>" size="20" maxlength="20">
  </form>
</table>

<script type="text/javascript">

var fm = iForm<%=Show_jsKey(id1)%>;
var fSubj = fm.InfSubj<%=sys27_Rnd(1)%>;
var fCont = fm.InfCont<%=sys27_Rnd(4)%>;
var fType = fm.InfType;
var dPar = parent.document;

function remPic(id)
{
  e = getElmID("ImgName1");
  e.outerHTML = e.outerHTML; 
  fSubj.value = "";
  
  pf = dPar.getElementById("iBox<%=id2%>");
  pf.parentNode.removeChild(pf);
  pf.innerHTML = "";
  pf.style.display='none';
  pf.style.visibility='hidden';
}

function chkPic(e)
{ 
  var v = e.value; 
  if(v.length>0){
	if(isIE){
	  var f1 = v.substring(1,3); 
	  if(f1!=":\\") { 
	    alert("附图地址错误<1>！"+e.value); 
	    e.outerHTML = e.outerHTML; 
	    return;
	  }
	  va = v.split("\\");
	  v = va[va.length-1];
	}
  }
  var ed=v.indexOf(".");
  var e4=v.substring(v.length-4,v.length);
  var e5=v.substring(v.length-5,v.length);
  if((ed<0)||((e4+e5).indexOf(".")<0)){
    alert("附图地址错误<2>！"); 
	e.outerHTML = e.outerHTML; 
	return;
  }
  fSubj.value = v.substring(0,ed);
  fType.value = dPar.getElementById("InfType").value;
  onUploadImgChange(e);
  //eImg = getElmID("ImgShow"); Url = e.value;
  //eImg.innerHTML = "<img src='"+Url+"' width='100' height='50' />";
  setCont(fSubj);
}

function setCont(e)
{ 
  tName = dPar.getElementById("InfType").options[dPar.getElementById("InfType").selectedIndex].text;
  tName = tName.replace("├","").replace("─","").replace("│","").replace(" ","") 
  fCont.value = e.value+" 的 详情(<%=ModName%>-"+tName+") <%=tmpCont%>";
}
</script>
</body>
</html>
