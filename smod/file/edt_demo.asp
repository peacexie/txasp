<!--#include file="../../upfile/sys/pcfg/editor.asp"-->
<!--#include file="../../smod/file/edt_config.asp"-->
<%
edTestBar = Request("edTestBar") 'edTestMod,edNameID
edTestPrt = Request("edTestPrt") 
edTestImg = Request("edTestImg")
If edTestPrt<>"" Then
  url = edcRootUrl&"/smod/file/edt_api.asp?edTestMod="&edNameID&"&edtAct="&edTestPrt&""
  If edTestPrt="mainLoad" Then
    jsLoad = "<"&flgScr&" src='"&url&"' "&fPara&"></"&flgScr&">"
	Response.Write "加载: <br>"&Show_Text(jsLoad)
	Response.End()
  Else
    url = url&"&edtID=EditID01&edtCont=InfCont&edtData=(boxHtml)&edtTool="&edTestBar
    Response.Redirect url
  End If
End If
InfConu = Request("InfConu")
If InfConu="" Then
  InfConu = ""
  InfConu = InfConu& "<blockquote> <h3 style='color:#FF00FF'>Peace目标：</h3></blockquote>"
  InfConu = InfConu& "  <li> 一个参数 切换各编辑器； </li>"
  InfConu = InfConu& "  <li> 一个API文件 控制所有编辑器； </li>"
  InfConu = InfConu& "  <li> 一个编辑器内核 应用于php,asp各语言； </li>"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%="("&edNameID&")"&edTempNM&" "%>--- Demos@ASP</title>
<script type="text/javascript" src="<%=edcRootUrl%>/inc/home/jsInfo.js"></script>
<script type="text/javascript" src="<%=edcRootUrl%>/inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtAct=mainLoad&edTestMod=<%=edNameID%>"></script>
<style type="text/css">
body, td, th {
	font-size: 12px;
	line-height:150%;
}
table.tbnote td {
	font-size: 12px;
	line-height:120%;
}
.dmform {
	clear:both;
	display:block;
	line-height:150%;
	border:1px solid #99F;
	padding:0px;
	margin:1px 0px 15px 0px;
}
.dmsend {
	clear:both;
	margin:3px;
}
.dmclear {
	clear:both;
}
.cmpPositive {
	color:#00F;
	text-align:center;
}
.cmpNegative {
	color:#F00;
	text-align:center;
}
.cmpLine {
	line-height:1px;
	text-align:center;
	color:#000;
	margin:3px;
}
div.cmpLine {
	padding:0px;
}
</style>
</head>
<body>
<div style="width:150px; position:absolute; right:10px; top:10px; border:1px solid #ccc; padding:10px 5px"> edNameID: <%=edNameID%>, <br />
  edNameNM: <%=edTempNM%>, <br />
  edToolbar: [<%=edTestBar%>], <br />
  edcRootUrl: <%=edcRootUrl%>, <br />
  edcEditUrl: <%=edcEditUrl%>, <br />
  Html5: <%=Request.ServerVariables("HTTP_CONTENT_DISPOSITION")%> </div>
<table width="600" border="1" class="dmform">
  <tr>
    <th>Editor(Mod)</th>
    <td><a 
    href="?edTestBar=<%=edTestBar%>&amp;edTestMod=edck3&amp;edTempNM=CKEditor&edTestImg=upload">CK3</a> - <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edkind&amp;edTempNM=KindEditor&edTestImg=imgFile">Kind</a> - <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edxh1&amp;edTempNM=xhEditor&edTestImg=filedata">xh1</a> - <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edue1&amp;edTempNM=UEditor&edTestImg=filename">ue1</a> - <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edtiny&amp;edTempNM=TinyMCE&edTestImg=(TinyMCE)">TinyMCE</a> - <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edfck&amp;edTempNM=FCKEditor&edTestImg=NewFile">FCK</a> # <a 
    href="?">(Defalut)</a> # <a 
    href="?edTestBar=<%=edTestBar%>&edTestMod=edeweb&amp;edTempNM=eWebEditor&edTestImg=(eWeb)">eWeb</a></td>
  </tr>
  <tr>
    <th>Toolbar(Bar)</th>
    <td><a 
    href="?edTestMod=<%=edNameID%>&amp;edTestBar=Full">Full</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestBar=Peace">Peace</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestBar=Basic">Basic</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestBar=vs">(vs)</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestBar=(API)&edTestImg=<%=edTestImg%>">Upload&amp;Notes</a></td>
  </tr>
  <tr>
    <th>Portion(Prt)</th>
    <td><a 
    href="?edTestMod=<%=edNameID%>&amp;edTestPrt=mainLoad">mainLoad</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestPrt=mainShow&amp;edTestBar=Full">(Full)</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestPrt=mainShow&amp;edTestBar=Peace">(Peace)</a> - <a 
    href="?edTestMod=<%=edNameID%>&amp;edTestPrt=mainShow&amp;edTestBar=Basic">(Basic)</a></td>
  </tr>
</table>
<div class="cmpLine" style="width:450px;background-color:#000;">&nbsp;</div>
<div class="cmpLine" style="width:500px;background-color:#F0F;">&nbsp;</div>
<div class="cmpLine" style="width:550px;background-color:#F00">&nbsp;</div>
<div class="cmpLine" style="width:600px;background-color:#00F;">&nbsp;</div>
<div class="cmpLine" style="width:650px;background-color:#000;">&nbsp;</div>
<div style="line-height:100%; margin:1px; padding-left:438px">450 - 500 - 550 - 600 - 650 (px)</div>
<%If edTestBar="vs" Then%>
<%
  h1 = 30
  If edNameID="edkind" Then  h1 = h1+50
  If edNameID="edfck" Then  h1 = h1+80
  h2 = 150
  If edNameID="edkind" Then  h2 = h2+50
  If edNameID="edfck" Then  h2 = h2+80
%>
<form id="frmDemo3" method="post" action="?" class="dmform">
  <textarea id="InfConv" name="InfConv" style="width:100%;height:<%=h1%>px;visibility:hidden;display:none"></textarea>
  <script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtID=EditID03&edtCont=InfConv&edtTool=Full&edTestMod=<%=edNameID%>"></script>
  <div class="dmclear"></div>
  <input type="submit" name="save3" value="Submit3" class="dmsend" />
  toolBar=Full(Defalut)
</form>
<form id="frmDemo1" method="post" action="?" class="dmform">
  <textarea id="InfCont" name="InfCont" style="width:600px;height:<%=h2%>px;visibility:hidden;display:none"></textarea>
  <script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtID=EditID01&edtCont=InfCont&edTestMod=<%=edNameID%>"></script>
  <div class="dmclear"></div>
  <input type="submit" name="save1" value="Submit1" class="dmsend" />
  toolBar=Peace(Null)
</form>
<%End If%>
<%If edTestBar="Full" Then%>
<form id="frmDemo3" method="post" action="?" class="dmform">
  <textarea id="InfConv" name="InfConv" style="width:100%;height:320px;visibility:hidden;display:none"></textarea>
  <script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtID=EditID03&edtCont=InfConv&edtTool=Full&edTestMod=<%=edNameID%>"></script>
  <div class="dmclear"></div>
  <input type="submit" name="save3" value="Submit3" class="dmsend" />
  toolBar=Full(Defalut)
</form>
<%End If%>
<%If edTestBar="Peace" Or edTestBar="" Then%>
<form id="frmDemo1" method="post" action="?" class="dmform">
  <textarea id="InfCont" name="InfCont" style="width:580px;height:320px;visibility:hidden;display:none"></textarea>
  <script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtID=EditID01&edtCont=InfCont&edtData=boxHtml&edTestMod=<%=edNameID%>"></script>
  <div class="dmclear"></div>
  <input type="submit" name="save1" value="Submit1" class="dmsend" />
  toolBar=Peace(Null)
</form>
<%End If%>
<%If edTestBar="Basic" Then%>
<form id="frmDemo2" method="post" action="?" class="dmform">
  <textarea id="InfConu" name="InfConu" style="width:450px;height:320px;visibility:hidden;display:none"><%=InfConu%></textarea>
  <script type="text/javascript" charset="utf-8" src="<%=edcRootUrl%>/smod/file/edt_api.asp?edtID=EditID02&edtCont=InfConu&edtTool=Basic&edTestMod=<%=edNameID%>"></script>
  <div class="dmclear"></div>
  <input type="submit" name="save2" value="Submit2" class="dmsend" />
  toolBar=Base
</form>
<%End If%>
<%If edTestBar="(API)" Then%>
<%
If edTBakID=edNameID Then
  edUpMsg = "文件提交["&edTestImg&"]"
  edUpDis = ""  
Else
  edUpMsg = "禁止提交["&edTestImg&"]"
  edUpDis = "disabled='disabled'"  
End If
%>
<form id="frmDemo4" method="post" action="<%=edcRootUrl%>/smod/file/edt_up.asp" class="dmform" enctype="multipart/form-data">
  <table width="100%" border="0">
    <tr>
      <td nowrap="nowrap">上传</td>
      <td><input type="file" name="<%=edTestImg%>" id="<%=edTestImg%>" style="width:380px" /></td>
      <td nowrap="nowrap"><input type="submit" name="button2" <%=edUpDis%> value="<%=edUpMsg%>" style="padding:0px 5px; margin:0px 5px" />
        <input name="edTestMod" type="hidden" id="edTestMod" value="<%=edTestMod%>" /></td>
    </tr>
  </table>
</form>
<table width="600" border="1" class="dmform tbnote">
  <tr>
    <th nowrap="nowrap">基本用法-<%=edNameID%></th>
    <th nowrap="nowrap">基本API</th>
    <th nowrap="nowrap">高级API</th>
  </tr>
  <tr>
    <td align="left" valign="top">apiGetValEditID01()<br />
      apiSetValEditID01(Val)<br />
      apiSetVIDEditID01(id)<br />
      apiInsertEditID01(Val)</td>
    <td colspan="2" align="left" valign="top">注意：$('#InfCont') = document.getElementById('InfCont')</td>
  </tr>
  <%If edNameID="edck3" Then%>
  <tr>
    <td align="left" valign="top" nowrap="nowrap">CKEDITOR.replace('InfCont',{<br />
      skin : 'office2003',<br />
      toolbar : 'toolBar_Basic',<br />
      width : $('#InfCont').style.width,<br />
      height : $('#InfCont').style.height,<br />
      language : 'zh-cn'<br />
      });</td>
    <td align="left" valign="top">var edObj = CKEDITOR.instances.EditID01;<br />
      <br />
      edObj.getData(); <br />
      edObj.setData(strHrml); <br />
      edObj.insertHtml(strHrml);<br /></td>
    <td align="left" valign="top">edObj.focus(); <br />
      //获取焦点</td>
  </tr>
  <%End If%>
  <%If edNameID="edkind" Then%>
  <tr>
    <td align="left" valign="top"> var edObj;<br />
      KindEditor.ready(function(K) {<br />
      edObj = K.create('#InfCont');<br />
      });</td>
    <td align="left" valign="top">edObj.html();<br />
      edObj.html(strHrml);<br />
      edObj.insertHtml(strHrml);</td>
    <td align="left" valign="top">edObj.remove()<br />
      edObj.exec('About')</td>
  </tr>
  <%End If%>
  <%If edNameID="edxh1" Then%>
  <tr>
    <td align="left" valign="top"> class=&quot;xheditor {skin:'default'}&quot;<br />
      $('#InfCont').xheditor(); 或者 <br />
      $('#InfCont').xheditor({tools:'mini'});<br />
      var edObj=$('#InfCont'). xheditor({skin:'default'});</td>
    <td align="left" valign="top">edObj.getSource();<br />
      edObj.setSource(strHrml);<br />
      edObj.pasteHTML(strHrml);</td>
    <td align="left" valign="top">edObj.focus(); </td>
  </tr>
  <%End If%>
  <%If edNameID="edue1" Then%>
  <tr>
    <td align="left" valign="top">var edObj = new baidu.editor.Editor({<br />
      //id: &quot;eBox_$edtID&quot;,<br />
      textarea : 'InfCont',<br />
      initialContent: strHrml,<br />
      minFrameHeight: 320,<br />
      autoHeightEnabled: true<br />
      }<br /></td>
    <td align="left" valign="top">edObj.getContent();<br />
      edObj.setContent(strHrml);<br />
      edObj.pasteContent(strHrml);</td>
    <td align="left" valign="top">&lt;div id='eBox_$edtID'&gt;&lt;/div&gt;</td>
  </tr>
  <%End If%>
  
  <%If edNameID="edtiny" Then%>
  <tr>
    <td align="left" valign="top">var strEditID = 'EditID01'; <br />
      var objContID = document.getElementById('InfCont');<br />
objContID.style.display = '';<br />
objContID.style.visibility = 'visible';<br />
tinyMCE.init({<br />
mode : 'textareas',<br />
plugins : toolBar_Plugs2,<br />
theme : 'advanced',<br />
theme_advanced_buttons1 : toolBar_Peace1,<br />
theme_advanced_buttons2 : toolBar_Peace2,<br />
theme_advanced_buttons3 : toolBar_Peace3,<br />
theme_advanced_buttons4 : toolBar_Peace4,<br />
content_css : '/asp/sadm/edtiny/themes/content.css',<br />
theme_advanced_toolbar_location : 'top',<br />
theme_advanced_toolbar_align : 'left',<br />
theme_advanced_statusbar_location : 'bottom',<br />
theme_advanced_resizing : true<br />
});<br /></td>
    <td colspan="2" align="left" valign="top">tinyMCE.get('$edtCont').getContent();<br />
      tinyMCE.get('$edtCont').setContent(Val);<br />
      tinyMCE.execCommand('mceInsertContent',false,Val);<br />
      <br />
      var objContID = document.getElementById('$edtCont');<br />
      objContID.style.display = '';<br />
    objContID.style.visibility = 'visible';</td>
  </tr>
  <%End If%>
  
  <%If edNameID="edfck" Then%>
  <tr>
    <td align="left" valign="top" nowrap="nowrap">var strEditID = 'EditID01' <br />
      var edObj = new FCKeditor('EditID01');<br />
      edObj.BasePath	= '/asp/sadm/edfck/'; <br />
      edObj.Width	= $('#InfCont').style.width;<br />
      edObj.Height	= $('#InfCont').style.heigh;<br />
      edObj.Value	= $('#InfCont').value;<br />
      edObj.ToolbarSet = 'toolBar_Basic';<br />
      edObj.Create() ;<br /></td>
    <td align="left" valign="top">var edObj = FCKeditorAPI.GetInstance ('EditID01');<br />
      <br />
      edObj.GetXHTML(true)<br />
      edObj.SetHTML(strHrml)<br />
      edObj.InsertHtml(strHrml)</td>
    <td align="left" valign="top">edObj.EditorDocument .body.innerText<br />
      opener.FCKeditorAPI .GetInstance('InstanceName');</td>
  </tr>
  <%End If%>
</table>
<table width="600" border="1" class="dmform">
  <tr>
    <th>编辑器</th>
    <th>产地</th>
    <th>稳定</th>
    <th>轻量</th>
    <th>支持</th>
    <th>主要优点</th>
    <th>主要不足</th>
    <th>Ver</th>
    <th>速</th>
    <th>肥</th>
  </tr>
  <tr>
    <td><a href="http://ckeditor.com/download" target="_blank">CKEditor</a></td>
    <td>国外老牌</td>
    <td class="cmpPositive">稳定</td>
    <td class="cmpNegative">否</td>
    <td class="cmpPositive">团队</td>
    <td>功能强大,稳定</td>
    <td>臃肿,加载慢</td>
    <td class="cmpPositive">3.6</td>
    <td class="cmpNegative">4</td>
    <td class="cmpNegative">0.90</td>
  </tr>
  <tr>
    <td><a href="http://www.kindsoft.net/down.php" target="_blank">KindEditor</a></td>
    <td>国产(上海-浩跃软件)</td>
    <td>&nbsp;</td>
    <td class="cmpPositive">轻量</td>
    <td>&nbsp;</td>
    <td>插件扩展</td>
    <td>&nbsp;</td>
    <td class="cmpPositive">4.0</td>
    <td class="cmpPositive">2</td>
    <td class="cmpPositive">0.24</td>
  </tr>
  <tr>
    <td><a href="http://xheditor.com/download" target="_blank">xhEditor</a></td>
    <td>国产(台州-[王一])</td>
    <td class="cmpNegative">差</td>
    <td class="cmpPositive">轻量</td>
    <td class="cmpNegative">个人</td>
    <td>迷你高效,插件扩展</td>
    <td><span class="cmpNegative">表格编辑</span>,不稳定</td>
    <td class="cmpNegative">1.1</td>
    <td class="cmpPositive">1</td>
    <td class="cmpPositive">0.49</td>
  </tr>
  <tr>
    <td><a href="http://ueditor.baidu.com/index.html" target="_blank">UEditor</a></td>
    <td>国产(百度)</td>
    <td>&nbsp;</td>
    <td class="cmpPositive">轻量</td>
    <td class="cmpPositive">百度</td>
    <td>小巧,分层架构</td>
    <td>&nbsp;</td>
    <td class="cmpNegative">1.1</td>
    <td class="cmpPositive">3</td>
    <td class="cmpPositive">0.44</td>
  </tr>
  <tr>
    <td><a href="http://www.tinymce.com/download/download.php" target="_blank">TinyMCE</a></td>
    <td>国外老牌</td>
    <td>&nbsp;</td>
    <td class="cmpPositive">轻量</td>
    <td>&nbsp;</td>
    <td>素雅清新,轻量级</td>
    <td>&nbsp;</td>
    <td class="cmpPositive">3.4</td>
    <td class="cmpNegative">6</td>
    <td class="cmpNegative">0.99</td>
  </tr>
  <tr>
    <td>FCKEditor</td>
    <td>国外老牌[已经退役]</td>
    <td align="center">---</td>
    <td align="center">---</td>
    <td align="center">---</td>
    <td>-----</td>
    <td>-----</td>
    <td align="center">2.6</td>
    <td class="cmpNegative">4</td>
    <td class="cmpNegative">1.11</td>
  </tr>
  <tr>
    <td><a href="http://www.ewebeditor.net/download.asp" target="_blank">eWebEditor</a></td>
    <td>国产(福州-极限软件)</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>功能齐全强大</td>
    <td><span class="cmpNegative">收费</span>,<span class="cmpNegative">要插件</span></td>
    <td class="cmpPositive">7.3</td>
    <td align="center">-</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="10" class="tbnote">&nbsp;* 肥:指体积大小,单位是MB</td>
  </tr>
</table>
<table width="600" border="1" class="dmform">
  <tr>
    <td align="left"><span class="cmpPositive">Peace目标</span>：<br />
      &nbsp;* 一个参数 切换各编辑器(CK3 - Kind - xh1 - eWeb - U1 - TinyMCE - FCK 等)；<br />
      &nbsp;* 一个API文件 控制所有编辑器，增加一个编辑器，只改一个文件，其它文件不变；<br />
      &nbsp;* 一个编辑器内核 应用于php,asp各语言；</td>
  </tr>
  <tr>
    <td align="left"><span class="cmpPositive">Editor背景</span>：<br />
      &nbsp;* IE6终奖退役，IE9及其它新浏览器必将来临，<br />
      &nbsp;* 目前我用的FCK2.6在官网已经退役，换编辑器势在必行！</td>
  </tr>
</table>
<%End If%>
<div id="boxHtml" style="width:560px; height:auto; border:1px solid #CCC; padding:5px; display:none">
  <p><strong>Editor(编辑器) - 基本属性使用说明</strong></p>
  <p><img src="<%=edcRootUrl%>/img/logo/y_logo.gif" alt="" width="49" height="42" hspace="5" vspace="0" align="left" /><strong>*** 两种换行：</strong><br />
    后台编辑框中的换行有两种(其实是网页的两种换行规则)：段落换行(直接按回车)<br />
    和Shift+回车换行；默认情况，前者换行间空隙比较大，后者换行间空隙比较小；根据需要设置和摸索。(本段一个左对齐测试图片。) </p>
  <p><strong>*** 从Word和网页中复制资料：<br />
    </strong>强烈建议不要直接从Word和网页中直接复制文本/图片到编辑框；建议使用“粘贴文本”命令或先把文本复制到记事本,把原有的CSS去掉后,再从记事本复制粘贴到编辑框后再编辑；如果一定要从word中复制文字，请找到"从word粘贴"按钮，之后可能出现如右图提示，点确定，然后按提示操作既可(此用于FCKEditor编辑器，一般的，较新版本出现此提示) ；</p>
  <p align="center"><img src="<%=edcRootUrl%>/tools/himg/fck-msg.JPG" alt="Copy Text" width="477" height="114" /></p>
  <p>这种方式可以 把word中的表格，其他网页中的图片一起粘贴的网页；如果粘贴的表格不满意，可以看以下"插入表格" 的说明；粘贴的图片，其实是一个地址，建议删除图片，重新上传一次；比如，Word文件资料的设置可能是要适合A4,A3纸张打印，而你的网页资料宽度可能要适合你前台网页显示的宽度，所以复制文字资料后，做一些编辑是非常必要的！从网络上贴来的图片要谨用, 一旦原网站的图片删除后,将不能正常显示该图片，原网站作了"防盗接"等技术后，也不能正常显示。 </p>
  <p><strong>*** 建议用默认字体：</strong><br />
    找到字体 下拉框，可以设置不同字体；但请注意，建议您使用默认字体，不用自己设置字体，特别是 隶书,幼圆等字体，因为显示字体取决于用户电脑 现有的字体；如果您设置了 幼圆等字体，在你电脑上显示非常漂亮，是因为你电脑安装了幼圆字体，在没有安装幼圆字体的电脑上，就显示默认的（可能是宋体）字体了!有需要的话，可设置文字属性：<br />
    <strong>加粗</strong>，<em>斜体</em>，<em>下划线</em>，<span style="color: #ff0000">文字红色</span>，<span style="background-color: #c0c0c0">灰色背景色</span>，等属性，可以把这些属性叠加。</p>
  <p><strong>*** 插入图片和图片属性：</strong><strong><img src="<%=edcRootUrl%>/upfile/myfile/xadv/temp_2008.jpg" alt="" width="180" height="140" hspace="5" vspace="0" border="1" align="right" /></strong><br />
    插入图片：a. 找到"插入(编辑)图片"按钮；<br />
    b. 上传(从本地)或浏览(网站已有的)一个图片；<br />
    c. 图片对齐：选中以上的图片（点击）；<br />
    d. 按右键，图片属性，设置它们各自的对齐属性；<br />
    另外设置间距，边框等属性看看效果……；<br />
    如果图片比较小，可以设置左右对齐，如果图片比较大，设置中间对齐比较合适，一切根据你的版面和内容设置，注意，设置中间对齐,建议图片在独立段落比较好些。(本段一个右对齐测试图片。)</p>
  <p><strong>*** 各种链接<br />
    </strong>1. 文字链接：<a href="http://www.dg.gd.cn/" target="_blank">myLink(我的链接)</a> - 打开新窗口 - 链接到dg.gd.cn<br />
    a. (输入文字) - myLink(我的链接) <br />
    b. (选中文字)<br />
    c. (找到插入(编辑)链接图标)<br />
    d. 地址填：http://www.dg.gd.cn/<br />
    e. 目标(选项卡)：(选) 空窗口(blank)</p>
  <p>2. 打开QQ聊天窗口：点我的QQ号码，打开QQ聊天窗口，与我聊天，这也是一个链接！<br />
    同上，但地址不同：tencent://message/?uin=[QQ号码]&amp;Site=[网站名称]&amp;Menu=yes 如：<br />
    &nbsp; -=&gt; 技术探讨QQ: <a href="tencent://message/?uin=80893510&amp;Site=[PeaceXieys]&amp;Menu=yes" target="_blank">80893510</a>。</p>
  <p>3. 图片链接：a. (插入一个图片，步骤见上，或) <strong><a href="http://www.dgchr.com/" target="_blank"><img src="<%=edcRootUrl%>/upfile/myfile/logo/logo-2009.jpg" alt="" width="120" height="60" hspace="5" vspace="0" border="1" align="right" /></a></strong><br />
    b. (插入一个表情图,第一行，倒数第五个按钮)<br />
    c. (选中图片,右键选图片属性)<br />
    d. 链接(选项卡)：地址：http://www.dgchr.com/<br />
    e. 目标：目标(选)新窗口 logo-2009.jpg</p>
  <p>4. 点击下载等的链接：a. (输入文字信息)<br />
    b. (选中文字)<br />
    c. (找到插入(编辑)链接图标) <br />
    d. 
    地址填：一个要下载的目标文件地址，这个地址可以是：<br />
    从本地上传 或 浏览网站已有的 一个文件 或 输入其它网站的一个文件地址；<br />
    &nbsp; -=&gt; 如：<a href="<%=edcRootUrl%>/upfile/myftp/down/ZoomImg.rar">下载压缩文件(批量图片处理工具)<br />
    </a>注意，大文件不适合程序直接上传！如果链接的是Word,Excel等Office文件，点击后是下载还是直接打开，取决与客户端安装的软件和浏览器设置。</p>
  <p><strong>*** 添加视频     步骤说明</strong><br />
    a. (找到插入(编辑)视频图标)<br />
    b. 从本地上传 或 浏览网站已有的 一个视频文件 或 输入其它网站的一个视频地址；<br />
    c. 设置必要的参数...<br />
    注意，大文件(视频)不适合程序直接上传！添加Flash文件类似；<br />
    插入 视频文件 | Flv文件 | Flash文件 等媒体文件示例，请点本文最上面标题下的相关连接……</p>
  <p><strong>*** 添加表格</strong>
  <table cellspacing="1" cellpadding="1" width="200" border="1" align="right">
    <tbody>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </tbody>
  </table>
  如果添加表格，同样可设置表格，单元格属性。<br />
  把光标移动到单元格，按右键，分别选行，列，表格属性，设置各种值测试。看看效果……
  </p>
</div>
</body>
</html>