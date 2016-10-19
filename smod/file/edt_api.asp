<!--#include file="../../upfile/sys/pcfg/editor.asp"-->
<!--#include file="edt_config.asp"-->
<% 
'//../../upfile/sys/pcfg/editor.asp
'//echo $edcEditID.'<br>';
'//echo $edcEditUrl.'<br>';

edtAct = Get_Para("edtAct","mainShow")
'jsConfig,jsPerm
'mainShow,mainLoad,(mainEnd)

If edtAct="jsConfig" Then ' // for fck fck_image.html
	
	Response.Write "var edrImg_Read, edrImg_Max, edrImgWidth, edrImgHeight;"&vbcrlf
	Response.Write "edrImg_Max = '"&edrImg_Max&"'"&vbcrlf
	Response.Write "edrImgWidth = '"&edrImgWidth&"'"&vbcrlf
	Response.Write "edrImgHeight = '"&edrImgHeight&"'"&vbcrlf
	Response.End
	
ElseIf edtAct="jsPerm" Then ' // for fck fck_dialog_common.js
	
    perm = Chk_PEditor("") '//Chk_Perm9("FileEditor","3");
    '// .... 
    Response.End
 
ElseIf edtAct="mainLoad" Then 

	' js 公共变量
	edcStr = "var edcEditID='"&edNameID&"',edcBasePath='"&edcRootUrl&"',edcFileExt='"&edcFileExt&"';"
	Response.Write ""&edcStr&""&vbcrlf 

	If edcEditID="edck3" Then
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/ckeditor.js' "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	ElseIf edcEditID="edkind" Then	   
	   Response.Write "document.write(""<link href='"&edcEditUrl&"/themes/default/default.css' rel='stylesheet' />"");"&vbcrlf 	
	   Response.Write "document.write(""<link href='"&edcEditUrl&"/plugins/code/prettify.css'  rel='stylesheet' />"");"&vbcrlf 
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/kindeditor.js'            "&flgPara&"></"&flgScr&">"");"&vbcrlf 	
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/lang/zh_CN.js'            "&flgPara&"></"&flgScr&">"");"&vbcrlf 	
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/plugins/code/prettify.js' "&flgPara&"></"&flgScr&">"");"&vbcrlf 	
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/_config.js'               "&flgPara&"></"&flgScr&">"");"&vbcrlf 		
	ElseIf edcEditID="edxh1" Then
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/xheditor.js'              "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/xheditor_peace/config.js' "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	   Response.Write "document.write(""<link href='"&edcEditUrl&"/xheditor_peace/style.css'      rel='stylesheet' />"");"&vbcrlf 
	ElseIf edcEditID="edue1" Then
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/editor_config.js'      "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/editor_ui_all.js'      "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	   Response.Write "document.write(""<link href='"&edcEditUrl&"/themes/default/ueditor.css' rel='stylesheet' />"");"&vbcrlf 
	ElseIf edcEditID="edtiny" Then
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/tiny_config.js' "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/tiny_mce.js'    "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	ElseIf edcEditID="edfck" Then
	   Response.Write "document.write(""<"&flgScr&" src='"&edcEditUrl&"/fckeditor.js' "&flgPara&"></"&flgScr&">"");"&vbcrlf 
	Else
	   Response.Write "alert('unKnow("&edNameID&")')"&vbcrlf 
	End If
	Response.End
	
ElseIf edtAct="mainEnd" Then 

	Response.End
    
Else '// mainShow Editor ... 

	edtCont = Get_Para("edtCont","InfCont")
	edtID   = Get_Para("edtID","EditID01")
	edtTool = Get_Para("edtTool","Peace") '//Peace,Basic,Full(def)
	edtData = Get_Para("edtData","")
	edtLine = Get_Para("edtLine","p") '//p/br  
	
	' js 公共变量
	Response.Write "var strEditID = '"&edtID&"'; "&vbcrlf 

	If edcEditID="edck3" Then  '// for CKEditor ////////////////////////////////	 
	
	  newLine = IIf(edtLine="br", "br", "p")
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);")
Response.Write "CKEDITOR.replace('"&edtCont&"',{"&vbcrlf
Response.Write "  skin : 'office2003',"&vbcrlf
If edtTool<>"Full" Then
Response.Write "  toolbar : 'toolBar_"&edtTool&"',"&vbcrlf
End If
Response.Write "  width : document.getElementById('"&edtCont&"').style.width,"&vbcrlf
Response.Write "  height : document.getElementById('"&edtCont&"').style.height,"&vbcrlf
Response.Write "  language : 'zh-cn'"&vbcrlf
Response.Write "});"&dataDID&""&vbcrlf 
Response.Write "var CK3_"&edtID&" = CKEDITOR.instances."&edtCont&";"&vbcrlf
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return CK"&edtID&".getData();"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  CK3_"&edtID&".setData(Val);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  CK3_"&edtID&".setData(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  CK3_"&edtID&".insertHtml(Val);"&vbcrlf
Response.Write "}"&vbcrlf

	ElseIf edcEditID="edkind" Then '// for KindEditor ///////////////////////////////////////
	
	  'newLine = IIf(edtLine="br", "br", "p")
	  'Response.Write "	newlineTag : '"&newLine&"',"&vbcrlf
	  'Response.Write "	fileManagerJson : '"&edcRootUrl&"/smod/file/edt_view.asp',"&vbcrlf 'edt_view.asp
	  'Response.Write "	allowFileManager : false,"&vbcrlf 'false,true
	  'Response.Write "	emoticonsPath : '"&edcRootUrl&"/img/',"&vbcrlf
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);")
Response.Write "var Kind_"&edtID&"; //放在.ready外面 "&vbcrlf
Response.Write "KindEditor.ready(function(K) {"&vbcrlf
Response.Write "  document.getElementById('"&edtCont&"').style.display = '';"&vbcrlf
Response.Write "  Kind_"&edtID&" = K.create('#"&edtCont&"',{"&vbcrlf
Response.Write "	uploadJson : '"&edcRootUrl&"/smod/file/edt_up.asp',"&vbcrlf
Response.Write "	urlType : 'absolute',"&vbcrlf 'relative为相对路径，absolute为绝对路径，domain为带域名的绝对路径
If edtTool<>"Full" Then
Response.Write "	items : toolBar_"&edtTool&","&vbcrlf
End If
Response.Write "	cssPath : '../plugins/code/prettify.css'"&vbcrlf
Response.Write "  }); "&dataDID&" //放在.ready里面"&vbcrlf
Response.Write "}); "&vbcrlf
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return Kind_"&edtID&".html(); //Kind_"&edtID&".sync();"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  Kind_"&edtID&".html(Val);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  Kind_"&edtID&".html(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  Kind_"&edtID&".insertHtml(Val);"&vbcrlf
Response.Write "}"&vbcrlf

	ElseIf edcEditID="edxh1" Then  '// for xhEditor ////////////////////////////////	

	  newLine = IIf(edtLine="br", "br", "p")
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);")
Response.Write "var xhEditor"&edtID&";"&vbcrlf
Response.Write "$(pageInit);"&vbcrlf
Response.Write "function pageInit()"&vbcrlf
Response.Write "{"&vbcrlf
Response.Write "    allPlugin = peacePlugins;"&vbcrlf
Response.Write "	xhEditor"&edtID&"=$('#"&edtCont&"').xheditor({"&vbcrlf
If edtTool<>"Full" Then
Response.Write "		plugins:allPlugin,"&vbcrlf
Response.Write "	    tools : toolBar_"&edtTool&","&vbcrlf
Else
Response.Write "	    tools : 'mfull',"&vbcrlf 'mfull,mfull,simple
End If
Response.Write "		loadCSS:'<style>pre{margin-left:2em;border-left:3px solid #CCC;padding:0 1em;}</style>',"&vbcrlf
Response.Write "		html5Upload: false,"&vbcrlf
Response.Write "		upLinkUrl: '"&edcRootUrl&"/smod/file/edt_up.asp',upLinkExt: 'zip,rar,txt',"&vbcrlf
Response.Write "		upImgUrl:  '"&edcRootUrl&"/smod/file/edt_up.asp',upImgExt:  'jpg,jpeg,gif,png',"&vbcrlf
Response.Write "		upFlashUrl:'"&edcRootUrl&"/smod/file/edt_up.asp',upFlashExt:'flv,swf',"&vbcrlf
Response.Write "		upMediaUrl:'"&edcRootUrl&"/smod/file/edt_up.asp',upMediaExt:'wmv,avi,wma,mp3,mid',"&vbcrlf
Response.Write "		emots:emots"&vbcrlf
Response.Write "	});"&dataDID&""&vbcrlf 
Response.Write "}"&vbcrlf
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return xhEditor"&edtID&".getSource();"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  xhEditor"&edtID&".setSource(Val);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  xhEditor"&edtID&".setSource(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  xhEditor"&edtID&".pasteHTML(Val);"&vbcrlf
Response.Write "}"&vbcrlf
	
	ElseIf edcEditID="edue1" Then  '// for UEditor ////////////////////////////////	
	
	  newLine = IIf(edtLine="br", "br", "p")
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);")
Response.Write "document.write(""<div id='eBox_"&edtID&"'></div>"");"&vbcrlf 
Response.Write "document.getElementById('eBox_"&edtID&"').style.width = document.getElementById('"&edtCont&"').style.width;"&vbcrlf
Response.Write "var eBoxHeight = document.getElementById('"&edtCont&"').style.height; eBoxHeight = eBoxHeight.replace('px',''); "&vbcrlf
Response.Write "var UE1_"&edtID&" = new baidu.editor.ui.Editor({"&vbcrlf
Response.Write "UEDITOR_HOME_URL: '"&edcEditUrl&"/',"&vbcrlf
Response.Write "id : 'eBox_"&edtID&"',"&vbcrlf
Response.Write "textarea : '"&edtCont&"',"&vbcrlf
Response.Write "iframeCssUrl :'"&edcEditUrl&"/themes/default/iframe.css',"&vbcrlf
Response.Write "initialContent: document.getElementById('"&edtCont&"').value,"&vbcrlf
Response.Write "minFrameHeight: eBoxHeight,"&vbcrlf
If edtTool<>"Full" Then
Response.Write "toolbars : toolBar_"&edtTool&","&vbcrlf
End If
Response.Write "autoHeightEnabled: false,"&vbcrlf
Response.Write "autoClearinitialContent:false"&vbcrlf
Response.Write "});"&dataDID&vbcrlf
Response.Write "UE1_"&edtID&".render('eBox_"&edtID&"');"&vbcrlf 
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return UE1_"&edtID&".getContent();"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  UE1_"&edtID&".setContent(Val) ; "&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  UE1_"&edtID&".setContent(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  UE1_"&edtID&".pasteContent(Val);"&vbcrlf
Response.Write "}"&vbcrlf

	ElseIf edcEditID="edtiny" Then '// for TinyEditor ////////////////////////////////	
	
	  newLine = IIf(edtLine="br", "br", "p")
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);") 
Response.Write "var objContID = document.getElementById('"&edtCont&"');"&vbcrlf
Response.Write "    objContID.style.display = '';"&vbcrlf
Response.Write "    objContID.style.visibility = 'visible';"&vbcrlf
Response.Write "tinyMCE.init({"&vbcrlf
Response.Write "  // General options"&vbcrlf
Response.Write "  mode : 'textareas',"&vbcrlf
Response.Write "  plugins : toolBar_Plugs2,"&vbcrlf
Response.Write "  theme : 'advanced',"&vbcrlf 'advanced,simple
Response.Write "  theme_advanced_buttons1 : toolBar_"&edtTool&"1,"&vbcrlf
Response.Write "  theme_advanced_buttons2 : toolBar_"&edtTool&"2,"&vbcrlf
Response.Write "  theme_advanced_buttons3 : toolBar_"&edtTool&"3,"&vbcrlf
Response.Write "  theme_advanced_buttons4 : toolBar_"&edtTool&"4,"&vbcrlf
Response.Write "  // Setting"&vbcrlf
Response.Write "  content_css : '"&edcEditUrl&"/themes/content.css',"&vbcrlf '../../sadm/edtiny
Response.Write "  theme_advanced_toolbar_location : 'top',"&vbcrlf
Response.Write "  theme_advanced_toolbar_align : 'left',"&vbcrlf
Response.Write "  theme_advanced_statusbar_location : 'bottom',"&vbcrlf
Response.Write "  theme_advanced_resizing : true"&vbcrlf
Response.Write "});"&dataDID&vbcrlf
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return tinyMCE.get('"&edtCont&"').getContent();"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  tinyMCE.get('"&edtCont&"').setContent(Val);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  tinyMCE.get('"&edtCont&"').setContent(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  tinyMCE.execCommand('mceInsertContent',false,Val);"&vbcrlf
Response.Write "}"&vbcrlf 
	
	ElseIf edcEditID="edfck" Then '// for FCKEditor /////////////////////////////////// 

	  newLine = IIf(edtLine="br", "br", "p")
	  dataDID = IIf(edtData="", "", "setTimeout(""apiSetVID"&edtID&"('"&edtData&"')"",500);")
Response.Write "var objFCK = new FCKeditor('"&edtID&"');"&vbcrlf
Response.Write "objFCK.BasePath	= '"&edcEditUrl&"/'; "&vbcrlf
Response.Write "objFCK.Width	= document.getElementById('"&edtCont&"').style.width;"&vbcrlf
Response.Write "objFCK.Height	= document.getElementById('"&edtCont&"').style.height;"&vbcrlf
Response.Write "objFCK.Value	= document.getElementById('"&edtCont&"').value;"&vbcrlf
If edtTool<>"Full" Then
Response.Write "objFCK.ToolbarSet = 'toolBar_"&edtTool&"';"&vbcrlf
End If
Response.Write "objFCK.Create(); "&dataDID&" //"&vbcrlf
Response.Write "function apiGetVal"&edtID&"(){"&vbcrlf
Response.Write "  return FCKeditorAPI.GetInstance('"&edtID&"').GetXHTML(true);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVal"&edtID&"(Val){"&vbcrlf
Response.Write "  var oEditor = FCKeditorAPI.GetInstance('"&edtID&"');"&vbcrlf
Response.Write "  oEditor.SetHTML(Val);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiSetVID"&edtID&"(id){"&vbcrlf
Response.Write "  var oEditor = FCKeditorAPI.GetInstance('"&edtID&"'); "&vbcrlf
Response.Write "  oEditor.SetHTML(document.getElementById(id).innerHTML);"&vbcrlf
Response.Write "}"&vbcrlf
Response.Write "function apiInsert"&edtID&"(Val){"&vbcrlf
Response.Write "  var oEditor = FCKeditorAPI.GetInstance('"&edtID&"');"&vbcrlf
Response.Write "  oEditor.InsertHtml(Val);"&vbcrlf
Response.Write "}"&vbcrlf

	Else
	   Response.Write "document.write('unKnow("&edNameID&")')"&vbcrlf 

	End If '//////////////////////////////////////////////////////////////////////////////	
	Response.End
	
End If

%>