<!--#include file="dinc/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>发布公文 -<%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../sadm/edfck/fckeditor.js" type="text/javascript"></script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  <%
  
send = Request("send")
ReEnd = Request("ReEnd")
InfType = RequestS("InfType","C",255)

If send="ins" Then
 If Len(Session("KeyID"))<15 Then 
  KeyID = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
 Else
  KeyID = Session("KeyID")
 End If
 upPath = upRoot&Replace(KeyID,"-","/")&"/" 
IP = Get_CIP()
InfSubj = RequestS("InfSubj",3,255) 
InfTo = RequestS("InfView",3,8000) : InfTo=Replace(InfTo,Session("InnID")&";","")
If Request("sAll01")="(__Public__)" Then InfTo="(__Public__)" 
'Response.Write InfTo
InfCont = RequestS("InfCont",3,960000) 
InfCont = Show_Html(InfCont)
 If Config_Cont="DB" Then
  xxxCont = InfCont
 Else
  xxxCont = ""
 End If
sql = " INSERT INTO [DocsNews] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfTyp2" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfTo" 
sql = sql& ", SetSubj" 
sql = sql& ", SetRead" 
sql = sql& ", SetHot" 
sql = sql& ", SetTop" 
sql = sql& ", SetShow" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & Get_FmtID("mdhnsx","")&"-"&Rnd_ID("KEY",6) &"'"  
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & RequestS("InfTyp2",3,48) &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & InfTo &"'" 
sql = sql& ", '" & RequestS("SetSubj",3,12) &"'" 
sql = sql& ", 0" 
sql = sql& ", '" & RequestS("SetHot",3,2) &"'" 
sql = sql& ", '" & RequestS("SetTop",3,12) &"'" 
sql = sql& ", '" & RequestS("SetShow",3,2) &"'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Session("InnID") &"'" 
sql = sql& ", '" & RequestS("LogATime","D",Now()) &"'" 
sql = sql& ")"
  If Trim(InfSubj)<>"" Then 
    Call rs_Dosql(conn,sql)	
	Call add_sfFile()
	
	ReDir = "info_list.asp"
	If ReEnd="Y" Then ReDir="info_add.asp?ReEnd="&ReEnd&"&InfType="&InfType&""
	Response.Write js_Alert("发布成功！","Redir",ReDir) 
	
  End If
End If

tmpCont = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp"&ModID&"'") 'Server.HTMLEncode()
'Response.Write Session("ImgList")

Session("KeyID") = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
InfTyp2 = rs_Val("","SELECT UsrType FROM AdmUser"&Adm_aUser&" WHERE UsrID='"&Session("InnID")&"'")

%>
  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
    <form name="fm01" id="fm01" action="?" method="post">
      <tr bgcolor="#FFFFFF">
        <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
        <td align="left" bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" size="60" maxlength="120">
          <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90 ">
            <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>-[默认]-</OPTION>
            <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
            <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
            <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
            <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
            <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
            <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
          </select></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" nowrap bgcolor="#FFFFFF">类别</td> 
        <td align="left" bgcolor="#FFFFFF"><select name="InfType" id="InfType">
            <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay")%>
          </select>
          &nbsp;&nbsp;
          发布部门 
          <select name="InfTyp2" id="InfTyp2">
            <option value='<%=InfTyp2%>'>[<%=GetGList("Val",InfTyp2)%>]</option>
          </select>
          </td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="#FFFFFF">查看人</td>
        <td align="left" bgcolor="#FFFFFF"><!--#include file="dinc/inc_group.asp"--></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="#FFFFFF">内容</td>
        <td align="left" bgcolor="#FFFFFF">
<script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID01' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 360 ;
oFCKeditor.Value	= '<%=Show_jsStr(tmpCont)%>' ;
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
		
          <input name="InfCont" type="hidden" value=""></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" bgcolor="#FFFFFF">属性</td>
        <td align="left" bgcolor="#FFFFFF"><select name="SetHot" id="SetHot">
            <option value="N" selected >平常</option>
            <option value="Y" >推荐</option>
          </select>
          <select name="SetShow" id="SetShow">
            <option value="Y" selected >显示</option>
            <option value="N" >隐藏</option>
          </select>
          <select name="SetTop" id="SetTop">
            <option value="888" selected >顺序</option>
            <%
	  For i = 0 to 9
	  sel = ""
	  %>
            <option value="<%=i%>"><%=i%></option>
            <%Next%>
          </select>
          时间
          <input name="LogATime" type="text" id="LogATime" value="<%=Now()%>" size="20" maxlength="20"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" nowrap bgcolor="#FFFFFF">返回</td>
        <td align="left" bgcolor="#FFFFFF"><input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
          添加资料后返回列表
          &nbsp;&nbsp;&nbsp;
          <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
          添加资料后继续</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="center" nowrap bgcolor="#FFFFFF">&nbsp;</td>
        <td align="left" bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
          <input name=view type=reset id="Button1" value="重    写">
          <input name="send" type="hidden" id="send" value="ins" /></td>
      </tr>
    </form>
  </table>
  <script type="text/javascript">
 function chkData()
 {
       
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }

	//document.fm01.InfCont.value = Editor01.getHTML();
	document.fm01.InfCont.value = FCKeditorAPI.GetInstance("EditID01").GetXHTML(true);
 if (document.fm01.InfCont.value.length==0) 
   {   
     //alert(" 内容 不能为空！"); 
     //eflag = 1; break;
   } 
   
 if (document.fm01.InfCont.value.length>=480000) 
   {   
     alert(" 内容 不能超过 480 K字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 document.fm01.submit();
		 //Editor01.remoteUpload("document.fm01.submit();"); // 
		 }
}

</script>
</div>
<%jsFlag="mTopPub"%>
<!--#include file="dinc/inc_bot.asp"-->
<script type="text/javascript">ChkAll('Pub');</script>
</body>
</html>
