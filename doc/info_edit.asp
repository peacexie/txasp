<!--#include file="dinc/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>修改公文 -<%=sysName%></title>
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

KW = RequestS("KW",3,24)
TP = RequestS("TP",3,240)
PG = RequestS("PG","N",1)
ID = RequestS("ID",3,48)

send = Request("send")

If send = "send" Then

SetUBB = RequestS("SetUBB",3,2)
InfCont = RequestS("InfCont",3,960000) 
InfCont = Show_Html(InfCont)
If Config_Cont="DB" Then
  xxxCont = InfCont
Else
  xxxCont = ""
End If
InfSubj = RequestS("InfSubj",3,255) 
InfTo = RequestS("InfView",3,8000) : InfTo=Replace(InfTo,Session("InnID")&";","")
If Request("sAll01")="(__Public__)" Then InfTo="(__Public__)" 
docUser = Session("InnID")&""
If docUser="" Then
  docUser = Session("UsrID")&""
End If
sql = " UPDATE [DocsNews] SET " 
sql = sql& " InfType = '" & RequestS("InfType",3,255) &"'" 
'sql = sql& ",InfTyp2 = '" & RequestS("InfTyp2",3,255) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",InfTo = '" & Replace(InfTo," ","") &"'" 
sql = sql& ",InfCont = '" & xxxCont &"'" 
sql = sql& ",SetSubj = '" & RequestS("SetSubj",3,12) &"'" 
sql = sql& ",SetHot = '" & RequestS("SetHot",3,2) &"'" 
sql = sql& ",SetTop = '" & RequestS("SetTop",3,12) &"'" 
sql = sql& ",SetShow = '" & RequestS("SetShow",3,2) &"'" 
sql = sql& ",LogATime = '" & RequestS("LogATime","D",Now()) &"'"
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & docUser &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "
'If RequestS("SetUBB",3,2)="X" Then
'sql = sql& ",ImgNam2 = '" & Session("ImgList") &"'" 
'End If
  Call rs_DoSql(conn,sql)

  upPath = upRoot&Replace(ID,"-","/")&"/" 
  KeyID = ID
  Call add_sfFile()

  If Request("Act")="Close" Then
    Response.Write js_Alert("修改成功!","Close","")
  Else
	Response.Redirect "info_list.asp?TP="&TP&"&KW="&KW&"&Page="&PG
  End If
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [DocsNews] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Form(rs("InfSubj"))
InfTo = rs("InfTo")
xxxCont = rs("InfCont") 
InfCont = Show_sfRead(ID,"/fcont.htx")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime") 
End If
rs.Close()
SET rs=Nothing 

If KeyID = "" Then Response.Redirect("info_list.asp")
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td align="left" bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120">
          <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90 ">
            <OPTION style='COLOR:#<%=SetSubj%>; BACKGROUND-COLOR:#<%=SetSubj%>' value='<%=SetSubj%>'>#<%=SetSubj%></OPTION>
            <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>#000000</OPTION>
            <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
            <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
            <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
            <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
            <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
            <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
          </select>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">类别</td>
      <td align="left" bgcolor="#FFFFFF">
        <select name="InfType" id="InfType">
		  <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
        </select> 
        &nbsp;&nbsp;
        <%If SwhDepSubs="Y" Then%>
        部门/分公司
	  <select name="InfTyp2" id="InfTyp2">
          <option value=''>[不限部门]</option>
          <%=getGList("",InfTyp2)%>
        </select>
        <%End If%>	  </td>
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
oFCKeditor.Value	= '<%=Show_jsStr(InfCont)%>' ;
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script><input name="InfCont" type="hidden" value="">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">属性</td>
      <td align="left" bgcolor="#FFFFFF"><select name="SetHot" id="SetHot">
        <option value="N" <%If SetHot="N" Then Response.Write("selected")%>>平常</option>
        <option value="Y" <%If SetHot="Y" Then Response.Write("selected")%>>推荐</option>
      </select>
        <select name="SetShow" id="SetShow">
          <option value="Y" <%If SetShow="Y" Then Response.Write("selected")%>>显示</option>
          <option value="N" <%If SetShow="N" Then Response.Write("selected")%>>隐藏</option>
        </select>
        <select name="SetTop" id="SetTop">
          <option value="<%=SetTop%>" selected >顺序</option>
          <%
	  For i = 0 to 9
	      sel = ""
	    If CStr(i)=CStr(SetTop) Then
		  sel = "selected"
		End If
	  %>
          <option value="<%=i%>" <%=sel%>><%=i%></option>
          <%Next%>
        </select>
        时间
        <input name="LogATime" type="text" id="LogATime" value="<%=LogATime%>" size="20" maxlength="20">        &nbsp; <a href="../smod/file/img_list.asp?send=PicList&ID=<%=ID%>" target="_blank">附件管理</a></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send">      </td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
      <input name="Act" type="hidden" id="Act" value="<%=Request("Act")%>"></td>
    </tr>
  </form>
</table>

<script type="text/javascript">

<%If InfTo="(__Public__)" Then Response.Write("ChkAll('Pub');")%>

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
         if (eflag==0){ document.fm01.submit(); }
}</script>
</div>
<%jsFlag="mTopAdm"%>
<!--#include file="dinc/inc_bot.asp"-->
</body>
</html>
