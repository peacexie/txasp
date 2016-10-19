<!--#include file="dinc/_config.asp"-->
<%

send = RequestS("send","C",12)
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,48)
MD = "DocD124"

ObjSubj = RequestS("ObjSubj","C",255) 
ObjID = RequestS("ObjID","C",48)
ObjUrl = RequestS("ObjUrl","C",255) 
If ObjUrl="" Then
ObjUrl = Request.Servervariables("HTTP_REFERER")
End If
'Response.Write ObjSubj&ObjID&ObjUrl
 
send = Request("send")
'Response.Write ChkCode&"dd"

If send="send" Then

  Call Chk_URL()	
  'Response.End()
  IP = Get_CIP()
  InfSubj = RequestS("InfSubj",3,255)
  InfCont = RequestS("InfCont",3,512)
  LogAUser = RequestS("LogAUser",3,48)
  InfCont = Show_Text(InfCont)

sql = " INSERT INTO [DocsRemark] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyCode" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont"  
sql = sql& ", SetShow"  
sql = sql& ", LogAddIP" 
sql = sql& ", LogATime" 
sql = sql& ", LogAUser" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & Get_AutoID(24) &"'" 
sql = sql& ", '" & MD & "'" 
sql = sql& ", '" & ObjID &"'" 
sql = sql& ", '" & RequestS("InfType","C",24) &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfCont &"'" 
sql = sql& ", '" & RemShow &"'" 
sql = sql& ", '" & IP &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '" & RequestS("LogAUser","C",48) &"'" 
sql = sql& ")"

  If InfSubj<>"" AND InfCont<>"" Then 
    Call rs_DoSql(conn,sql) 
    rMsg = "提交OK!"
	Response.Write js_Alert(rMsg,"Redir","?ModID="&MD&"&ObjID="&ObjID&"&ObjSubj="&Server.URLEncode(ObjSubj)&"&ObjUrl="&ObjUrl&"")
  Else
      rMsg = vIRem_MsgNG&fADD
  End If ' End Pass
  
End If

ObjSubj = Show_Text(ObjSubj)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>评论公文 - <%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  

<FIELDSET style="padding:3px; text-align:left">
<LEGEND style="padding:5px;"><strong>公文评论</strong></LEGEND>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr align="center" bgcolor="#FFFFFF">
    <td width="60%" align="left" bgcolor="#FFFFFF">原文: <a href="<%=ObjUrl%>"><%=ObjSubj%></a></td>
    <form name="fm02" method="post" action="?">
      <td align="right" nowrap><font color="#FF0000">
        <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
        <input name="ObjID" type="hidden" id="ObjID" value="<%=ObjID%>" />
        <input name="ObjSubj" type="hidden" id="ObjSubj" value="<%=ObjSubj%>" />
        <input name="ObjUrl" type="hidden" id="ObjUrl" value="<%=ObjUrl%>" />
        <%=rMsg%></font>&nbsp;关键字:
        <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
        <input type="submit" name="Submit" value="搜索">
        </td>
    </form>
  </tr>
</table>
</FIELDSET>
<div style="line-height:8px;">&nbsp;</div>
<%
sqlK = ""
KW = RequestS("KW","C",24)
If KW<>"" Then
  sqlK=sqlK&" AND (InfSubj LIKE '%"&KW&"%' OR InfCont LIKE '%"&KW&"%') "
End If

    sql = " SELECT DocsRemark.* FROM [DocsRemark] "
	sql =sql& " WHERE KeyMod='"&MD&"' AND KeyCode='"&ObjID&"' "&sqlK
	sql =sql& " ORDER BY KeyID DESC" 'SetTop,
   rs.Open Sql,conn,1,1
   rs.PageSize = 24 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If
%>
<%
  
  If rs.eof then
  %>
<FIELDSET style="padding:3px;">
<LEGEND style="padding:5px;">无资料</LEGEND>
</FIELDSET>
<%
  Else
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
LogAddIP = rs("LogAddIP") 
InfSubj = Show_Text(InfSubj) 
InfCont = rs("InfCont")
	  %>
<FIELDSET style="background:<%=col%>">
<LEGEND><B><%=InfSubj%></B> <a href="<%=ObjUrl%>">原文</a></LEGEND>
<table width="100%"  border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td align="left" valign="top"><%=InfCont%></td>
    <td width="15%" align="left" valign="top" nowrap style="line-height:120%;"><%=LogATime%><br>
      评论人: <%=LogAUser%><br>
      IP:<%=LogAddIP%> </td>
  </tr>
</table>
</FIELDSET>
<div style="line-height:8px;">&nbsp;</div>
<%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
<FIELDSET style="padding:3px;">
<%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&ObjID="&ObjID&"&ObjSubj="&ObjSubj&"&KW="&KW&"",1)%>
</FIELDSET>
<%    
  End If
	  rs.Close()
	  
	  %>
<div style="line-height:8px;">&nbsp;</div>
<FIELDSET style="padding:3px;">
<LEGEND style="padding:5px;"><B>添加评论</B> <a href="<%=ObjUrl%>">原文</a></LEGEND>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td align="left" bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="评论: <%=ObjSubj%>" size="60" maxlength="120">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">姓名</td>
      <td align="left" bgcolor="#FFFFFF"><input name="LogAUser" type="text" id="LogAUser" value="<%=Session("InnID")%>" size="18" maxlength="24" />
        &nbsp;&nbsp;态度:
        <select name="InfType" id="InfType">
          <option value="支持">支持</option>
          <option value="反对">反对</option>
          <option value="中立" selected="selected">中立</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">内容<br />
        (255字内)</td>
      <td align="left" bgcolor="#FFFFFF"><textarea name="InfCont" cols="50" rows="8" id="InfCont"></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send" />
        </td>
      <td align="left" bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="提交" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重写">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
        <input name="ObjID" type="hidden" id="ObjID" value="<%=ObjID%>">
        <input name="ObjSubj" type="hidden" id="ObjSubj" value="<%=ObjSubj%>" />
        <input name="ObjUrl" type="hidden" id="ObjUrl" value="<%=ObjUrl%>" /></td>
    </tr>
  </form>
</table>
</FIELDSET>
<script type="text/javascript">
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" <%=vGbo_jsSubj%>"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length==0) 
   {   
     alert(" <%=vGbo_jsCont%>"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length>=255) 
   {   
     alert(" <%=vGbo_jsCMax%>255!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>

  
</div>
<%jsFlag="mTopHome"%>
<!--#include file="dinc/inc_bot.asp"-->
<%
SET rs=Nothing 
%>
</body>
</html>
