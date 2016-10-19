
<!--#include file="head_config.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<%

MD = RequestS("MD",3,48)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,240)
PG = RequestS("PG","N",1)
ID = RequestS("ID",3,48)

send = Request("send")

If send = "send" Then

InfCont = RequestS("InfCont",3,480) 
InfCont = Show_Html(InfCont)
InfSubj = RequestS("InfSubj",3,255) 
sql = " UPDATE [InfoHead] SET " 
sql = sql& " InfTyp2 = '" & RequestS("InfTyp2",3,255) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",InfCont = '" & InfCont &"'" 
sql = sql& ",SetSubj = '" & RequestS("SetSubj",3,12) &"'" 
sql = sql& ",ImgName = '" & RequestS("ImgNam3",3,180) &"'" 
sql = sql& ",LogATime = '" & RequestS("LogATime","D",Now()) &"'"
sql = sql& ",LogAUser = '" & RequestS("LogAUser","C",48) &"'"
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "
  Call rs_DoSql(conn,sql)
  If Request("Act")="Close" Then
    Response.Write js_Alert("修改成功!","Alert","")
  Else
	Response.Write js_Alert("修改成功!","Redir","head_edit.asp?ID="&ID&"&TP="&TP&"&KW="&KW&"&MD="&MD&"&Page="&PG)
	Msg = "修改成功!"
	'Response.Redirect "head_edit.asp?ID="&ID&"&TP="&TP&"&KW="&KW&"&MD="&MD&"&Page="&PG
  End If
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [InfoHead] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyRE = rs("KeyRE") 
KeyMod = rs("KeyMod")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = rs("InfCont")
InfCont = Show_Form(InfCont) 'Server.HTMLEncode() 
SetSubj = rs("SetSubj")
ImgName = rs("ImgName")
LogATime = rs("LogATime") 
LogAUser = rs("LogAUser") 
End If
rs.Close()
SET rs=Nothing 

If KeyID = "" Then Response.Redirect("head_list.asp")
%>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
      <td align="center" bgcolor="#FFFFFF"><strong>头条编辑</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF">
        颜色
        <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:120px">
          <OPTION style='COLOR:#<%=SetSubj%>; BACKGROUND-COLOR:#<%=SetSubj%>' value='<%=SetSubj%>'>#<%=SetSubj%></OPTION>
          <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>#000000</OPTION>
          <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
          <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
          <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
          <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
          <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
          <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
        </select>
        &nbsp; 位置
        <select name="InfTyp2" id="InfTyp2" style="width:120px">
          <option value=''>[位置]</option>
          <%=hOpt(TP2)%>
        </select> <a href="info_view.asp?ID=<%=KeyRe%>&ModID=<%=KeyMod%>" target="_blank">原文</a> </td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">内容</td>
      <td bgcolor="#FFFFFF"><textarea name="InfCont" cols="56" rows="8" id="InfCont"><%=InfCont%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">属性</td>
      <td bgcolor="#FFFFFF">
        时间
        <input name="LogATime" type="text" id="LogATime" value="<%=LogATime%>" size="20" maxlength="20" style="width:120px;"> 
        &nbsp; 用户
        <input name="LogAUser" type="text" id="LogAUser" value="<%=LogAUser%>" size="20" maxlength="20" style="width:120px;"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">缩略图</td>
      <td bgcolor="#FFFFFF"><input name="ImgNam3" type="text" id="ImgNam3" value="<%=ImgName%>" size="60" maxlength="120">
        <input name=view2 type=button id="Button2" value="选择" onClick="window.open('../file/file_view.asp?yPath=myfile/head/')">
        <input name="InfSpeci" type="hidden" id="InfSpeci" value="<%=InfSpeci%>" size="12" maxlength="12">        <input name="InfPrice" type="hidden" id="InfPrice" value="<%=InfPrice%>" size="12" maxlength="10"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send">      </td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="MD" type="hidden" id="MD" value="<%=MD%>">
<input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
      <input name="Act" type="hidden" id="Act" value="<%=Request("Act")%>"> 
      &nbsp;<a href="?head_list.asp?ID=<%=ID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&MD=<%=MD%>">&lt;&lt;&lt;返回</a> 
      &nbsp;<span class="colF00"><%=Msg%></span></td>
    </tr>
  </form>
</table>
<div style="line-height:8px">&nbsp;</div>
<%

If ImgName<>"" Then
  InfSubj = "<img src='"&Replace(Config_Path&"/"&ImgName,"//","/")&"' height='32' alt='"&InfSubj&"'>"
Else
  InfSubj = Show_sTitle(InfSubj,SetSubj)
End If
%>
<table width="640" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="F0F0F0">
    <tr bgcolor="#FFFFFF">
      <td width="10%" align="center" bgcolor="#FFFFFF"><p>标题</p></td>
      <td align="center" bgcolor="#FFFFFF"><%=InfSubj%></td>
  </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">内容</td>
      <td bgcolor="#FFFFFF"><%=InfCont%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">注意</td>
      <td bgcolor="#FFFFFF" width="572">图片显示依前台标准。</td>
    </tr>
</table>

<script type="text/javascript">

  var xFile,xSize;
    function owFileGet()
    {
	  var tFile = xFile; 
	  var nPos = tFile.indexOf('/upfile/');
	  var sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  document.fm01.ImgNam3.value = tFile;
	  nPos = tFile.indexOf('.');
	  sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  document.fm01.InfSpeci.value = tFile;
	  document.fm01.InfPrice.value = xSize;
    }

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

 if (document.fm01.InfCont.value.length==0) 
   {   
     alert(" 内容 不能为空！"); 
     eflag = 1; break;
   }   
   
 if (document.fm01.InfCont.value.length>=250) 
   {   
     alert(" 内容 不能超过 240 字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
