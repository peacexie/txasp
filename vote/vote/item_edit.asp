<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../sadm/edfck/fckeditor.js"></script>
</head>
<body>
<%
KW = RequestS("KW",3,24)
ID = RequestS("ID",3,48)
iID = RequestS("iID",3,48)
send = Request("send")

If send = "send" Then

SetUBB = RequestS("SetUBB",3,2)
InfRem = RequestS("InfRem",3,9600) 
InfRem = Show_Html(InfRem)
sql = " UPDATE [VoteItem] SET " 
sql = sql& " InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",InfRem = '" & InfRem &"'" 
sql = sql& ",SetTop = '" & RequestS("SetTop","N",1024) &"'"
sql = sql& ",SetVote = '" & RequestS("SetVote","N",0) &"'"
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&iID&"' "
  Call rs_DoSql(conn,sql) 
  'Call rs_GetFile(ID,"")
  If Request("Act")="Close" Then
  Response.Write js_Alert("修改成功!","Close","")
  Else
  Response.Redirect "item_list.asp?ID="&ID&"&KW="&KW&"&Page="&PG
  End If
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [VoteItem] WHERE KeyID='"&iID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = Show_Form(rs("InfSubj"))
InfRem = rs("InfRem")&""
SetTop = rs("SetTop")
SetVote = rs("SetVote")
ImgName = rs("ImgName")
If ImgName<>"" then ronly = "readonly"
  else
KeyID = ""  
  end if 
rs.Close()
SET rs=Nothing 

If KeyID = "" Then Response.Redirect("info_list.asp")

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
      <td align="center" bgcolor="#FFFFFF"><strong>[投票项目]编辑</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120">      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">顺序</td>
      <td bgcolor="#FFFFFF"><input name="SetTop" type="text" id="SetTop" value="<%=SetTop%>" size="24" maxlength="24">        
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;选票
        <input name="SetVote" type="text" id="SetVote" value="<%=SetVote%>" size="24" maxlength="24"></td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">投票须知<br>
        </td>
      <td bgcolor="#FFFFFF">
<script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID01' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 360 ;
oFCKeditor.Value	= '<%=Show_jsStr(InfRem)%>' ;
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
	  <input name="InfRem" type="hidden" id="InfRem" value="">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">图片1</td>
      <td bgcolor="#FFFFFF"><IFRAME marginWidth=0 marginHeight=0
         src="../file/img_set.asp?TabID=VoteItem&ID=<%=KeyID%>"
          width="420" height='24' frameBorder=0 scrolling=no>
        </IFRAME>      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send">      </td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
      <input name="Act" type="hidden" id="Act" value="<%=Request("Act")%>">
      <input name="iID" type="hidden" id="iID" value="<%=iID%>"></td>
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
 
	document.fm01.InfRem.value = FCKeditorAPI.GetInstance("EditID01").GetXHTML(true);
 if (document.fm01.InfRem.value.length==0) 
   {   
     alert(" 内容 不能为空！"); 
     eflag = 1; break;
   }  
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
