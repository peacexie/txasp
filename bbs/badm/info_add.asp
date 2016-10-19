<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>后台管理中心</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<!--#include file="conpub.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%

send = Request("send") 

if send = "ins" then

KeyID = get_AutoID(24)

sql = " INSERT INTO [BPKTitle] (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfOrg" 
sql = sql& ", InfOUrl" 
sql = sql& ", InfMemb" 
sql = sql& ", InfStart" 
sql = sql& ", InfEnd" 
sql = sql& ", InfView1" 
sql = sql& ", InfView2" 
sql = sql& ", InfVote1" 
sql = sql& ", InfVote2" 
sql = sql& ", SetRead" 
sql = sql& ", SetHot" 
sql = sql& ", SetShow" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & RequestS("InfType",3,255) &"'" 
sql = sql& ", '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ", '" & RequestS("InfCont",3,4000) &"'" 
sql = sql& ", '" & RequestS("InfOrg",3,255) &"'" 
sql = sql& ", '" & RequestS("InfOUrl",3,255) &"'" 
sql = sql& ", '(Admin)'" 'InfMemb
sql = sql& ", '" & RequestS("InfStart","D","1900-12-31") &"'" 
sql = sql& ", '" & RequestS("InfEnd","D","1900-12-31") &"'" 
sql = sql& ", '" & RequestS("InfView1",3,255) &"'" 
sql = sql& ", '" & RequestS("InfView2",3,255) &"'" 
sql = sql& ", " & RequestS("InfVote1","N",0) &"" 
sql = sql& ", " & RequestS("InfVote2","N",0) &"" 
sql = sql& ", " & RequestS("SetRead","N",0) &"" 
sql = sql& ", 'N'" 'SetHot
sql = sql& ", 'Y'" 'SetShow 
sql = sql& ", '" & RequestS("ImgName",3,48) &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("USID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
Call rs_DoSql(conn,sql) 
Response.Redirect "info_list.asp"
msg = "增加成功！"
end if


%>
<br>
<table width="540"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="f4f4f4">
  <form name="fm01" method=post action="?">
    <tr>
      <td height="24">&nbsp;</td>
      <td><strong>论题增加</strong></td>
    </tr>
    <tr>
      <td colspan="2" align="right" bgcolor="#999999"></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">标题:</td>
      <td><input name=InfSubj id=InfSubj size=48 
maxlength=48 ></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">类别:</td>
      <td><select name="InfType" id="InfType" style="width:300px; ">
          <option value="">[无分类]</option>
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"'",TP)%>
        </select></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">说明:</td>
      <td><textarea name="InfCont" cols="48" rows="8" id="InfCont" ></textarea></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">正方观点:</td>
      <td><input name=InfView1 id=InfView1 size=48 
maxlength=48 ></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">反方观点:</td>
      <td><input name=InfView2 id=InfView2 size=48 
maxlength=48 ></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">辩论起止:</td>
      <td><input name=InfStart id=InfStart value="<%=Date()%>" size=12 
maxlength=10 >
        ~
        <input id=InfEnd 
maxlength=10 size=12 name=InfEnd value="<%=DateAdd("m",1,Date())%>" >
        (格式:yyyy-MM-DD)</td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">论题来源:</td>
      <td><input name=InfOrg id=InfOrg size=48 
maxlength=48 ></td>
    </tr>
    <tr bgcolor="ffffff">
      <td align="right">来源地址:</td>
      <td><input name=InfOUrl id=InfOUrl size=48 
maxlength=255 ></td>
    </tr>
    <tr bgcolor="ffffff">
      <td>&nbsp;</td>
      <td><input type="button" name="Button" value=" 提 交 " onClick="chkData()">
        &nbsp;
        <input name="send" type="hidden" id="send" value="ins">
        <font color="#FF0000"><%=MSG%>
        <input type="reset" name="Reset" value=" 重设 " >
        </font></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </form>
</table>
<script type="text/javascript">

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfView1.value.length==0) 
   {   
     alert(" 正方观点 不能为空！"); 
     document.fm01.InfView1.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfView2.value.length==0) 
   {   
     alert(" 反方观点 不能为空！"); 
     document.fm01.InfView2.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont.value.length>1200) 
   {   
     alert(" 说明 不能超过1200字!"); 
     document.fm01.InfCont.value = document.fm01.InfCont.value.substr(0,1200-3)+"...";
     document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 //tmv = chkF_Date(document.fm01.InfMemb,"日期 不规范！");
 //if (tmv=='ER') 
   {         //alert(""); //document.fm01.VSStart.focus();
     //eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
