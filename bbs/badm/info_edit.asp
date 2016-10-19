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

TP = RequestS("TP",3,48)	
KW = RequestS("KW",3,12)
Page = RequestS("Page","N",1)
KeyID = RequestS("KeyID",3,48)
send = Request("send")

if send = "edt" then
sql = " UPDATE [BPKTitle] SET "
sql = sql& " InfType = '" & RequestS("InfType",3,255) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",InfCont = '" & RequestS("InfCont",3,4000) &"'" 
sql = sql& ",InfOrg = '" & RequestS("InfOrg",3,255) &"'" 
sql = sql& ",InfOUrl = '" & RequestS("InfOUrl",3,255) &"'" 
'sql = sql& ",InfMemb = '" & RequestS("InfMemb",3,48) &"'" 
sql = sql& ",InfStart = '" & RequestS("InfStart","D","1900-12-31") &"'" 
sql = sql& ",InfEnd = '" & RequestS("InfEnd","D","1900-12-31") &"'" 
sql = sql& ",InfView1 = '" & RequestS("InfView1",3,255) &"'" 
sql = sql& ",InfView2 = '" & RequestS("InfView2",3,255) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("USID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&KeyID&"' "
Call rs_DoSql(conn,sql) 
  Response.Redirect "info_list.asp?Page="&Page&"&TP="&TP&"&KW="&KW&""
end if

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [BPKTitle] WHERE KeyID='"&KeyID&"' ",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
InfType = rs("InfType")
InfSubj = rs("InfSubj")
InfCont = rs("InfCont")
InfOrg = rs("InfOrg")
InfOUrl = rs("InfOUrl")
InfMemb = rs("InfMemb")
InfStart = rs("InfStart")
InfEnd = rs("InfEnd")
InfView1 = rs("InfView1")
InfView2 = rs("InfView2")
InfVote1 = rs("InfVote1")
InfVote2 = rs("InfVote2")
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")

else
  Response.Write js_Alert("错误,论坛不存在!","Back",-1)   
end if 
rs.Close()
SET rs=Nothing

%>
        <table border="0" align="center" cellpadding="1" cellspacing="1">

          <tr>
            <td valign="top"><table width="540"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="f4f4f4">
              <form name="fm01" method=post action="?">
                <tr>
                  <td height="24"><a href="#">返回</a></td>
                  <td><strong>论题修改</strong></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" bgcolor="#999999"></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">标题:</td>
                  <td><input name=InfSubj id=InfSubj value="<%=InfSubj%>" size=48 
maxlength=48 ></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">类别:</td>
                  <td><select name="InfType" id="InfType" style="width:300px; ">
                      <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"'",InfType)%>
                  </select></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">说明:</td>
                  <td><textarea name="InfCont" cols="48" rows="8" id="InfCont" ><%=InfCont%></textarea></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">正方观点:</td>
                  <td><input name=InfView1 id=InfView1 value="<%=InfView1%>" size=48 
maxlength=48 ></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">反方观点:</td>
                  <td><input name=InfView2 id=InfView2 value="<%=InfView2%>" size=48 
maxlength=48 ></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">辩论起止:</td>
                  <td><input name=InfStart id=InfStart value="<%=InfStart%>" size=12 
maxlength=10 >
        ~
          <input id=InfEnd 
maxlength=10 size=12 name=InfEnd value="<%=InfEnd%>" >
        (格式:yyyy-MM-DD)</td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">论题来源:</td>
                  <td><input name=InfOrg id=InfOrg value="<%=InfOrg%>" size=48 
maxlength=48 ></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td align="right">来源地址:</td>
                  <td><input name=InfOUrl id=InfOUrl value="<%=InfOUrl%>" size=48 
maxlength=255 ></td>
                </tr>
                <tr bgcolor="ffffff">
                  <td><font color="#FF0000">
                    <input name="TP" type="hidden" id="TP" value="<%=TP%>">
                    <input name="KW" type="hidden" id="KW" value="<%=KW%>">
                    <input name="Page" type="hidden" id="Page" value="<%=Page%>">
                  </font></td>
                  <td><input type="button" name="Button" value=" 提 交 " onClick="chkData()">
&nbsp;
        <input name="send" type="hidden" id="send" value="edt">
        <font color="#FF0000">
        <input type="reset" name="Reset" value=" 重设 " >
        <%=MSG%>        
        <input name="KeyID" type="hidden" id="KeyID" value="<%=KeyID%>">
</font></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </form>
            </table>
            </td>
            <td align="center" valign="top" width="320">&nbsp;</td>
          </tr>
</table>

<script type="text/javascript">

function owInfoUpd()
    { 
	switch (owupdNO)
	{
	  case "fm01" : 
	  { document.fm01.InfStart.value=owArr[0];
	    document.fm01.InfStartC.value=owArr[1];
	  break;}
	  case "fm02" : 
	  { document.fm01.InfEnd.value=owArr[0];
	    document.fm01.InfEndC.value=owArr[1];
	  break;}
	  case "fm03" : 
	  { document.ff.Addr3.value=owArr[0];
	    document.ff.Addr3X.value=owArr[1];
	  break;}
	  default : { break; }
	}
}

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
     document.fm01.View1.focus();
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
