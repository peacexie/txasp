<!--#include file="config.asp"-->
<!doctype html>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jsInfo.js"></script>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtAct=mainLoad"></script>
<style type="text/css">
tr, td {
	background-color:#FFF;
}
</style>
</head>
<body>
<%
 	
  reUrl1 = Config_Path&"smod/gbook/info_list.asp"
  reUrl2 = Config_Path&"smod/gbook/info_edit.asp"
  If Chk_URL3(reUrl1)="OK" Or Chk_URL3(reUrl2)="OK" Then
  Else
    Response.End()
  End If	

If (ModID="MemB224" AND PrmFlag="(Mem)") OR (ModID="MemB524" AND PrmFlag="(Inn)") Then
  Response.Write "<h1 style='line-height:180%; padding:50px'>系统通知 不能修改</h1>"
  Response.End()
End If
		 
send = Request("send")
PG = RequestS("PG","N",1)
KW = RequestS("KW",3,48)
ID = RequestS("ID",3,48)
TP = RequestS("TP",3,48)
InfSubj = RequestS("InfSubj",C,255)

If send="ins" Then 

InfCont = Show_Html(RequestS("InfCont",C,24000))
InfReply = Show_Html(RequestS("InfReply",C,24000))
 If Config_Cont="DB" Then
  xxxCont = InfCont
  xxxReply = InfReply
 Else
  xxxCont = ""
  xxxReply = ""
 End If
sql = " UPDATE "&ModTab&" SET " 
sql = sql& " InfSubj = '" & InfSubj &"'" 
sql = sql& ",InfType = '" & RequestS("InfType",C,255) &"'" 
sql = sql& ",InfCont = '" & xxxCont &"'" 
sql = sql& ",InfReply = '" & xxxReply &"'"  
sql = sql& ",SetShow = '" & RequestS("SetShow",C,2) &"'"
sql = sql& ",LnkName = '" & RequestS("LnkName",C,48) &"'"
sql = sql& ",LnkEmail = '" & RequestS("LnkEmail",C,120) &"'"
sql = sql& ",ImgName = '" & RequestS("ImgName",C,255) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Get_PUser(PrmFlag) &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "
 If Len(InfSubj)>0 And Len(InfCont)>0 Then
  Call rs_DoSql(conn,sql) 
  upPath = upRoot&Replace(ID,"-","/") 
  Call add_sfFile()
 End If
  jsMsg = "信息修改/回复成功！" ': Response.Write sql
  'Response.Write sql
  Response.Write js_Alert(jsMsg,"Redir","info_list.asp?KW="&KW&"&TP="&TP&"&Page="&PG&"&ModID="&ModID&"&PrmFlag="&PrmFlag&"") 

End If



SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"'",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = Show_Form(rs("InfType"))
InfSubj = Show_Form(rs("InfSubj"))
xxxCont = rs("InfCont") 
InfCont = Show_sfRead(ID,".org.htm")
xxxReply = rs("InfReply") 
InfReply = Show_sfRead(ID,".rep.htm")
LnkName = Show_Form(rs("LnkName"))
LnkEmail = Show_Form(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
ImgName = Show_Form(rs("ImgName"))
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
ArrFlag = Split(KeyFlag&"^","^")
KeyFlag = Replace(KeyFlag,"^"," / ")
end if 
rs.Close()
SET rs=Nothing 

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="right"><strong>信息增加</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>主题</td>
      <td><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120"></td>
    </tr>
    <tr>
      <td align="center">类别</td>
      <td><select name="InfType" style="width:120px; " id="InfType">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypID",InfType)%>
        </select>
        &nbsp; &nbsp;
        姓名
        <input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="24" maxlength="24" <%If ModID="Gbo936" Then Response.Write("readonly") %> ></td>
    </tr>
    <tr>
      <td align="center" nowrap>审核</td>
      <td><select name="SetShow" id="SetShow" style="width:120px;">
          <option value="Y"  <%if SetShow="Y" then Response.Write("selected")%>>已审核</option>
          <option value="N"  <%if SetShow="N" then Response.Write("selected")%>>未审核</option>
        </select>
        &nbsp; &nbsp; 邮箱
        <input name="LnkEmail" type="text" id="LnkEmail" value="<%=LnkEmail%>" size="24" maxlength="60"></td>
    </tr>
    <tr>
      <td align="center">内容</td>
      <td>
	  
<textarea id="InfCont" name="InfCont" style="width:580px;height:360px;visibility:hidden;display:none"><%=Show_Form(InfCont)%></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID01&edtCont=InfCont&edtTool=Basic"></script>

	  </td>
    </tr>
    <%If (PrmFlag="(Mem)" AND ModID="MemB124") OR (PrmFlag="(Inn)" AND ModID="MemB424") Then%>
    <%Else%>
    <tr>
      <td align="center">回复</td>
      <td>
	  
<textarea id="InfReply" name="InfReply" style="width:580px;height:240px;visibility:hidden;display:none"><%=Show_Form(InfReply)%></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID02&edtCont=InfReply&edtTool=Basic"></script>

      </td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" nowrap><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
        注意:留言12K字内</td>
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

  document.fm01.InfCont.value = apiGetValEditID01();
  <%If (PrmFlag="(Mem)" AND ModID="MemB124") OR (PrmFlag="(Inn)" AND ModID="MemB424") Then%>
  <%Else%>
  document.fm01.InfReply.value = apiGetValEditID02();
  <%End If%>
 if (document.fm01.InfCont.value.length>=12000) 
   {   
     alert(" 内容 不能超过 12K!"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }


         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
