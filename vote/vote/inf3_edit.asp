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
</head>
<body>
<%
DeMax = "1"'RequestSafe(rs_Val("","SELECT ParNum FROM AdmPara WHERE ParCode='n"&ModID&"'"),"N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,24)
PG = RequestS("PG","N",1)
ID = RequestS("ID",3,48)

send = Request("send")
Session("KeyID") = ID

If send = "send" Then

SetUBB = RequestS("SetUBB",3,2)
InfRem1 = RequestS("InfRem1",3,9600) 
InfRem1 = Show_Html(InfRem1)
If InfRem1="" Then InfRem1=" "

sql = " UPDATE [VoteInfo] SET " 
sql = sql& " KeyCode = '" & RequestS("KeyCode",3,255) &"'" 
sql = sql& ",InfType = '" & RequestS("InfType",3,255) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",InfRem1 = '" & InfRem1 &"'" 

sql = sql& ",InfTime1='" & RequestS("InfTime1","D",Date()) &"'" 
sql = sql& ",InfTime2='" & RequestS("InfTime2","D",Date()) &"'" 
sql = sql& ",InfNum1=" & RequestS("InfNum1","N",1) &"" 
sql = sql& ",InfNum2=" & RequestS("InfNum2","N",10) &"" 
sql = sql& ",InfCard='" & RequestS("InfCard","C",2) &"'" 
sql = sql& ",InfVNum='" & RequestS("InfVNum","C",24) &"'" 

sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "
  Call rs_DoSql(conn,sql) 
  'Call rs_GetFile(ID,"")
  If Request("Act")="Close" Then
  Response.Write js_Alert("修改成功!","Close","")
  Else
  Response.Redirect "inf3_list.asp?TP="&TP&"&KW="&KW&"&Page="&PG
  End If
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = Show_Form(rs("InfSubj"))
InfRem1 = rs("InfRem1")&""
InfRem1 = Server.HTMLEncode(InfRem1) 

InfTime1 = rs("InfTime1")
InfTime2 = rs("InfTime2")
InfCard = rs("InfCard")
InfVNum = rs("InfVNum")
InfNum1 = rs("InfNum1")
InfNum2 = rs("InfNum2")

ImgName = rs("ImgName")
If ImgName<>"" then ronly = "readonly"
  else
KeyID = ""  
  end if 
rs.Close()
SET rs=Nothing 

If KeyID = "" Then Response.Redirect("ind_list.asp")

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
      <td align="center" bgcolor="#FFFFFF"><strong>[简易调查]编辑</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120">      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">类别</td>
      <td bgcolor="#FFFFFF"><select name="InfType" id="InfType" style="width:165px;">
        <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay")%>
      </select>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;编号
<input name="KeyCode" type="text" id="KeyCode" value="<%=KeyCode%>" size="24" maxlength="24"></td>
    </tr>

    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">调查说明</td>
      <td bgcolor="#FFFFFF"><textarea name="InfRem1" id="InfRem1" cols="50" rows="6"><%=InfRem1%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">调查日期</td>
      <td bgcolor="#FFFFFF"><input name="InfTime1" type="text" id="InfTime1" value="<%=InfTime1%>" size="14" maxlength="10">
        ~
        <input name="InfTime2" type="text" id="InfTime2" value="<%=InfTime2%>" size="14" maxlength="10">
        &nbsp;(格式:2008-12-31)</td>
    </tr>
    <!--
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">验证规则</td>
      <td bgcolor="#FFFFFF"><select name="InfCard" id="InfCard" style="width:120px;">
          <option value='N' <%If InfCard="N" Then Response.Write("selected")%>>无身份证号验证</option>
          <option value='Y' <%If InfCard="Y" Then Response.Write("selected")%>>有身份证号验证</option>
        </select>
        ~
        <select name="InfVNum" id="InfVNum" style="width:120px;">
          <option value='D1T1' <%If InfVNum="D1T1" Then Response.Write("selected")%>>每IP每天可投1次</option>
          <option value='DNT1' <%If InfVNum="DNT1" Then Response.Write("selected")%>>每IP共可问卷1次</option>
          <option value='DXTN' <%If InfVNum="DXTN" Then Response.Write("selected")%>>[无限制] 问卷</option>
        </select></td>
    </tr>
    -->
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
 
 if (document.fm01.InfRem1.value.length==0) 
   {   
     //alert(" 内容 不能为空！"); 
     //eflag = 1; break;
   }  
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
