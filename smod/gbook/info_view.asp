<!--#include file="config.asp"-->
<%
send = Request("send")
PG = RequestS("PG","N",1)
KW = RequestS("KW",3,48)
ID = RequestS("ID",3,48)
TP = RequestS("TP",3,48)
ModID = RequestS("ModID",3,48)

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [GboInfo] WHERE KeyID='"&ID&"'",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = Show_Form(rs("InfType"))
InfSubj = Show_Form(rs("InfSubj"))
xxxCont = Show_Html(rs("InfCont"))
xxxReply = Show_Html(rs("InfReply"))
LnkName = Show_Form(rs("LnkName"))
LnkEmail = Show_Form(rs("LnkEmail"))
SetRead = rs("SetRead")
SetShow = rs("SetShow")
ImgName = Show_Form(rs("ImgName"))
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
end if 
rs.Close()
SET rs=Nothing 

TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
If TypName="" Then
  TypName = MDName
End If

If KeyMod="GboU324" Then
  If SetShow<>"Y" Then
    InfCont = Show_Text(InfCont)
  End If
End If
%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%></title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <tr bgcolor="#FFFFFF">
    <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
    <td align="right" bgcolor="#FFFFFF"><strong>信息增加</strong></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
    <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center" bgcolor="#FFFFFF">姓名</td>
    <td bgcolor="#FFFFFF"><input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="18" maxlength="24" style="width:90 ">
&nbsp;邮箱
<input name="LnkEmail" type="text" id="LnkEmail" value="<%=LnkEmail%>" size="18" maxlength="24" style="width:90 ">
&nbsp;审核
<select name="SetShow" id="SetShow">
  <option value="Y"  <%if SetShow="Y" then Response.Write("selected")%>>已审核</option>
  <option value="N"  <%if SetShow="N" then Response.Write("selected")%>>未审核</option>
</select>
<%=TypName%></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center" bgcolor="#FFFFFF">内容</td>
    <td bgcolor="#FFFFFF"><table  border="0" cellpadding="3" cellspacing="5" bgcolor="#E0E0E0">
      <tr>
        <td width="540" height="120" valign="top" bgcolor="#FFFFFF"><%Call Show_sfGbook(ID,".org.htm")%></td>
      </tr>
    </table>
    </td>
  </tr>
  <%If KeyMod<>"GboU324" Then%>
  <tr bgcolor="#FFFFFF">
    <td align="center" bgcolor="#FFFFFF">回复</td>
    <td bgcolor="#FFFFFF"><table  border="0" cellpadding="3" cellspacing="5" bgcolor="#E0E0E0">
        <tr>
          <td width="540" height="120" valign="top" bgcolor="#FFFFFF"><%Call Show_sfGbook(ID,".rep.htm")%></td>
        </tr>
    </table></td>
  </tr>
  <%End If%>
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

 if (document.fm01.InfCont.value.length>=1800) 
   {   
     alert(" 内容 不能超过 1800字!"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }


         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
