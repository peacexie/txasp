<!--#include file="config.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="mconfig.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/WinFunc.js" type="text/javascript"></script>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

send = Request("send") 
ID = RequestS("ID","C",128)
PW = RequestS("PW","C",1200)
If PW<>"" Then
 Response.Write Rsa_pDec(PW,ID)
 Response.End()
End If

MemGrade = RequestS("MemGrade",C,255)
MemGrade = Replace(MemGrade," ","")

If send = "upd" Then
sql = " UPDATE [Member"&Mem_aMemb&"] SET " 
sql = sql& " MemGrade = '" & MemGrade &"'" 
sql = sql& ",MemFlag = '" & RequestS("MemFlag",C,12) &"'" 
sql = sql& ",LogAUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogEUser = '" & Now() &"'" 
sql = sql& " WHERE MemID='"& ID &"' "
Call rs_DoSql(conn,sql)
'Response.Write sql
End If

sql = "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&ID&"' "
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
if NOT rs.EOF then
MemID = rs("MemID")
MemPW = rs("MemPW")
EncPW = " &nbsp; <a href='?PW="&MemPW&"&ID="&ID&"'>解密</a>" 'Rsa_Dec(MemPW,MemID)' Rsa_pDec
MemQu = rs("MemQu")
MemAn = rs("MemAn")
MemType = rs("MemType")
MemTyp2 = rs("MemTyp2")
MemGrade = rs("MemGrade")
MemName = rs("MemName")
MemNam2 = rs("MemNam2")
MemSex = rs("MemSex")
MemCard = rs("MemCard")
MemBirth = rs("MemBirth")
MemMobile = rs("MemMobile")
MemTel = rs("MemTel")
MemEmail = rs("MemEmail")
MemFrom = rs("MemFrom")
MemExp = rs("MemExp")
MemFlag = rs("MemFlag")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
LogEditIP = rs("LogEditIP")
LogEUser = rs("LogEUser")
LogETime = rs("LogETime")
else
Response.Write "Error!"
end if
  rs.close()
  set rs = nothing

sql = " SELECT * FROM [MemSyst] WHERE SysType='Type' ORDER BY SysID" 
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
s = ""
Do While NOT rs.EOF
  SysID = rs("SysID")
  SysName = rs("SysName")
	sqli = "SELECT SysID,SysName FROM MemSyst WHERE SysType='Grade' AND SysID LIKE 'g"&Mid(SysID,4,3)&"%'"
	OptStr = Get_rsOpt2(conn,sqli,MemGrade)
	OptStr = "<select name='MemGrade' id='MemGrade' style='width:180px;'>"&OptStr&"</select>"
	s = s&vbcrlf&OptStr&"<br>"
rs.movenext
Loop
rs.close()
set rs = nothing

%>
<br>
<%
If Request("Act")="View" Then
MemTName = Get_SOpt(mCfgCode,mCfgName,MemType,"Val")
%>
<table width="540" align="center" cellpadding="2" cellspacing="1" bgcolor="CCCCCC">

    <tr bgcolor="#FFFFFF">
      <td colspan="2" align="left" nowrap bgcolor="#FFFFFF"><strong>&nbsp;会员类型：</strong> <%=MemTName%></td>
    </tr>
    <tr bgcolor="#FFFFFF">

      <td align="right" nowrap bgcolor="#FFFFFF"><FONT color=#ff0000>*</FONT>用 户 名:</td>
      <td nowrap bgcolor="#FFFFFF"><input name='MemID3' type='text' value="<%=MemID%>" size='24' maxlength='64' disabled></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">用户类别:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><%=Get_SOpt(mCfgCode,mCfgName,MemType,"Val")%> (在“修改密码”中修改)</td>
    </tr>
    <%If MemType="Corp" Then%>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">公司名称:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName2' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">行业类型:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><select name="MemTyp" id="MemTyp" style="width:150px;">
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='Fields'",MemTyp2)%>
      </select></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">营业执照:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard2' type='text' id="MemCard2" value="<%=MemCard%>" size='24' maxlength='24'>
        &nbsp;</td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">成立日期:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth2' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
        (如:1900-12-31)</td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">联 系 人:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam' type='text' id="MemNam" value="<%=MemNam2%>" size='24' maxlength='12'>
        <select name="MemSex2" id="MemSex2" style="width:80 ">
          <%=Get_SOpt("F;M;X","女;男;保密",MemSex,"")%>
        </select></td>
    </tr>
    <%If 1=2 Then%>
    <tr>
      <td colspan="2" bgcolor="#FFFFEE">个人</td>
    </tr>
    <%End If%>
    <%ElseIf MemType="Privy" Then%>
    <input type='hidden' name='MemNam' id="MemNam" size='24' value="<%=MemNam2%>">
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">姓&nbsp;&nbsp;&nbsp;名:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName2' type='text' id="MemName" value="<%=MemName%>" size='24' maxlength='24'>
        <select name="MemSex2" id="MemSex2" style="width:80 ">
          <%=Get_SOpt("F;M;X","女;男;保密",MemSex,"")%>
        </select></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">身份证号:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard2' type='text' id="MemCard2" value="<%=MemCard%>" size='24' maxlength='24'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">出生日期:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth2' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
        (如:1900-12-31)</td>
    </tr>
    <%If 1=2 Then%>
    <tr>
      <td colspan="2" bgcolor="#FFFFEE">机构</td>
    </tr>
    <%End If%>
    <%ElseIf MemType="Gov" Then%>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">机构名称:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName2' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">相关证件:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard2' type='text' id="MemCard2" value="<%=MemCard%>" size='24' maxlength='24'>
        &nbsp;如果无资料，填个人身份证号</td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">联 系 人:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam' type='text' id="MemNam" value="<%=MemNam2%>" size='24' maxlength='12'>
        <select name="MemSex2" id="MemSex2" style="width:80 ">
          <%=Get_SOpt("F;M;X","女;男;保密",MemSex,"")%>
        </select></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">出生日期:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth2' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
        (如:1900-12-31)</td>
    </tr>
    <%If 1=2 Then%>
    <tr>
      <td colspan="2" bgcolor="#FFFFEE">团体</td>
    </tr>
    <%End If%>
    <%ElseIf MemType="Org" Then%>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">团体名称:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemName2' type='text' id="MemName" value="<%=MemName%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">相关证件:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemCard2' type='text' id="MemCard2" value="<%=MemCard%>" size='24' maxlength='24'>
        &nbsp;如果无资料，填个人身份证号</td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">成立日期:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemBirth2' type='text' value="<%=MemBirth%>" size='24' maxlength='10'>
        (如:1900-12-31)</td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">联 系 人:</td>
      <td align="left" nowrap bgcolor="#FFFFFF"><input name='MemNam' type='text' id="MemNam" value="<%=MemNam2%>" size='24' maxlength='12'>
        <select name="MemSex2" id="MemSex2" style="width:80 ">
          <%=Get_SOpt("F;M;X","女;男;保密",MemSex,"")%>
        </select></td>
    </tr>
    <%End If%>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">联系电话:</td>
      <td nowrap bgcolor="#FFFFFF"><input name='MemTel2' type='text' id="MemTel2" value="<%=MemTel%>" size='24' maxlength='24'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">手 &nbsp;&nbsp;&nbsp;机:</td>
      <td nowrap bgcolor="#FFFFFF"><input name='MemMobile2' type='text' id="MemMobile2" value="<%=MemMobile%>" size='24' maxlength='24'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">电子邮件:</td>
      <td nowrap bgcolor="#FFFFFF"><input name='MemEmail2' type='text' value="<%=MemEmail%>" size='24' maxlength='120'></td>
    </tr>
    <tr>
      <td align="right" nowrap bgcolor="#FFFFFF">联系地址:</td>
      <td nowrap bgcolor="#FFFFFF"><input name='MemFrom' type='text' id="MemFrom" value="<%=MemFrom%>" size='24' maxlength='120'></td>
    </tr>


</table>
<%
Else
%>
<table width="540" align="center" cellpadding="2" cellspacing="1">
  <form name='fm01' method='post' action='?'>
    <tr>
      <td colspan="2" bgcolor="#999999"></td>
    </tr>
    <tr bgcolor="f4f4f4">
      <td height="24" colspan="2"><strong>&nbsp;会员设定</strong></td>
    </tr>
    <tr>
      <td>MemID</td>
      <td><input name='MemID' type='text' value="<%=MemID%>" size='24' maxlength='64' readonly disabled></td>
    </tr>
    <tr>
      <td>MemName</td>
      <td><input name='MemName' type='text' value="<%=MemName%>" size='24' maxlength='24' readonly disabled></td>
    </tr>
    <tr>
      <td>MemCard</td>
      <td><input name='MemCard' type='text' id="MemCard" value="<%=MemCard%>" size='24' maxlength='48' readonly disabled></td>
    </tr>
    <tr>
      <td>MemTel</td>
      <td><input name='MemTel' type='text' value="<%=MemTel%>" size='24' maxlength='48' readonly disabled></td>
    </tr>
    <tr>
      <td>MemMobile</td>
      <td><input name='MemMobile' type='text' value="<%=MemMobile%>" size='24' maxlength='48' readonly disabled></td>
    </tr>
    <tr>
      <td>MemEmail</td>
      <td><input name='MemEmail' type='text' value="<%=MemEmail%>" size='24' maxlength='48' readonly disabled></td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#999999"></td>
    </tr>
    <tr>
      <td>登陆</td>
      <td><select name="MemFlag" id="MemFlag" style="width:120px; ">
          <option value="-">[-]默认</option>
          <option value="N" <%If MemFlag="N" Then Response.Write("selected") End If%>>冻结</option>
          <option value="Y" <%If MemFlag="Y" Then Response.Write("selected") End If%>>通过</option>
        </select></td>
    </tr>
    <tr>
      <td>会员社区/等级</td>
      <td><%=s%></td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#999999"></td>
    </tr>
    <tr>
      <td>EditUser</td>
      <td><input name='LogAUser' type='text' id="LogAUser" value="<%=LogAUser%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr>
      <td>EditTime</td>
      <td><input name='LogEUser' type='text' id="LogEUser" value="<%=LogEUser%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr>
      <td>LogAddIP</td>
      <td><input name='LogAddIP' type='text' id="LogAddIP" value="<%=LogAddIP%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr>
      <td>LogATime</td>
      <td><input name='LogATime' type='text' id="LogATime" value="<%=LogATime%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr>
      <td>LogEditIP</td>
      <td><input name='LogEditIP' type='text' id="LogEditIP" value="<%=LogEditIP%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr>
      <td>LogETime</td>
      <td><input name='LogETime' type='text' id="LogETime" value="<%=LogETime%>" size='18' maxlength='8' disabled></td>
    </tr>
    <tr bgcolor="f4f4f4">
      <td><%=EncPW%></td>
      <td bgcolor="f4f4f4"><input type='submit' name='Submit' value='提交'>
        &nbsp;&nbsp; &nbsp;
        <input type='reset' name='Reset' value='重设'>
        <input name='send' type='hidden' id='send' value='upd'>
        <input name='ID' type='hidden' id='ID' value='<%=ID%>'></td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#999999"></td>
    </tr>
  </form>
</table>
<%
End If
%>
</body>
</html>