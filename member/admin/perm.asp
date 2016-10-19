<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

ID  = RequestS("ID",3,48)
sql = "SELECT SysName,SysPerm,SysType FROM [MemSyst] WHERE SysID='"&ID&"'"
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
If NOT rs.EOF Then
  sName = rs("SysName")
  sPerm = rs("SysPerm")
  sType = rs("SysType")
End If
rs.close()
set rs = nothing


send  = RequestS("send",3,48) 
TPM = RequestS("TPM",3,24) 
If send="send" Then
  PrmList  = RequestS("PrmList",3,8000)
  PrmArr = Split(PrmList,",")
  s = ""
For i = 0 To uBound(PrmArr)
  If PrmArr(i)&""<>"" Then
    s = s &"("& Trim(PrmArr(i)) &");"
  End If
Next
  s = ""&s '{"&UsrType&"};
  Call rs_DoSql(conn,"UPDATE [MemSyst] SET SysPerm='"&s&"' WHERE SysID='"&ID&"'")
  sPerm = s
  msg = "权限修改成功!!"
End If


sID = ""
sNM = ""
sql = " SELECT * FROM [MemSyst] "
sql = sql& " WHERE SysType='Type' "
sql = sql& " ORDER BY SysTop,SysID " 
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
Do While NOT rs.EOF
  sID = sID&rs("SysID")&";"
  sNM = sNM&rs("SysName")&";"
rs.MoveNext
Loop
rs.close()
set rs = nothing
aID = Split(sID,";")
aNM = Split(sNM,";")
'<input type='checkbox' name='[ChkBox01]' value='"&SysID&"' "&ch1&">"&SysName&"("&SysID&")
sql = " SELECT * FROM [MemSyst] "
sql = sql& " WHERE SysID LIKE 'x%'  " 'AND (SysDef='(Public)' OR SysDef LIKE 'Mod"&Left(ID,3)&"%' )
'sql = sql& " WHERE SysType='System' "
sql = sql& " ORDER BY SysID " 
%>
        <br>
        <table width="620" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
          <tr align="center">
            <td width="30%" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>会员等级权限设定</strong> <br>
            <font color="#0000FF">[<%=ID%>]<%=sName%></font></td>
            <td colspan="2" align="right" bgcolor="#FFFFFF"> | <a href="type.asp">社区</a> | <a href="system.asp">模块</a> | <a href="grade.asp">等级&nbsp;</a> </td>
          </tr>
          <tr align="center">
            <td width="50%" align="center" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
            <td width="50%" align="center" bgcolor="#FFFFFF"><a href="grade.asp?TPU=Grade&TPM=<%=TPM%>">返回会员等级</a></td>
          </tr>
		  <form name="fm01" action="?">
          <tr align="center">
            <td colspan="3" align="center" bgcolor="#FFFFFF">
              <table width="98%" border="0" cellpadding="2" cellspacing="1">
			  <%
			  For i=0 To uBound(aID)-1
			  t = Mid(aID(i),4,3)
			  sql3 = "SELECT SysID,SysName FROM [MemSyst] WHERE SysID LIKE 'x"&t&"%' "
			  s = Get_rsCBox(conn,sql3,sPerm) 
			  s = Replace(s,"[ChkBox01]","PrmList") 
			  s = Replace(s,"<br>","　　")
			  %>
                <tr>
                  <td style="color:#666666;font-weight:bold;">　　<%=aNM(i)%>[<%=aID(i)%>]</td>
                </tr>
                <tr>
                  <td><%=s%></td>
                </tr>
                <tr>
                  <td bgcolor="#CCCCCC"></td>
                </tr>
			  <%
			  Next
			  %>
                <tr>
                  <td align="center"><table width="100%"  border="0">
                    <tr>
                      <td width="50%" align="center"><input name="SetOK" type="submit" id="SetOK" value="确认">
                        <input name="send" type="hidden" id="send" value="send">
                        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
                        <input name="TPM" type="hidden" id="TPM" value="<%=TPM%>">
&nbsp;</td>
                      <td width="50%" align="center"><a href="grade.asp?TPU=Grade&TPM=<%=TPM%>">返回会员等级</a></td>
                    </tr>
                  </table>
                    </td>
                </tr>
            </table></td>
          </tr>
          
		  </form>
</table>
        
        
</body>
</html>
