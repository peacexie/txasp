<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.pItem {
	width:120px;
	height:auto;
	margin:0px;
	padding:0px 2px 0px 2px;
	float:left;
	overflow:hidden;
	border-right:1px dashed #EEEEEE;
}
.pItm2 {
	width:20px;
	height:auto;
	margin:0px;
	float:left;
}
.pIRow1 {
	border-top:1px dashed #333333;
}
.pIRow2 {
	padding-bottom:5px;
}
.unLine {
    text-decoration:underline;	
}
.ovLine {
	text-decoration:overline;
}
-->
</style>
</head>
<body>

<%

PrmUser = RequestS("PrmUser",3,48)
ModID = RequestS("ModID",3,48)
sql = "SELECT UsrName,UsrPerm,UsrType FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&PrmUser&"'"
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
If NOT rs.EOF Then
  UsrName = rs("UsrName")
  UsrPerm = rs("UsrPerm")
  UsrType = rs("UsrType")
  sPerm = UsrPerm
End If
rs.close()
set rs = nothing


send  = RequestS("send",3,48) 
If send="send" Then
  PrmList  = RequestS("PrmList",3,8000)
  PrmArr = Split(PrmList,",")
  s = ""
For i = 0 To uBound(PrmArr)
  If PrmArr(i)&""<>"" Then
    s = s &"("& Trim(PrmArr(i)) &");"
  End If
Next
  s = "{"&UsrType&"};"&s
  Call rs_DoSql(conn,"UPDATE [AdmUser"&Adm_aUser&"] SET UsrPerm='"&s&"' WHERE UsrID='"&PrmUser&"'")
  sPerm = s
  msg = "权限修改成功!!"
End If


s1 = ""
s2 = ""
s = ""

sql = " SELECT * FROM [AdmSyst] WHERE SysType='Module' AND SysTop<'U' ORDER BY SysTop,SysID " 
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
  i=i+1
  SysID = rs("SysID")
  SysName = rs("SysName")
    ch1 = ""
  If inStr(sPerm,SysID)>0 Then
    ch1 = "checked"
  End If
  s=s& "<tr><td class='pIRow1'><font style='color:#6666FF'> * "&i&". <input type='checkbox' name='[ChkBox01]' value='"&SysID&"' "&ch1&">"&SysName&" ----- ["&SysID&"]</font></td></tr>"
  s=s& CPart(SysID)
rs.MoveNext
Loop
s = Replace(s,"[ChkBox01]","PrmList")
s = s1&s&s2
   
   
Function CPart(xType)
Dim sql,s,st,mdStr,i,j,p1,s1,s2
sql = "SELECT SysID,SysName FROM AdmSyst WHERE SysType='"&Mid(xType,4)&"' AND SysTop<'u' ORDER BY SysTop,SysID "
sCBox = Get_rsCBox(conn,Sql,sPerm) 
sCArr = Split(sCBox,"<br>")
s="" :st="" :mdStr=""'"InfN124,PicS124" '
For i = 0 To uBound(sCArr)
  If sCArr(i)&""<>"" Then
   '' <input type='checkbox' name='PrmList' value='GboK124'>
     p1 = inStr(sCArr(i),"value='")+7 : s1 = Mid(sCArr(i),p1)
   If inStr(s1,"'")>0 Then
	 p1 = inStr(s1,"'")-1            : s1 = Mid(s1&"",1,p1)
   Else
	 p1 = inStr(s1,"'")-1            : s1 = Mid(s1&"",1,p1)
   End If
   '' Response.Write " "&s1 '' GboK124
   If inStr(mdStr,s1)>0 Then
	s = s &"<span class='pItem'>"& sCArr(i) &"</span>"
	s2 = "<span class='pItem col666'>"&sCArr(i)&"&gt;</span>"
	st = st& "<tr><td class='pIRow2'>"&Replace(s2,"' value='","x' disabled xvalue='")&CPartT(s1)&"</tr>"
   Else
    s = s &"<span class='pItem'>"& sCArr(i) &"</span>"
   End If
  End If
Next
If s<>"" Then
  s = "<tr><td class='pIRow2'>"&s&"</tr>"
End If
CPart = s&st
End Function

Function CPartD()
Dim sql,s,i,j
sql = "SELECT SysID,SysName FROM AdmSyst WHERE SysType='Depart' ORDER BY SysTop,SysID "
sCBox = Get_rsCBox(conn,Sql,sPerm) 
sCArr = Split(sCBox,"<br>")
s = ""
For i = 0 To uBound(sCArr)
  If sCArr(i)&""<>"" Then
    s = s & "<span class='pItem'>&nbsp;"&sCArr(i)&"</span>"
  End If
Next
s = Replace(s,"[ChkBox01]","PrmList")
CPartD = s
End Function

Function CPartT(xType)
Dim sql,s,i,j
sql = "SELECT TypID,TypName FROM WebTyps WHERE TypDeep=1 AND TypMod='"&xType&"' ORDER BY TypID "
sCBox = Get_rsCBox(conn,Sql,sPerm) 
sCArr = Split(sCBox,"<br>")
j = 0 
s = ""
For i = 0 To uBound(sCArr)
 If sCArr(i)&""<>"" Then
  s = s & "<span class='pItem unLine'>"&sCArr(i)&"</span>"
 End If
Next
CPartT = s
End Function

%>
<table width="620" border="0" align="center" cellpadding="3" cellspacing="1">
  <tr align="center">
    <td width="30%" rowspan="2" align="center"><strong>管理者权限设定</strong> <br>
      <font color="#0000FF">[<%=PrmUser%>]<%=UsrName%></font></td>
    <td align="center"><font color="#FF0000"><%=msg%></font></td>
    <td width="20%"><a href="user.asp?ModID=<%=ModID%>">返回用户列表</a></td>
  </tr>
  <tr align="center">
    <td colspan="2" align="right">
<a xhref="user_perm1.asp?PrmUser=<%=PrmUser%>&ModID=Admin">Ver1</a> 
<a href="user_perm2.asp?PrmUser=<%=PrmUser%>&ModID=Admin">Ver2</a> 
<a xhref="user_perm3.asp?PrmUser=<%=PrmUser%>&ModID=Admin">Ver3</a>
    </td>
  </tr>
  <form name="fm01" action="?">
    <tr align="center">
      <td colspan="3" align="center"><table width="640" border="1">
          <tr>
            <td><table width='100%' border=0 align='center'>
                <%=s%>
                <%If SwhDepSubs="Y" Then %>
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * D.
                    <input name='xxxList' type='checkbox' id="xxxList" value='xxxSystem' disabled>
                    科室部门 ----- [SysDepart]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'><%=CPartD()%></td>
                </tr>
                <%End If%>
                <!--InsHere-->
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * X.
                    <input name='xxxList' type='checkbox' id="xxxList" value='xxxSystem' disabled>
                    附加权限 ----- [SysAppend]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'>
                    <input type='checkbox' name='PrmList' value='SetShow' <%If inStr(sPerm,"SetShow")>0 Then Response.Write("checked")%>>
                    发布信息不需要审核(默认都需要审核)
                    &nbsp;
                    <input type='checkbox' name='PrmList' value='SetOher' <%If inStr(sPerm,"SetOher")>0 Then Response.Write("checked")%>>
                    管理查看他人信息(默认只管理查看个人的信息) 
                  </td>
                </tr>
                <!--Ins End-->
                <!--TestHere--/>
                <tr>
                  <td class='pIRow2'><%=CPartT("PicS124")%></td>
                </tr>
                <!--Test End-->
            </table></td>
          </tr>
          <tr>
            <td align="center"><input name="SetOK" type="submit" id="SetOK" value="确认">
              <input name="send" type="hidden" id="send" value="send">
              <input name="PrmUser" type="hidden" id="PrmUser" value="<%=PrmUser%>">
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
</body>
</html>
