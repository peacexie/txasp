<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

Set rs=Server.Createobject("Adodb.Recordset")

'OrgTCode = "InfA124;InfN1;InfN2;PicS124"
'OrgTName = "介绍信息;塘厦教育;塘厦教研;资源中心"
OrgTCode = "InfA124;InfD124;InfC124;InfN124;PicS124;PicV124"
OrgTName = "企业介绍;部门介绍;课程介绍;新闻中心;产品图片;视频下载"

sql2 = "SELECT TypID,TypName FROM WebType WHERE TypMod='InfHead' ORDER BY TypTop, TypID "
rs.Open sql2,conn,1,1 
Do While Not rs.EOF 
  s1 = rs("TypID")
  s2 = rs("TypName")
  OrgTCode = OrgTCode&";"&s1
  OrgTName = OrgTName&";"&s2
rs.MoveNext
Loop
rs.Close()

a1 = Split(OrgTCode,";")
a2 = Split(OrgTName,";")

yAct = Request("yAct") 
KW = RequestS("KW",3,24)

If KW&"" <> "" Then
  sqlK = sqlK & " AND ( UsrID LIKE '%"&KW&"%' "
  sqlK = sqlK & " OR UsrName LIKE '%"&KW&"%' "
  sqlK = sqlK & " ) " 
End If

If yAct="xLogATime"Then
  ''//
  Msg = cID&" 条记录 设置成功!"
End If

 sql = " SELECT * FROM [AdmUser"&Adm_aUser&"] "
 sql = sql& " WHERE UsrType LIKE 'Adm%' "&sqlK
 sql = sql& " ORDER BY UsrType, UsrID " 
 rs.Open Sql,conn,1,1

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="150" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[信息发布统计] </strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>

  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="center" nowrap bgcolor="#FFFFFF">NO</td>
    <th nowrap bgcolor="#FFFFFF">帐号</th>
    <th nowrap bgcolor="#FFFFFF">单位/姓名</th>
    <%
	For i=0 To uBound(a1)
	%>
    <td nowrap bgcolor="#FFFFFF"><%=a2(i)%></td>
    <%
	Next
	%>
    <th nowrap bgcolor="#FFFFFF">登陆IP</th>
    <th nowrap bgcolor="#FFFFFF">最后登陆</th>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="150" align="right" nowrap></td>
  </tr>

    <%
  j = 0
  Do While NOT rs.EOF
  j =j+1
		If j mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
		UsrID = rs("UsrID")
		UsrName = rs("UsrName") 
		UsrLogIP = rs("UsrLogIP")
		UsrLTime = rs("UsrLTime")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=j%></td>
      <td><%=UsrID%></td>
      <td align="center" nowrap><%=UsrName%></td>
    <%
	For i=0 To uBound(a1)
	'InfN124 PicS124 HD1212 HD1216 HD1220 HD1224 HD1228 
	If Left(a1(i),2)="HD" Then
	  iTab = "InfoHead"
	  iMod = "InfTyp2"
	ElseIf a1(i)="PicS124" OR a1(i)="PicV124" Then
	  iTab = "InfoPics"
	  iMod = "KeyMod"
	Else
	  iTab = "InfoNews"
	  iMod = "KeyMod"
	End If
	If Len(a1(i))=7 Then 
	  wMod = " "&iMod&"='"&a1(i)&"' "
	Else
	  wMod = " "&iMod&" LIKE '"&a1(i)&"%' "
	  If Left(a1(i),3)="Pic" Then 
		iTab = "InfoPics"
	  Else
		iTab = "InfoNews"
	  End If 
	End If 
	iCnt = rs_Count(conn," "&iTab&" WHERE "&iMod&"='"&a1(i)&"' AND LogAUser='"&UsrID&"' ")
	%>
      <td align="center" nowrap><%=iCnt%></td>
    <%
	Next
	%>
      
      <td align="center" nowrap><%=UsrLogIP%></td>
      <td align="center" nowrap><%=UsrLTime%></td>
    </tr>
    <%
  rs.Movenext
  Loop
%>
    <tr bgcolor="E0E0E0">
      <td height="21" colspan="150" align="left" nowrap bgcolor="#FFFFFF">注意：</td>
    </tr>
    <%  
  rs.Close()
  Set rs = Nothing  
	  %>
    <tr bgcolor="#999999">
      <td colspan="150" align="right"></td>
    </tr>

</table>
<script type="text/javascript">


function fAct()
{
  var frmID = document.flist;
  var vAct = frmID.yAct.value; 
  if(vAct=='SetTop'){
    nAct.innerHTML = "设置值在[100~999]之间";
	frmID.yVal.value = '888';
  }
  if(vAct=='SetSubj'){
    nAct.innerHTML = "设置格式[RRGGBB]";
	frmID.yVal.value = '000000';
  }
  if(vAct=='LogATime'){
    nAct.innerHTML = "设置格式[yyyy-mm-dd HH:mm:ss]";
	frmID.yVal.value = '<%=Now()%>';
  }
  //alert(XAct);
}
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>
</body>
</html>
