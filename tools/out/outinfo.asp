<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../himg/tconfig.asp"-->
<!--#include file="outfunc.asp"-->
<%
Call Chk_Perm1("SysTools","")
conn = conDB 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>外部信息收集</title>
<style type="text/css">
body, td, th {
	font-size: 12px;
}
select {
	width:100px;
}
.DivTest {
	width:180px;
	height:18px;
	display:block;
	padding:3px;
	float:left;
	border:1px solid #CCCCCC;
	margin:1px;
}
body {
	margin: 5px;
}
</style>
</head>
<body>
<%

'// 接收值
yUrl=Request("yUrl") 
yCSet=Request("yCSet")
yID=Request("yID")
YName=Request("YName")
yAct=Request("yAct")
ySave=Request("ySave")
chPos1=Request("chPos1")
chPos2=Request("chPos2")
chNum1=Request("chNum1")
chNum2=Request("chNum2")
chOrg=Request("chOrg")
chObj=Request("chObj")
chKil=Request("chKil")

If yAct="View" Then
 SET rs=Server.CreateObject("Adodb.Recordset") 
 rs.Open "SELECT * FROM [WebAdvert] WHERE KeyID='"&yID&"'",conn,1,1 
 if NOT rs.eof then 
 KeyID = rs("KeyID")
 KeyMod = rs("KeyMod")
 InfType = rs("InfType")
 InfName = rs("InfName")
 InfCont = rs("InfCont")
 InfPath = rs("InfPath")
 InfPara = rs("InfPara")
 end if 
 rs.Close()
 SET rs=Nothing 
 InfContA = Split(InfCont,"^|")
 InfParaA = Split(InfPara,"^|")
 yCSet=InfParaA(0)
 chNum1=InfParaA(1)
 chNum2=InfParaA(2)
 yUrl=InfPath 
 YName=InfName
 yAct=InfType
 ySave="SCode"
 chPos1=InfContA(0)
 chPos2=InfContA(1)
 chOrg=InfContA(2)
 chObj=InfContA(3)
 chKil=InfContA(4)
ElseIf yAct="Del" Then
 yAct=InfType
 Call rs_DoSql(conn,"DELETE FROM WebAdvert WHERE KeyID='"&yID&"'")
 ySave="SCode"
End If


'// 处理默认值
If yUrl="" Then yUrl="http://www.hao123.com/"
If yCSet="" Then yCSet="gb2312" 
If chPos1="" Then chPos1="<body"
If chPos2="" Then chPos2="</body>"
If chNum1="" Then chNum1="0"
If chNum2="" Then chNum2="0"
If yID="" Then yID="Out_"&Replace(Time(),":","-")&".htm"
If YName="" Then YName=yID
If yAct="" Then yAct="LText" 
If ySave="" Then ySave="SCode" 


'// 处理 显示内容
'yUrl = "http://shop.dg.gd.cn/"
'yCSet = "UTF-8"
sHtml=OutPage(yUrl,yCSet)&""
'Response.Write sHtml 

p1=OutSPos(sHtml,chPos1) : p1=p1+Int(chNum1)
p2=OutSPos(sHtml,chPos2) : p2=p2+Int(chNum2)

If p1>0 And p2-p1>0 Then
  sHtml=Mid(sHtml,p1,p2-p1)
Else
  sHtml="(Error)"
End If

If chKil<>"" Then
aKil = Split(chKil,"|")
For i=0 To uBound(aKil)
 If aKil(i)<>"" Then 
  sHtml=Replace(sHtml,aKil(i),"")
 End If 
Next
End If

If chOrg<>"" Then
aOrg = Split(chOrg,"|")
aObj = Split(chObj,"|")
For i=0 To uBound(aOrg)
 If aOrg(i)<>"" Then 
  sHtml=Replace(sHtml,aOrg(i),aObj(i))
 End If 
Next
End If



' // 取内容
If yAct="Page" Then
  '
ElseIf yAct="PText" Then
  Set rExp = New RegExp                ' 建立正则表达式。
    rExp.Global = True                 ' 设置为全局
    rExp.Pattern = "<.*?>|<[^<]*>|<!--[^<]*-->"              ' 设置模式。
    rExp.IgnoreCase = true             ' 设置是否区分大小写。
    sHtml = rExp.Replace(sHtml,"") ' 作替换。
  Set regEx = Nothing
ElseIf yAct="Links" Or yAct="LPics" Or yAct="LText" Then
  ExPar = "<a[^>]*(href=[^>]*)[^>]*>([^<]*|<img[^<]*>)</a>"
  If yAct="LPics" Then ExPar = "<a[^>]*(href=[^>]*)[^>]*>(<img[^<]*>)</a>"
  If yAct="LText" Then ExPar = "<a[^>]*(href=[^>]*)[^>]*>([^<]*)</a>"
  sHtml=OutLinks(sHtml,ExPar)
Else
  '
End If



' // 内容显示,处理：
sBak=sHtml
If ySave="Show" Then
  'Attribute
ElseIf ySave="SCode" Then
  sHtml = ShowText(sHtml)
ElseIf ySave="SLCss" Then
  sHtml=Replace(sHtml,"||","")
ElseIf ySave="SavDB" Then

  InfCont = chPos1&"^|"&chPos2&"^|"&chOrg&"^|"&chObj&"^|"&chKil
  InfPara = yCSet&"^|"&chNum1&"^|"&chNum2

  sqlE = "SELECT * FROM [WebAdvert] WHERE KeyID='"&yID&"' "
  fExist = rs_Exist(conn,sqlE)
  If fExist = "EOF" Then
  sql = " INSERT INTO [WebAdvert] (" 
  sql = sql& "  KeyID,KeyMod" 
  sql = sql& ", InfType, InfName, InfCont" 
  sql = sql& ", InfPath, InfPara" 
  sql = sql& ", LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & yID &"','OutLink'" 
  sql = sql& ", '" & yAct &"'" 
  sql = sql& ", '" & YName &"'" 
  sql = sql& ", '" & Replace(InfCont,"'","''") &"'"  
  sql = sql& ", '" & yUrl &"'" 
  sql = sql& ", '" & Replace(InfPara,"'","''") &"'" 
  sql = sql& ", '" & Get_CIP() &"'" 
  sql = sql& ", '" & Session("UsrID") &"'" 
  sql = sql& ", '" & Now() &"'" 
  sql = sql& ")" ':Response.Write sql
  Else
  sql = " UPDATE [WebAdvert] SET " 
  sql = sql& " InfType = '" & yAct &"'" 
  sql = sql& ",InfName = '" & YName &"'" 
  sql = sql& ",InfCont = '" & Replace(InfCont,"'","''") &"'" 
  sql = sql& ",InfPath = '" & yUrl &"'" 
  sql = sql& ",InfPara = '" & Replace(InfPara,"'","''") &"'" 
  sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
  sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
  sql = sql& ",LogETime = '" & Now() &"'" 
  sql = sql& " WHERE KeyID='"&yID&"' " ':Response.Write sql
  End If
  Call rs_DoSql(conn,sql) 
ElseIf ySave="SavFL" Then
  Call OutFAdd(Server.MapPath(Config_Path&"upfile/#dbf#/"&yID&""),sHtml,"utf-8")
Else
  sHtml = ShowText(sHtml)
End If
If Left(yAct,1)="L" Then
  sHtml=Replace(sHtml,"||",vbcrlf&"<br>")
End If

%>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <th width="20%" bgcolor="#F0F0F0">ID</th>
    <th width="15%" bgcolor="#F0F0F0">Name</th>
    <th bgcolor="#F0F0F0">Url</th>
    <th width="10%" bgcolor="#F0F0F0">Type</th>
    <th width="10%" bgcolor="#F0F0F0">Para</th>
    <th width="5%" bgcolor="#F0F0F0">Edit</th>
    <th width="5%" bgcolor="#F0F0F0">Del</th>
  </tr>
  <%
 sql = " SELECT * FROM [WebAdvert] WHERE KeyMod='OutLink' "
 sql =sql& " ORDER BY KeyID DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfType = rs("InfType")
InfName = rs("InfName")
InfCont = rs("InfCont")
InfPara = rs("InfPara")
InfPath = rs("InfPath")
%>
  <tr>
    <td bgcolor="#FFFFFF"><%=KeyID%></td>
    <td bgcolor="#FFFFFF"><%=InfName%></td>
    <td bgcolor="#FFFFFF"><%=InfPath%></td>
    <td align="center" bgcolor="#FFFFFF"><%=InfType%></td>
    <td align="center" bgcolor="#FFFFFF"><%=InfPara%></td>
    <td align="center" bgcolor="#FFFFFF"><a href="?yID=<%=KeyID%>&yAct=View">Edit</a></td>
    <td align="center" bgcolor="#FFFFFF"><a href="?yID=<%=KeyID%>&yAct=Del">Del</a></td>
  </tr>
  <%
 rs.MoveNext()
 Loop
%>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <form id="fm01" name="fm01" method="post" action="?">
    <tr>
      <td width="20%" align="center" bgcolor="#FFFFFF">Url: </td>
      <td bgcolor="#FFFFFF"><input name="yUrl" type="text" id="yUrl" value="<%=yUrl%>" size="36" maxlength="120" />
        <select name="yCSet" id="yCSet">
          <option value="<%=yCSet%>">.<%=yCSet%>.</option>
          <option value="gb2312" >gb2312</option>
          <option value="UTF-8"  >UTF-8</option>
          <option value="big5"   >big5</option>
        </select></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Pos1:</td>
      <td bgcolor="#FFFFFF"><input name="chPos1" type="text" id="chPos1" value="<%=chPos1%>" size="36" maxlength="30" />
        <input name="chNum1" type="text" id="chNum1" value="<%=chNum1%>" size="6" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Pos2:</td>
      <td bgcolor="#FFFFFF"><input name="chPos2" type="text" id="chPos2" value="<%=chPos2%>" size="36" maxlength="30" />
        <input name="chNum2" type="text" id="chNum2" value="<%=chNum2%>" size="6" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Replace.:</td>
      <td nowrap="nowrap" bgcolor="#FFFFFF"><input name="chOrg" type="text" id="chOrg" value="<%=chOrg%>" size="48" maxlength="50" />
        <br />
        <input name="chObj" type="text" id="chObj" value="<%=chObj%>" size="48" maxlength="50" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Kill Strs:</td>
      <td bgcolor="#FFFFFF"><input name="chKil" type="text" id="chKil" value="<%=chKil%>" size="48" maxlength="50" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">ID / File:</td>
      <td bgcolor="#FFFFFF"><input name="yID" type="text" id="yID" value="<%=yID%>" size="36" maxlength="24" />
        <select name="yAct" id="yAct">
          <option value="<%=yAct%>">.<%=yAct%>.</option>
          <option value="Page"  >Page--页面效果</option>
          <option value="PText" >PText-页面文本</option>
          <option value="Links" >Links-所有连接</option>
          <option value="LPics" >LPics-图片连接</option>
          <option value="LText" >LText-文字连接</option>
        </select>
        <a href="?">重载本页</a></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">Name:</td>
      <td bgcolor="#FFFFFF"><input name="yName" type="text" id="yName" value="<%=yName%>" size="36" maxlength="24" />
        <select name="ySave" id="ySave">
          <option value="<%=ySave%>">.<%=ySave%>.</option>
          <option value="Show" >Show--显示网页</option>
          <option value="SCode">SCode-显示代码</option>
          <option value="SLCss">SLCss-样式显示</option>
          <option value="SavDB">SavDB-存数据库</option>
          <option value="SavFL">SavFL-存入文件</option>
          <option value=""     >-无！</option>
        </select>
        <input type="submit" name="button" id="button" value="提交" />
        &nbsp;</td>
    </tr>
  </form>
</table>
<hr />
<%=sHtml%>
</body>
</html>
