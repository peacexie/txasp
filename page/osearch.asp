<!--#include file="_config.asp"-->
<%

If verNow="2" Then

  vOrd_m6Time = "Order Time"
  vOrd_m6Rem = "Please remember this NO."
  vOrd_m6Unit = "RMB"
  
Else

  vOrd_m6Time = "订购时间"
  vOrd_m6Rem = "请记录此定单号"
  vOrd_m6Unit = "元"
  
End If

KW = RequestS("KeyWD",3,255)
KW = Replace(KW,"'",";") : KW = Replace(KW,"‘","")
KW = Replace(KW,",",";") : KW = Replace(KW,"，",";")
KW = Replace(KW," ",";") : KW = Replace(KW,"　",";")
KW = Replace(KW,vbcrlf,";") : KW = Replace(KW,"；",";")
KW = Replace(KW,vbcr,";")   : KW = Replace(KW,vblf,";")
aKW = Split(KW,";") : sqlK = ""
For i=0 To uBound(aKW)
 If Len(aKW(i))>8 Then
  sqlK=sqlK&" OR KeyCode LIKE '%"&aKW(i)&"%' "
 End If
Next

sql = " SELECT * FROM [OrdInfo] WHERE 1=2 "&sqlK&" ORDER BY KeyID DESC" 
'Response.Write sql
rs.Open sql,conn,1,1
    

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Order Search</title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.OrdFLeg {
	padding:2px 5px 2px 5px;
	font-weight:bold;
}
.OrdFSet {
	padding:5px 5px 5px 5px;
	margin:2px 5px;
}
</style>
</head>
<body style='background-color:#FFF; padding:12px 0px'>


<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">


  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">单号</td>
    <td height="27" align="center" nowrap>标题</td>
    <td align="center" nowrap>姓名</td>
    <td width="12%" align="center" nowrap>日期</td>
    <td width="12%" align="center" nowrap>状态</td>
    <td width="12%" height="27" align="center" nowrap>查看</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
  </tr>
  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  i = 0 
  Do While NOT rs.EOF 
    i = i + 1
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If

KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
InfSubj = rs("InfSubj")

LnkName = rs("LnkName")
LnkSex  = rs("LnkSex")

LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfSubj = Show_Text(InfSubj) 

SetState = rs("SetState") : If SetState="-" Then SetState="N"
If PrmFlag<>"(Mem)" Then
  dFlag = true
ElseIf SetState="N" And PrmFlag="(Mem)" Then
  dFlag = true
Else
  dFlag = false
End If
SetState = Get_State(SetState,"N;D;P;S","新建;处理中;已付款;已发货")

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%></td>
      <td align="center"><a href="ord_view.asp?ID=<%=KeyID%>" target="_blank"><%=KeyCode%></a></td>
      <td align="center" nowrap><%=InfSubj%></td>
      <td align="center" nowrap><%=LnkName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetState%></td>
      <td align="center" nowrap>
      <%If LogAUser=Session("MemID") Then%>
      <a href="../smod/info/ord_view.asp?ID=<%=KeyID%>&PrmFlag=(Mem)" target="_blank">查看</a>
      <%Else%>
      <span class="fntCCC">查看</span>
      <%End If%>
      </td>
    </tr>
    <%
  rs.Movenext
  Loop
%>

    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="9">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="9" align="right"></td>
    </tr>
  </form>
</table>

</body>
</html>
