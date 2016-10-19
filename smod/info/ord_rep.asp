<!--#include file="config.asp"-->
<!--#include file="../../pfile/lang/vmemb.asp"-->
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
If Session("UsrID")&""="" Then
  Response.End()
End If
mStart = RequestS("mStart",3,24) 
mShow = RequestS("mShow",3,24) 
If mStart="" Then
  d = Date()
  y = DatePart("yyyy",d)
  m = DatePart("m",d) :if m<10 Then m="0"&m
  mStart = y&"-"&m&"-01"
End If
%>

<table width="100%" border="0" cellpadding="2" cellspacing="1">
  <tr>
    <td colspan="2"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>定单统计报表</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="mStart" type="text" id="mStart" value="<%=mStart%>" size="12">
            <input type="submit" name="Submit2" value="<%=vMem_ODA3%>"></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="20%" align="left" valign="top" style="padding-left:5px;"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td align="center" nowrap bgcolor="#E0E0E0">月份</td>
          <td align="right" nowrap bgcolor="#E0E0E0">单数</td>
          <td align="right" nowrap bgcolor="#E0E0E0">总金额</td>
        </tr>
<%
  sn = 0 :  ss = 0
  For i=0 To 11
		If i mod 2 = 0 Then
		  col = "#ffffff"
		Else
		  col = "#F0F0F0"
		End If
	d = DateAdd("m",-1*i,mStart)
	y = DatePart("yyyy",d)
	m = DatePart("m",d) :if m<10 Then m="0"&m
	m = y&"-"&m
	d1 = m&"-1"
	d2 = DateAdd("m",1,d1)
	w = "WHERE LogATime>=#"&d1&"# AND LogATime<=#"&d2&"#"
	n = rs_Count(conn,"OrdInfo "&w)
	s = rs_Val(conn,"SELECT SUM(InfNum) AS mInfSum FROM OrdInfo "&w&" ")
	s = FormatNumber(RequestSafe(s,"N",0),2)
	sn = sn + n
	ss = ss + s
	If i=0 And mShow="" Then
	  mShow = m
	End If
%>
        <tr bgcolor="<%=col%>">
          <td align="center" nowrap>
          <%If mShow=m Then%>
          <span class="colF00"><%=m%></span>
          <%Else%>
          <a href="?mShow=<%=m%>&mStart=<%=mStart%>"><%=m%></a>
          <%End If%>
          </td>
          <td align="right" nowrap><%=n%></td>
          <td align="right" nowrap><%=s%></td>
        </tr>
<%
  Next
  ss = FormatNumber(ss,2)
%>
        <tr bgcolor="#FFFFFF">
          <td align="center" nowrap>合计</td>
          <td align="right" nowrap><%=sn%></td>
          <td align="right" nowrap><%=ss%></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td colspan="3" align="left" style="line-height:150%;">注意：<br>
&nbsp;　 点左边月份，则右边显示这月定单列表；<br>
          &nbsp;　 右上输入年-月(如:2010-12)搜索，则统计从此月开始一年内的定单</td>
        </tr>
    </table></td>
    <td align="left" valign="top" style="padding-right:5px;">
    
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="#E0E0E0">
    <td width="5%" align="center" nowrap>NO</td>
    <td width="18%" align="center"><%=mShow%>(月份)</td>
    <td align="center" nowrap><%=vMem_ODA6%></td>
    <td width="12%" align="center" nowrap><%=vMem_ODA7%></td>
    <td width="12%" align="center" nowrap><%=vMem_ODA8%></td>
    <td width="12%" align="center" nowrap>金额</td>
  </tr>
<%
  d1 = mShow&"-1"
  d2 = DateAdd("m",1,d1)
  w = "WHERE LogATime>=#"&d1&"# AND LogATime<=#"&d2&"#"
  sql = " SELECT * FROM [OrdInfo] "&w
  sql =sql& " ORDER BY KeyID DESC" 'SetTop,
  Set rs=Server.Createobject("Adodb.Recordset")
  rs.Open Sql,conn,1,1
  i = 0
  If not rs.eof then
  Do While NOT rs.EOF
    i = i+1
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F0F0F0"
		End If

KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
InfSubj = rs("InfSubj")

LnkName = rs("LnkName")
LnkSex  = rs("LnkSex")

LogATime = rs("LogATime") :LogATime = Mid(LogATime,1,Len(LogATime)-3) 'FormatDateTime(,2)
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
If verMemb="2" Then
  SetState = Get_State(SetState,"N;D;P;S","New;Doing;Pay;Send")
Else
  SetState = Get_State(SetState,"N;D;P;S","新建;处理中;已付款;已发货")
End If
InfNum = rs("InfNum")
'aSum = rs_Val(conn,"SELECT SUM(InfPrice*InfCount) AS aSum FROM OrdItem WHERE KeyCode='"&KeyID&"'")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      </td>
      <td align="center"><a href="ord_view.asp?ID=<%=KeyID%>&PrmFlag=<%=PrmFlag%>&verMemb=<%=verMemb%>" target="_blank"><%=KeyCode%></a></td>
      <td align="center" nowrap><%=LnkName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetState%></td>
      <td align="center" nowrap><%=InfNum%></td>
    </tr>
    <%
    rs.Movenext
  Loop  
  Else
%>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="10">无信息</td>
    </tr>
<%
  End If
  rs.Close()
  Set rs = Nothing
%>
</table>
    </td>
  </tr>
</table>

</body>
</html>
