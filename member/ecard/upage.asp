<!--#include file="../../page/_config.asp"-->
<%
MD="MemC124"
MDName = GetMName(MD)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%></title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--Item Start-->
<style type="text/css">
.tb2Line{ 
  border-top:1px solid #CCC;
  border-bottom:1px solid #CCC;
}
.topLine{ 
  border-top:1px solid #CCC;
}
.botLine{ 
  border-bottom:1px solid #CCC;
}
.topLine{ 
  border-top:1px solid #CCC;
}
</style>
<!--- //////////// -->
<%
CrdNO = RequestS("CrdNO","C",48)
CrdType = RequestS("CrdType","C",48)
CrdName = RequestS("CrdName","C",48)

sql = "SELECT "
sql = sql & " * FROM [MemCard] "
If CrdNO<>"" Then
 sql =sql& " WHERE CrdNO='"&CrdNO&"' "
 'sql =sql& " WHERE CrdNO LIKE '%"&CrdNO&"%' "
End If
If CrdType<>"" Then
 'sql =sql& " AND CrdType='"&CrdType&"' "
End If 
If CrdName<>"" Then
 sql =sql& " AND CrdName='"&CrdName&"' "
End If 
'Response.Write sql

%>
<%If Request("send")<>"send" Then%>
<table width="500" height="162" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="fm01" action="?" xtarget="_blank">
    <tr bgcolor="f4f4f4">
      <td colspan="2" align="left" class="tb2Line"><strong>&nbsp;档案查询</strong></td>
    </tr>
    <tr>
      <td width="120" align="center">&nbsp;</td>
      <td align="left">请输入证书编号(结业证号)</td>
    </tr>
    <tr>
      <td align="right">证书编号</td>
      <td align="left"><input name="CrdNO" type="text" id="CrdNO" style="width:130px; height:18px" maxlength="24"></td>
    </tr>
    <tr>
      <td align="right">客户姓名</td>
      <td align="left"><input name="CrdName" type="text" id="CrdName" style="width:130px; height:18px" maxlength="24"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td align="left"><input type="button" name="Button" value="查询" style="width:130px;" onClick="chkData()"></td>
    </tr>
    <tr>
      <td class="botLine line01" colspan="2" align="left">&nbsp;</td>
    </tr>
    <input name="send" type="hidden" value="send">
  </form>
</table>
<script type="text/javascript">
var fm = document.fm01;
function chkData()
{
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////
   
  if ((fm.CrdNO.value.length==0)&&(fm.CrdName.value.length==0))
  {   
     alert("证书号码,姓名 不能同时为空");
	 fm.CrdNO.focus();
	 eflag = 1; break;
  }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 fm.submit(); 
		 }
}
</script>
<%Else%>
<table width="500" align="center" cellpadding="3" cellspacing="1">
  <tr>
    <td colspan="6" bgcolor="#999999"></td>
  </tr>
  <tr bgcolor="f4f4f4">
    <td colspan="6"><strong>&nbsp;查询结果</strong> <font color="#FF0000"><%=Msg%></font></td>
  </tr>
  <%
  
   rs.Open Sql,conn,1,1
if NOT rs.EOF then
%>
  <tr align="center">
    <td class="tb2Line" nowrap>结业证号</td>
    <td class="tb2Line" nowrap>姓名</td>
    <td class="tb2Line" nowrap>性别</td>
    <td class="tb2Line" nowrap>头衔</td>
    <td class="tb2Line" nowrap>报学课程</td>
    <td class="tb2Line" nowrap>时间</td>
  </tr>
  <%
  i = 1
Do While NOT rs.EOF 
  i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f8f8f8"
		end if
CrdID = rs("CrdID")
CrdNO = rs("CrdNO")
CrdNCode = rs("CrdNCode")
CrdType = rs("CrdType")
CrdName = rs("CrdName")
CrdSpeci = rs("CrdSpeci")
CrdTime = rs("CrdTime")
CrdRem = Show_Form(rs("CrdRem"))
f="Y"


  

  %>
  <tr bgcolor="<%=col%>">
    <td nowrap><%=CrdNO%></td>
    <td nowrap><%=CrdName%></td>
    <td align="center" nowrap><%=CrdNCode%></td>
    <td align="center" nowrap><%=CrdSpeci%></td>
    <td align="center" nowrap><%=CrdType%></td>
    <td align="center" nowrap><%=CrdTime%></td>
  </tr>
  <%
  rs.MoveNext()
Loop
rs.close()
%>
  <%
  Else
  %>
  <tr>
    <td colspan="6" class="tb2Line">很抱歉我无法查询到贵用户数据!</td>
  </tr>
  <%End If%>
  <tr>
    <td colspan="6" class="tb2Line">继续查询 　　 <a href="?" class="cRed">请返回&gt;&gt;&gt;</a></td>
  </tr>
</table>
<%End If%>
<!--Item End-->
<%Set rs=Nothing %>
</body>
</html>
