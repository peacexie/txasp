<!--#include file="config.asp"-->
<%
Response.ContentType = "application/zip" 'application/vnd.ms-Excel;application/zip
Response.Addheader "Content-Disposition","attachment;Filename=member_"&Get_FmtID("mdhnsx","")&".xls" 
Response.Charset = "UTF-8" 
r = Request("r")
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<style>
<!--
table {
	mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";
}
@page {
margin:1.0in .75in 1.0in .75in;
 mso-header-margin:.5in;
 mso-footer-margin:.5in;
}
tr {
	mso-height-source:auto;
	mso-ruby-visibility:none;
}
col {
	mso-width-source:auto;
	mso-ruby-visibility:none;
}
br {
	mso-data-placement:same-cell;
}
td {
	mso-style-parent:style0;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:宋体;
	mso-generic-font-family:auto;
	mso-font-charset:134;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;
}
ruby {
	ruby-align:left;
}
rt {
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:宋体;
	mso-generic-font-family:auto;
	mso-font-charset:134;
	mso-char-type:none;
	display:none;
}
-->
</style>
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>WebMembers</x:Name>
    <x:WorksheetOptions>
     <x:DefaultRowHeight>285</x:DefaultRowHeight>
     <x:Selected/>
     <x:Panes>
      <x:Pane>
       <x:Number>3</x:Number>
       <x:ActiveRow>31</x:ActiveRow>
       <x:ActiveCol>1</x:ActiveCol>
      </x:Pane>
     </x:Panes>
     <x:ProtectContents>False</x:ProtectContents>
     <x:ProtectObjects>False</x:ProtectObjects>
     <x:ProtectScenarios>False</x:ProtectScenarios>
    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
  <x:WindowHeight>9225</x:WindowHeight>
  <x:WindowWidth>14940</x:WindowWidth>
  <x:WindowTopX>240</x:WindowTopX>
  <x:WindowTopY>105</x:WindowTopY>
  <x:ProtectStructure>False</x:ProtectStructure>
  <x:ProtectWindows>False</x:ProtectWindows>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body link=blue vlink=purple>
<%

   sql = "SELECT * FROM Member"&Mem_aMemb&" WHERE 1=1 ORDER BY LogATime DESC"
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 32000 

%>
<table x:str border=0 cellpadding=0 cellspacing=0 width=936 style='border-collapse:
 collapse;table-layout:fixed;width:702pt'>
  <col width=72 span=13 style='width:54pt'>
  <tr>
    <td>NO</td>
    <td>帐号</td>
    <td>公司名称</td>
    <td>联系人</td>
    <td>性别</td>
    <td>证件号码</td>
    <td>手机</td>
    <td>电话</td>
    <td>邮件</td>
    <td>地址</td>
    <td>生日</td>
    <td>类别</td>
    <td>注册</td>
  </tr>
  <%

  if not rs.eof then
  rs.AbsolutePage = 1
  for i = 1 to rs.PageSize

MemID = rs("MemID")
MemType = rs("MemType")
MemName = Show_Text(rs("MemName"))
MemNam2 = Show_Text(rs("MemNam2"))     : If MemNam2="" Then MemNam2="&nbsp;"
MemSex = rs("MemSex")
MemCard = Show_Text(rs("MemCard"))
MemBirth = rs("MemBirth")

MemMobile = Show_Text(rs("MemMobile")) : If MemMobile="" Then MemMobile="&nbsp;"
MemTel = Show_Text(rs("MemTel"))       : If MemTel="" Then MemTel="&nbsp;"  
MemEmail = Show_Text(rs("MemEmail"))   : If MemEmail="" Then MemEmail="&nbsp;"
MemFrom = Show_Text(rs("MemFrom"))     : If MemFrom="" Then MemFrom="&nbsp;"

LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = Left(rs("LogATime"),13)
if MemSex = "F" then
  MemSex = "女"&MemMarry
elseif MemSex = "M" then
  MemSex = "男"&MemMarry
else
  MemSex = "未知"
end if
If MemBirth="1900-1-1" Or MemBirth="1900-12-31" Then
  MemBirth = "[---]"
ElseIf inStr(MemBirth," ")>0 Then
  MemBirth = FormatDateTime(MemBirth,2)
End If

If MemNam2="&nbsp;" And MemType="Privy" Then
  MemNam2 = MemName
  MemName = "&nbsp;"
End If

	  %>
  <tr>
    <td><%=i%></td>
    <td><%=MemID%></td>
    <td><%=MemName%></td>
    <td><%=MemNam2%></td>
    <td><%=MemSex%></td>
    <td><%=MemCard%></td>
    <td><%=MemMobile%></td>
    <td><%=MemTel%></td>
    <td><%=MemEmail%></td>
    <td><%=MemFrom%></td>
    <td><%=MemBirth%></td>
    <td><%=MemType%></td>
    <td><%=LogATime%></td>
  </tr>
  <%
  rs.movenext
  if rs.eof then exit for
  next

  rs.Close
  end if
  set rs = nothing
%>
</table>
</body>
</html>
