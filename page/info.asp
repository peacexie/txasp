<!--#include file="_config.asp"-->
<%
MD = RequestS("ModID",3,240) : If MD="" Then MD="PicS124" 
TP = RequestS("TypID",3,240) 
If TP="" AND inStr("(InfA124)",MD)>0 Then 
  TP = GetFirstTypeID(MD,"",1)
End If

KW = RequestS("KeyWD",3,24)
Page = RequestS("Page","N",1)
ID = RequestS("KeyID",3,48)

MDName = GetMName(MD)
ModTab = rel_ModTab(MD)

If inStr("(InfA124,PicS124,InfN124)",MD)>0 Then
  If TP="" Then
    TPName = MDName&vPMsg_TName
  Else
	TPName = GetTName(MD,TP,"")
  End If
  sTitle = MDName
  sTPMod = MDName
  sTPOne = TPName 
  sTPLay = TPName '"<a href='info.asp?ModID="&MD&"&TypID="&TP&"'>"&TPName&"</a>"
Else
  TPName = GetTName(MD,TP,"")
  sTitle = MDName
  sTPMod = "关于我们"
  sTPOne = MDName
  sTPLay = ""'MDName
  'sTPLay = " <a href='info.asp?ModID="&MD&"'>"&MDName&"</a> &gt;&gt; <a href='info.asp?ModID="&MD&"&TypID="&TP&"'>"&TPName&"</a>"
End If

sqlK = " WHERE ( KeyMod='"&MD&"' ) " 
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" Then
  sqlK = sqlK & " AND ( InfType LIKE '%"&TP&"%' ) " 
End If

tmpList = get_TmpID("List",MD&";"&TP) 'Cont,Next,FAQ,NList,PList,News,Pics,Jobs,Vdos 
'Response.Write sqlK&ModTab&tmpList&TypMD

'本节Code 与 以下 include file "../pfile/pinc/info_code.asp" 只用一种

%>
<!--#include file="../pfile/pinc/info_code.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-<%=vPMsg_WName%></title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="_head.asp"-->
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle"><%=MDName%></td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <%If MD="PicS124" Then%>
          <td height="30" align="left" class="SysTit02">测试P: <a href="?ModID=<%=MD%>&amp;TypMD=S110012">衣</a> | <a href="?ModID=<%=MD%>&amp;TypMD=S110016">食</a> | <a href="?ModID=<%=MD%>&amp;TypMD=S110020">住</a></td>
          <%ElseIf MD="InfC124" Then%>
          <td height="30" align="left" class="SysTit02"><%=DPName%>: <a href="?ModID=<%=MD%>&DepID=ClsTyp212">电脑打字</a> | <a href="?ModID=<%=MD%>&DepID=ClsTyp216">会计考证</a> | <a href="?ModID=<%=MD%>&DepID=ClsTyp220">商务英语</a></td>
          <%ElseIf MD="InfD124" Then%>
          <td height="30" align="left" class="SysTit02"><%=DPName%>: <a href="?DepID=DepTech">技术设计</a> | <a href="?DepID=DepSalse">业务部</a> | <a href="?DepID=DepTest">测试</a></td>
          <%End If%>
          <td align="right" class="SysTit03"><%=vPMsg_WSite%><%=vPMsg_WHome%> &gt;&gt; <%=MDName%> <%=TPName%> &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><table cellspacing="0" cellpadding="0" width="99%" border="0">
        <tr>
          <td height="500" colspan="4" align="left" valign="top" style="padding:5px;"><!--Item Start-->
            <!-- #include file="../pfile/pinc/info_page.asp" -->
            <!--Item End--></td>
        </tr>
      </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240">
	<!--Side Start-->
	<!-- #include file="iside.asp" -->
    <!--Side End-->
    </td>
  </tr>
</table>
<!--#include file="_foot.asp"-->
</body>
</html>
