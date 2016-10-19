<!--#include file="_config.asp"-->
<!--#include file="../pfile/pinc/info_decode.asp"-->
<%
'Response.Write InfTyp2
If 1=2 Then

  xxxCont = Replace(xxxCont,vbcrlf&"<p>","")
  xxxCont = Replace(xxxCont,vbcrlf&"</p>","")
  xxxCont = Replace(xxxCont,vbcrlf&"  <p>","")
  xxxCont = Replace(xxxCont,vbcrlf&"  </p>","")
  xxxCont = Replace(xxxCont,"(TypName)","专家介绍")
  xxxCont = Replace(xxxCont,"(ImgName)",aImgs(1))
  xxxCont = Replace(xxxCont,"(RegLink)","#/page/duty.asp?Flag=Reg")
  xxxCont = Replace(xxxCont,"&#79;nload=","onload=")
  InfCont = xxxCont
  
  If xxxCont&InfCont="" Then
    xxxCont = "无资料，请添加!!!"
  End If
  
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=InfSubj%> - <%=MDName%> - <%=vPMsg_WName%></title>
<meta name="keywords" content="<%=KeyCode%>,<%=aPara(1)&aPara(3)%>" />
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
          <td width="210" height="30" align="left" class="SysTit02">&nbsp;</td>
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
              <!-- #include file="../pfile/pinc/info_view.asp" -->
            <!--Item End--></td>
          </tr>
        </table></td>
      <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240">
	<!--Side Start-->
	<!-- #include file="iside.asp" -->
    <!--Side End-->
    </td>    </tr>
</table>
<!--#include file="_foot.asp"-->
</body>
</html>
