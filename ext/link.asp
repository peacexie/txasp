<!--#include file="../page/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE>友情连接-<%=Config_Name%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<style type="text/css">
.ItmLExt{
  width:140px;
  height:18px;
  border:1px solid #CCC;
  line-height:150%;
  padding:4px 5px 3px 8px;
  margin:2px;5px;
  float:left;
  overflow:hidden;
}
</style>
</HEAD>
<BODY>
<!--#include file="../page/_head.asp"-->
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle">友情连接</td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="210" height="30" align="left" class="SysTit02">&nbsp;</td>
          <td align="right" class="SysTit03"><%=vPMsg_WSite%><%=vPMsg_WHome%> &gt;&gt; 友情连接 &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><div style="line-height:12px;">&nbsp;</div>
<%

   sMD = "HomLnk1"
   aCode = Split(Eval("s"&sMD&"Code"),"|")
   aName = Split(Eval("s"&sMD&"Name"),"|")
   aNam2 = Split(Eval("s"&sMD&"Nam2"),"|")
  For i=0 to uBound(aCode)
    If aCode(i)<>"" Then
      iFlag = aNam2(i)
%>   
      <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="210" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle"><%=aName(i)%></td>
          <td width="500" align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif">&nbsp;</td>
        </tr>
        <tr>
          <td height="120" colspan="2" align="left" valign="top" style="padding:8px 8px;">
          <%
		  'sLink = ListLink(aCode(i),"")
		  'Response.Write sLink
		  
 xTemp="<div class='ItmLExt'><a href='($InfUrl)' target='_blank'>($InfSubj)</a></div>"
 s = ""
 tSql = " SELECT InfSubj,InfUrl,ImgName,SetSubj FROM GboLink WHERE InfType='"&aCode(i)&"' ORDER BY SetTop,KeyID DESC "
 rs.Open tSql,conn,1,1 
 Do While NOT rs.EOF
   InfSubj = rs("InfSubj")
   InfUrl = rs("InfUrl")
   ImgName = rs("ImgName")&""
   SetSubj = rs("SetSubj")
   InfSubj = Show_sTitle(InfSubj,SetSubj)
   s0 = xTemp
   s0 = Replace(s0,"($InfUrl)",InfUrl)
   s0 = Replace(s0,"($InfSubj)",InfSubj)
   s0 = Replace(s0,"($ImgName)",ImgName)
   s=s&s0
 rs.MoveNext
 Loop
 rs.Close()
 Response.Write s
		  
		  %>

          </td>
        </tr>
      </table>
      <div class="line10">&nbsp;</div>
<%
    End If
  Next
%>
      </td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240"><!-- #include file="ext_side.asp" --></td>
  </tr>
</table>
<!--#include file="../page/_foot.asp"-->
</BODY>
</HTML>
