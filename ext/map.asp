
<!--#include file="../page/_config.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE>站点导航-<%=Config_Name%></TITLE>
<META http-equiv=Pragma content=no-cache>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<style type="text/css">
.ItmMExt{
  width:90px;
  height:16px;
  border:1px solid #CCC;
  padding:4px 2px 3px 5px;
  margin:2px;1px;
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
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle">站点导航</td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="210" height="30" align="left" class="SysTit02">&nbsp;</td>
          <td align="right" class="SysTit03"><%=vPMsg_WSite%><%=vPMsg_WHome%> &gt;&gt; 站点导航 &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><div style="line-height:12px;">&nbsp;</div>
<%

a2Code = Split(sysModCode,"|")
a2Name = Split(sysModName,"|")
For i=0 To uBound(a2Code)
 If inStr("InfA,InfN,InfJ,InfP,InfV,PicS,PicR,PicT,PicV",Left(a2Code(i),4))>0 Then 'Left(a2Code(i),4)="InfN"

%>   
      <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" class="pgBorder">
        <tr>
          <td width="210" height="30" align="center" background="../pfile/pimg/qqimg21_12_01.gif" class="SysTitle"><a href="xxx<%=a2Code(i)%>"><%=a2Name(i)%></a></td>
          <td width="500" align="right" nowrap="nowrap" background="../pfile/pimg/qqimg21_12_02.gif">&nbsp;</td>
        </tr>
        <tr>
          <td height="30" colspan="2" align="left" valign="top" style="padding:8px 8px;">
          <%

   sMD = a2Code(i)
   bCode = Split(Eval("s"&sMD&"Code"),"|")
   bName = Split(Eval("s"&sMD&"Name"),"|")
   bNam2 = Split(Eval("s"&sMD&"Nam2"),"|")
  For j=0 to uBound(bCode)
    If bCode(j)<>"" Then
      iFlag = bNam2(j)
%>          
        <div class='ItmMExt'><a href="?TypID=<%=bCode(j)%>" class="as01" ><%=bName(j)%></a></div>
<%

    End If
  Next
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
