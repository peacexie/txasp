<!--#include file="inc/_config.asp"-->
<!--#include file="../member/admin/mconfig.asp"-->
<%
MDName = "会员黄页 "
Page = RequestS("Page","N",1)

TP = RequestS("TypID",3,240)
If TP&"" <> "" Then
  sqlK = " AND ( InfType LIKE '%"&TP&"%' ) " 
End If 
KW = RequestS("KeyWD",3,240)
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-商务与供求-<%=Config_Name%></title>
<link href="inc/style.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" type="text/css" href="../img/rnd_nid/box_nid.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../ext/api/play/jsPlayer.js" type="text/javascript"></script>
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
        <td align="right" class="SysTit03">您的位置: &nbsp;  &gt;&gt; 商务与供求 &gt;&gt;&nbsp; <%=MDName%> &nbsp; </td>
      </tr>
    </table></td>
  </tr>
</table>
<div class="PubClear"></div>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><div class="line05">&nbsp;</div>
    
    <!-- Content Start -->

<%
    sql = " SELECT * FROM [TradeCorp] WHERE 1=1 "&sqlK '" 'Corp 
	sql =sql& " ORDER BY SetTop,LogATime DESC "
	'Response.Write sql
   'Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 12
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>

<table width="96%"  border="0" align="center" cellpadding="2" cellspacing="1">

  <tr class="fnt666">
    <th height="27" nowrap>&nbsp;</th>
    <th align="left" nowrap>公司或单位名称 / 个人姓名</th>
    <th width="5%" align="center" nowrap>类别</th>
    <th width="5%" align="center" nowrap>行业</th>
    <th width="5%" align="center" nowrap>供求</th>
    <th width="15%" align="center" nowrap>发布日期</th>
  </tr>
  <tr>
    <td colspan="6" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
  </tr>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
KeyID = rs("KeyID")	
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfType&"'")
InfTyp2 = rs("InfTyp2")
TypNam2 = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfTyp2&"'")
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
SetRead = rs("SetRead")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfSubj = Show_sTitle(InfSubj,SetSubj)
	  %>
  <tr>
    <td width="3%" height="24" align="center">&middot;</td>
    <td> <a href="corp.asp?UsrID=<%=LogAUser%>" xtarget="_blank"><%=InfSubj%></a> <font color="#CCCCCC"><%=Get_SOpt(mCfgCode,mCfgName,KeyMod,"Val")%></font></td>
    <td width="5%" align="center" nowrap><%=TypName%></td>
    <td width="5%" align="center" nowrap><%=TypNam2%></td>
    <td width="5%" align="center" nowrap>&nbsp;<%=SetRead%>&nbsp;</td>
    <td width="15%" align="center" nowrap><%=LogATime%></td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
  <tr>
    <td colspan="6" nowrap bgcolor="#999999" class="td1px">&nbsp;</td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="6" align="center" bgcolor="#FFFFFF"><%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&KeyWD="&KW&"&TypID="&TP&"&TypMD="&TypMD&"&TPLay="&TPLay&"&Flag="&Flag&"",1)%></td>
  </tr>
  <%  
  Else
  %>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="6"><%=vPMsg_NoRec%></td>
  </tr>
  <%
  End If
	  
	  rs.Close()
	  'Set rs = Nothing
	  
	  %>
</table>

    <!-- Content Start -->

    
    
    </td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240"><!-- #include file="_side.asp" --></td>
  </tr>
</table>

<!--#include file="_foot.asp"-->
</body>
</html>
