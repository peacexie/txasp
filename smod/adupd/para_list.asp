<!--#include file="config.asp"-->
<!--#include file="../../upfile/sys/para/rkeyid.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

sCode = Replace(Replace(sEditParCode," ",""),",",";")
sFlag = Replace(Replace(sEditParFlag," ",""),",",";")
sMod  = Replace(Replace(sEditParMod," ",""),",",";")

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)

sqlK = ""
If KW&"" <> "" Then
  sqlK = sqlK & " ( ParCode LIKE '%"&KW&"%' "
  sqlK = sqlK & " OR ParFlag LIKE '%"&KW&"%' "
  sqlK = sqlK & " ) " 
End If

cID = 0
sID = ""

sqlM = ""
aCode = Split(sCode,";")
aMod  = Split(sMod,";")
For i=0 To uBound(aCode)
 If aMod(i)="Mod" Then
  sqlM = sqlM& " OR ParFlag='"&aCode(i)&"' "
 Else
  sqlM = sqlM& " OR ParCode='"&aCode(i)&"' "
 End If
Next
sqlM = Mid(sqlM,4)
If sqlK<>"" Then
 sqlK = sqlK&" AND ("&sqlM&") "
Else
 sqlK = sqlM
End If

	sql = " SELECT * FROM [AdmPara] "
	sql =sql& " WHERE "&sqlK&"" 
	sql =sql& " ORDER BY ParFlag DESC,ParCode "  ':Response.Write sql '[SetTop,
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table 
	width="100%" border="0" cellpadding="3" cellspacing="0">
        <form name="fm01" method="post" action="?">
          <tr align="center" bgcolor="#FFFFFF">
            <td width="30%" align="center" bgcolor="#FFFFFF"><strong>版权杂项 参数设置</strong></td>
            <td align="right" bgcolor="#FFFFFF">&nbsp;&nbsp;</td>
            <td width="30%" align="right" nowrap><A href="../../sadm/system/upd_para.asp">刷新参数</A> | <a href="?">重载本页</a>&nbsp;</td>
          </tr>
        </form>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" align="center" nowrap bgcolor="#F0F0F0">&nbsp;</td>
    <td width="20%" align="center" bgcolor="#F0F0F0">编号</td>
    <td align="center" nowrap bgcolor="#F0F0F0">[组别]名称</td>
    <td align="center" nowrap bgcolor="#F0F0F0">&nbsp;</td>
    <td colspan="2" align="center" nowrap bgcolor="#F0F0F0">修改</td>
  </tr>
  <%If Chk_PermSP() Then%>
  <tr bgcolor="#ffffff">
    <td align="right" nowrap bgcolor="#F8F8F8">---</td>
    <td bgcolor="#F8F8F8" class="col00F">rKeyID.asp </td>
    <td align="left" nowrap bgcolor="#F8F8F8">[ParRem]保留字,刷新项</td>
    <td align="center" nowrap bgcolor="#F8F8F8"></td>
    <td width="20%" align="center" nowrap bgcolor="#F8F8F8"><a href="para_set1.asp?ID=rKeyID.asp&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a></td>
    <td width="20%" rowspan="6" align="center" nowrap class="colF0F">右侧参数,已经配置好<br>
      正常运行的系统<br>
      请不要随便编辑<br>
      否则后果自负！！</td>
  </tr>
  <tr bgcolor="#ffffff">
    <td align="right" nowrap>---</td>
    <td bgcolor="#ffffff" class="col00F">rTmsPara</td>
    <td align="left" nowrap>[ParRem]模版参数备份</td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><a href="para_set1.asp?ID=rTmsPara&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a></td>
  </tr>
  <tr bgcolor="#ffffff">
    <td align="right" nowrap>---</td>
    <td bgcolor="#ffffff" class="col00F">tmsPara.asp </td>
    <td align="left" nowrap>[ParTemp]模版参数</td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><a href="para_set1.asp?ID=tmsPara.asp&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a></td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td align="right" nowrap>---</td>
    <td class="col00F">tmsTyp2.asp </td>
    <td align="left" nowrap>[ParTemp]次类别参数</td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><a href="para_set1.asp?ID=tmsTyp2.asp&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a></td>
  </tr>
  
  <tr bgcolor="#F8F8F8">
    <td align="right" nowrap>---</td>
    <td class="col00F">tmp_File.asp</td>
    <td align="left" nowrap>[ParTemp]文件模板</td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><a href="para_set1.asp?ID=tmp_File.asp&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a>
    </td>
  </tr>
  <tr bgcolor="#F8F8F8">
    <td align="right" nowrap>---</td>
    <td class="col00F">remind.js</td>
    <td align="left" nowrap>[ParRem]提醒参数</td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><a href="para_set1.asp?ID=remind.js&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a>
    </td>
  </tr>
  
  
  <tr bgcolor="#999999">
    <td colspan="6" align="center" nowrap style="line-height:1px;"></td>
  </tr>
  <%End If%>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
ParFlag = rs("ParFlag")
ParName = rs("ParName")
ParCode = rs("ParCode")
flag1 = Get_SOpt(sCode,sFlag,ParCode,"Val")
flag2 = Get_SOpt(sCode,sFlag,ParFlag,"Val")
'flag3 = Get_SOpt(sCode,sFlag,ParFlag,"Val")
	  %>
  <tr bgcolor="<%=col%>">
    <td align="right" nowrap><%=i%></td>
    <td><%=ParCode%></td>
    <td align="left" nowrap>[<%=ParFlag%>]<%=ParName%></td>
    <td align="center" nowrap></td>
    <td align="center" nowrap><%If flag1="Num" Then%>
      <a href="para_set1.asp?ID=<%=ParCode%>&Flag=Num&nLen=6&fRet=List">文本(Text代码)编辑</a>
      <%ElseIf flag1="Text" Then%>
      <a href="para_set1.asp?ID=<%=ParCode%>&Flag=Text&nLen=18&fRet=List">文本(Text代码)编辑</a>
	  <%Else%>
      <a href="para_set1.asp?ID=<%=ParCode%>&Flag=Rem&nLen=18&fRet=List">文本(Text代码)编辑</a>
      <%End If%></td>
    <td width="20%" align="center" nowrap><%If inStr(flag1&flag2,"Html")>0 Then%>
      <a href="para_set1.asp?ID=<%=ParCode%>&Flag=Editor&nLen=18&fRet=List">可视(FCK插件)编辑</a>
	  <%Else%>
      <%=flag1%>
	  <%End If%>
    </td>
  </tr>
  <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
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
	  Set rs = Nothing
	  
	  %>
  <tr align="center" bgcolor="#f4f4f4">
    <td colspan="9" align="left" bgcolor="#FFFFF0">注意：本列表在  系统与设置 &gt;&gt; 参数 中一般都可找到，这里只是把常用的参数提取出来，方便管理。</td>
  </tr>
</table>
<script type="text/javascript">
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>
</body>
</html>
