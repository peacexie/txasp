<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title>导入查询数据</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../sadm/func1/WinFunc.js" type="text/javascript"></script>
</head>
<body>
      
<%

send   = Request("send") 
If send = "dFile" Then
  fp=Server.MapPath(Config_Path&"upfile/#dbf#/"&Request("fiName")) 'Server.MapPath("/stuinfo/#dbf#/")
  SET FSO = Server.CreateObject("Scripting.FileSystemObject")
  FSO.DeleteFile fp,TRUE
  Set FSO = Nothing
End If
%>
<table width="620" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#0066FF">
  <tr align="center" bgcolor="E0E0E0">
    <td  colspan="2" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>导入查询数据</strong></td>
          <td width="50%" align="left" bgcolor="#FFFFFF">&nbsp;<font color="#FF0000"><%=msg%>&nbsp;</font></td>
          <td align="right" nowrap>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <form name="form1" method="post" action="admimp2.asp" enctype="multipart/form-data">
    <tr bgcolor="#F0F0F0">
      <td align="right" bgcolor="#FFFFFF">(Excel文件名)</td>
      <td bgcolor="#FFFFFF"><input type="file" name="CrdFile" id="CrdFile" style="width:360px">
      <br>
      建议Excel文件在200KB以内</td>
    </tr>
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="ins"></td>
      <td bgcolor="#FFFFFF"><input type="submit" name="Button" value="提交" ></td>
    </tr>
    <tr bgcolor="#e8e8e8">
      <td colspan="2" align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;<span class="colF00">注意:</span> .Excel文件模板：<a href="../../upfile/<%=Server.URLEncode("#dbf#")%>/ysWeb_Demo_Card.xls">dyCard2.xls(点击下载或打开)</a></td>
    </tr>
    <tr bgcolor="#e8e8e8">
      <td colspan="2" align="left" bgcolor="#FFFFFF"><table width="100%" border='0' align="center" cellpadding='1' cellspacing='1' bordercolor='#E8E8E8' bgcolor="#EEEEEE">
        <tr bgcolor="fffff8">
          <td colspan='2' bgcolor="#FFFFFF"> ||==> <strong>Excel文件管理</strong> <span class="colF0F">(请经常清理这些垃圾文件)</span></td>
          <td colspan='4' bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <%
i=0 : s=0
SET objFSO = Server.CreateObject("Scripting.FileSystemObject")
IF objFSO.FolderExists(Server.MapPath(Config_Path&"upfile/#dbf#/")) THEN 
%>
        <tr bgcolor='#CCCCFF'>
          <td nowrap>Files/Folders</td>
          <td align='right' nowrap>Object.Size</td>
          <td align="center" nowrap>Created.Time</td>
          <td align="center" nowrap>&nbsp;</td>
          <td align="center" nowrap>&nbsp;Delete&nbsp;</td>
        </tr>
        <%
SET objFolder = objFSO.GetFolder(Server.MapPath(Config_Path&"upfile/#dbf#/"))
FOR EACH objFile IN objFolder.Files
fiName = objFile.Name
If inStr(fiName,"-")>0 And inStr(uCase(fiName),".XLS")>0 Then
i=i+1 : s=s+objFile.Size
if i mod 2 = 0 then
    row_col = "bgcolor=#f4f4f4"
else 
    row_col = "bgcolor=#FFFFFF"
end if
%>
        <tr <%=row_col%>>
          <td nowrap><%=fiName%></td>
          <td align='right' nowrap><%=objFile.Size%></td>
          <td align="center" nowrap class="txtSC"><%=objFile.DateLastModified%></td>
          <td align="center" nowrap class="txtSC">&nbsp;</td>
          <td align="center" nowrap class="txtSC"><a href='?send=dFile&fiName=<%=objFile.Name%>'>删除</a></td>
        </tr>
        <%
End If
NEXT

%>
        <tr bgcolor='#CCCCFF'>
          <td colspan='6' nowrap>&nbsp;All:&nbsp;<font color="#FF0000">[<%=i%>]</font>File(s) &nbsp;<font color="#FF0000">[<%=FormatNumber(s/1000,2)%>]</font>KB&nbsp;</td>
        </tr>
        <%
ELSE
%>
        <tr bgcolor='#CCCCFF'>
          <td colspan='6' nowrap bgcolor="#CCCCFF"> No files or folder...</td>
        </tr>
        <%
END IF
SET objFolder = Nothing
SET objFSO = Nothing
%>
      </table></td>
    </tr>
  </form>
</table>
</body>
</html>
