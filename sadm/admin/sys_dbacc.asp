<!--#include file="config.asp"-->

<html>
<head>
<title>数据库备份，压缩</title>
<meta http-equiv="Pragma" content="no-cache">
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>

<%

send = Request("send") 
DBType = cfgDBType ' Access,MSSQL
SET FSO = Server.CreateObject("Scripting.FileSystemObject")


'/////////////////////////////////////// DB-Access
If DBType = "Access" Then
'///////////////////////////////////////


prov = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

  oldDB = Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_"&Config_Code&".Peace!DB" )
  Set fObj = FSO.getfile(oldDB) 
    fSiz = FormatNumber(fObj.Size/1024,2)
  Set fObj = Nothing
If send = "dFile" Then
  fp=Server.MapPath(Config_Path&"upfile/#dbf#/"&Request("fiName")) 'Server.MapPath("/stuinfo/#dbf#/")
  FSO.DeleteFile fp,TRUE
  msg = "删除成功！"
ElseIf send="Back" Then
  bakDB = Server.MapPath(Config_Path&"upfile/#dbf#/Back_"&Get_yyyymmdd("")&Get_hhmmss()&Rnd_ID("KEY",3)&".Peace!Bak" )
  FSO.CopyFile oldDB, bakDB, true
  msg = "备份成功！"
ElseIf send="Pack" Then
  '(这过程也同时修复数据库)
  newDB = Server.MapPath(Config_Path&"upfile/#dbf#/Pack_"&Get_yyyymmdd("")&Get_hhmmss()&Rnd_ID("KEY",3)&".Peace!Bak" )
  Set Engine = CreateObject("JRO.JetEngine") 
   Engine.CompactDatabase prov & OldDB, prov & newDB 
  Set Engine = nothing 
  msg = "压缩成功！"
ElseIf send="Resave" Then
  fiName = Request("fiName")
  newDB = Server.MapPath(Config_Path&"upfile/#dbf#/"&fiName&"" )
  If FSO.FileExists(newDB) Then
   FSO.DeleteFile oldDB 
   FSO.CopyFile newDB, oldDB, true
  End If
  msg = "还原成功！"
  Call Add_Log(conn,Session("UsrID"),"还原成功","dbmaster",Msg)
ElseIf send="PExt02" Then 
  oldDB2 = Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_PubData.Peace!DB" )
  filID2 = Mid(Get_yyyymmdd("")&Get_hhmmss()&Rnd_ID("KEY",3),3,8)&"PubData"
  newDB2 = Server.MapPath(Config_Path&"upfile/#dbf#/Pack_"&filID2&".Peace!Bak" )
  Set Engine = CreateObject("JRO.JetEngine") 
   Engine.CompactDatabase prov & OldDB2, prov & newDB2 
  Set Engine = nothing  
  If FSO.FileExists(newDB2) Then
   FSO.DeleteFile oldDB2 
   FSO.CopyFile newDB2, oldDB2, true
   FSO.DeleteFile newDB2
   msg = "压缩PubData成功！"
  End If  
End If


'/////////////////////////////////////// DB-MSSQL
ElseIf DBType = "MSSQL" Then
'///////////////////////////////////////


sqlServer = cfgSqlServer   'sql服务器   
sqlUser = cfgSqlUser       '用户名   
sqlPassword = cfgSqlPassword    '密码   
sqlTimeout = 15   '登陆超时   
sqlDatabase = cfgSqlDatabase  

If send = "dFile" Then
  fp=Server.MapPath(Config_Path&"upfile/#dbf#/"&Request("fiName"))
  FSO.DeleteFile fp,TRUE
  msg = "删除成功！"
ElseIf send="NoLog" Then
  sql="backup log "&sqlDatabase&" with no_log" :Call rs_Dosql(conn,sql)
  sql="DBCC SHRINKDATABASE("&sqlDatabase&")"   :Call rs_Dosql(conn,sql)
  msg = "清理日志成功！"
ElseIf send="Backup" Then
  bak_file = Mid(Get_yyyymmdd("")&Get_hhmmss()&Rnd_ID("KEY",3),3,8)&"PubData"
  bak_file = Server.MapPath(Config_Path&"upfile/#dbf#/Pack_"&bak_file&".Peace!Bak" )
  Set srv=Server.CreateObject("SQLDMO.SQLServer")   
  srv.LoginTimeout = sqlTimeout 
  srv.Connect sqlServer,sqlUser, sqlPassword
  Set bak = Server.CreateObject("SQLDMO.Backup")   
  bak.Database=sqlDatabase   
  bak.Devices=Files   
  bak.Files=bak_file   
  bak.SQLBackup srv   
  if err.number>0 then   
    msg = "Error["&err.number&"]:"&err.description&"" 
  else
    msg = "备份成功!"  
  end if 
  Set srv = Nothing
  Set bak = Nothing  
ElseIf send="Resave" Then
  '恢复时要在没有使用数据库时进行！   
  fiName = Request("fiName")
  newDB = Server.MapPath(Config_Path&"upfile/#dbf#/"&fiName&"" )
  Set srv=Server.CreateObject("SQLDMO.SQLServer")   
  srv.LoginTimeout = sqlTimeout   
  srv.Connect sqlServer,sqlUser, sqlPassword
  Set rest=Server.CreateObject("SQLDMO.Restore")   
  rest.Action=0   '   full   db   restore   
  rest.Database=sqlDatabase    
  rest.Devices=Files   
  rest.Files=newDB   
  rest.ReplaceDatabase=True   'Force   restore   over   existing   database  
  if err.number>0 then   
    msg = "Error["&err.number&"]:"&err.description&"" 
  else
    msg = "备份成功!"  
  end if 
  Set srv = Nothing
  Set rest = Nothing
End If
  
  
'/////////////////////////////////////// DB-End
End If
'///////////////////////////////////////

%>
<br>
<table width="620" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#0066FF">
  <tr align="center" bgcolor="E0E0E0">
    <td colspan="2" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td align="center" bgcolor="#FFFFFF"><strong>数据库 备份压缩 / 临时文件 管理</strong></td>
          <td width="10%" align="right" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <td align="right" nowrap><%If Config_Mode = "isExpert" Then%>
            <a href="sys_dbinfo.asp">dbinfo</a> - 
            <a href="../../tools/base/clear.asp?SysMod=DBClear">dbtabs</a> - 
            <a href="../../tools/base/dbm.asp">dbmaster</a> - 
            <a href="../../tools/base/clear.asp?SysMod=FileClear">files</a>
          <%End If%></td>
        </tr>
      </table></td>
  </tr>
  <%If DBType = "Access" Then%>
  <form name="fm1" method="post" action="?">
    <tr bgcolor="#F0F0F0">
      <td align="left" bgcolor="#FFFFFF"> 当前数据库文件大小: <%=fSiz%> KB
<%
oldDB2 = Server.MapPath(Config_Path&"upfile/#dbf#/ysWeb_PubData.Peace!DB" )
If FSO.FileExists(oldDB2) Then
 Set fObj2 = FSO.getfile(oldDB2) 
 fSiz2 = FormatNumber(fObj2.Size/1024,2)
%>
        &nbsp; <a href="?send=PExt02"><span style="color:#FFF"><%=fSiz2%>KB</span></a>
<%
  Set FSO = Nothing
End If
%></td>
      <td width="20%" align="center" bgcolor="#FFFFFF"><input type="submit" name="Button2" value="无压缩备份" >
        <input name="send" type="hidden" id="send" value="Back"></td>
    </tr>
  </form>
  <form name="fm2" method="post" action="?">
    <tr bgcolor="#F0F0F0">
      <td align="left" bgcolor="#FFFFFF"> 提示，依次点：压缩备份&gt;修复(<span class="colF0F">选最近的备份文件</span>)&gt;删除，可优化数据库，节约空间。</td>
      <td align="center" bgcolor="#FFFFFF"><input type="submit" name="Button" value="-压缩备份-" >
        <input name="send" type="hidden" id="send" value="Pack"></td>
    </tr>
  </form>
  <%Else%>
  <form name="fm1" method="post" action="?">
    <tr bgcolor="#F0F0F0">
      <td align="left" bgcolor="#FFFFFF"> 还原时要在没有使用数据库时进行！当前数据库文件大小:[未知]</td>
      <td width="20%" align="center" bgcolor="#FFFFFF"><input type="submit" name="Button2" value="备份[Back]" >
      <input name="send" type="hidden" id="send" value="Backup"></td>
    </tr>
  </form>
  <form name="fm2" method="post" action="?">
    <tr bgcolor="#F0F0F0">
      <td align="left" bgcolor="#FFFFFF"> 提示，执行：压缩[NoLog]，可优化数据库，节约空间。</td>
      <td align="center" bgcolor="#FFFFFF"><input type="submit" name="Button" value="压缩[NoLog]" >
      <input name="send" type="hidden" id="send" value="NoLog"></td>
    </tr>
  </form>
  <%End If%>
  <tr bgcolor="#e8e8e8">
    <td colspan="2" align="left" bgcolor="#FFFFFF"><table width="100%" border='0' align="center" cellpadding='1' cellspacing='1' bordercolor='#E8E8E8' bgcolor="#EEEEEE">
        <tr bgcolor="fffff8">
          <td colspan='6' bgcolor="#FFFFFF"> ||==> <strong>备份文件 / 临时文件 管理</strong></td>
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
          <td align="center" nowrap>Resave</td>
          <td align="center" nowrap>&nbsp;Delete&nbsp;</td>
        </tr>
        <%
SET objFolder = objFSO.GetFolder(Server.MapPath(Config_Path&"upfile/#dbf#/"))
FOR EACH objFile IN objFolder.Files
fiName = objFile.Name
fiExt = Mid(fiName,InStrRev(fiName,"."),12)
If inStr(".Peace!Bak.Peace!Del.txt.xls",fiExt)>0 AND Left(fiName,5)<>"ysWeb" Then 
i=i+1 : s=s+objFile.Size
if i mod 2 = 0 then
    row_col = "bgcolor=#f4f4f4"
else 
    row_col = "bgcolor=#FFFFFF"
end if
iMsg = "郑重提示：您确认要恢复\n"&objFile.DateLastModified&"\n之前的数据？请小心操作！\n如误操作，后果自负！！！"
%>
        <tr <%=row_col%>>
          <td nowrap><%If inStr(".txt.xls",fiExt)>0 Then Response.Write("<a href='"&Config_Path&Server.URLEncode("upfile/#dbf#/"&objFile.Name)&"' target='_blank'>")%><%=fiName%></a></td>
          <td align='right' nowrap><%=FormatNumber(objFile.Size/1024,2)%> KB</td>
          <td align="center" nowrap class="txtSC"><%=objFile.DateLastModified%></td>
          <td align="center" nowrap class="txtSC">
          
          <%If fiExt=".Peace!Bak" Then%>
          <a href="#" onClick="Del_YN('?send=Resave&fiName=<%=objFile.Name%>','<%=iMsg%>')">修复/还原</a>
          <%Else%>
          <font color="#999999">修复/还原</font>
          <%End If%>
          
          </td>
          <td align="center" nowrap class="txtSC"><a href='?send=dFile&fiName=<%=objFile.Name%>'>删除</a></td>
        </tr>
        <%
End If
NEXT

%>
        <tr bgcolor='#CCCCFF'>
          <td colspan='6' nowrap>&nbsp;All:&nbsp;<font color="#FF0000">[<%=i%>]</font>File(s) &nbsp;<font color="#FF0000">[<%=FormatNumber(s/1024,2)%>]</font>KB&nbsp;</td>
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
Set fObj = Nothing 
Set fObj2 = Nothing 
%>
      </table></td>
  </tr>
  </tr>
        <tr bgcolor='#CCCCFF'>
          <td colspan='6' nowrap bgcolor="#FFFFCC">注意：后缀是.Peace!Bak的为备份文件, 建议一周或一月备份/压缩一次,保留最近3~5个备份文件;<br>
            后缀是.txt的为 特定场合下的 系统记录 文件, 可作为需要时分析, 建议保留最近3~5个.txt文件; <br>
          后缀是.Peace!Del的为 过滤的 临时文件, 可作为需要时分析, 建议保留最近几个临时文件,其它的删除;<br>
          后缀是.xls等文件 为 临时文件, 可直接删除; </td>
  </tr>
</table>
</body>
</html>
