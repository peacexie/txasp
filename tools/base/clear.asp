<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../himg/tconfig.asp"-->
<!--#include file="cfunc.asp"-->
<%Call Chk_Perm1("SysTools","")%>
<%

Act = Request("Act")
SysMod = clr_Reset("SysMod","DBClear") 'DBClear,FileClear
listTab = ""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>清洁工-木马扫把：清理数据库木马，文件木马，数据库压缩等</title>
<link rel="stylesheet" type="text/css" href="../himg/tstyle.css">
</head>
<body>
<%
If Session("SysMod")="" Then
  Response.Write "<center><br><a href='?SysMod=FileClear'>文件</a> | <a href='?SysMod=DBClear'>数据库</a></center>"
End If
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
' DBClear Start 
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
If Session("SysMod")="DBClear" Or Session("SysMod")="DBOut" Then
Set rs = Server.Createobject("Adodb.Recordset")
'123456789-123456789-123456789-123456789-123456789-123456789-123456789-
'<script src=http://g.cn/x.js></script-> Len > 24
TabName = Request("TabName") ' 表名
KeyName = Request("KeyName") ' 主键,
KeyChar = Request("KeyChar") ' 主键分割号 '或空
StrCols = Request("StrCols") ' Fld1,Fld2,Fld3 字段名
StrMuma = Request("StrMuma") ' <script src=http://g.cn/x.js></script->| 要清除的字符串
StrWhr = Request("StrWhr")   ' 条件,查看数据
strPars = "&KeyChar="&KeyChar&"&StrMuma="&Server.URLEncode(StrMuma)&"&StrWhr="&Server.URLEncode(StrWhr)&""
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="1" class="bgPub bgLBot">
  <tr>
    <td width="120" rowspan="2" align="left" valign="top" nowrap="nowrap">
      <table width="100%" border="0">
        <tr>
          <td colspan="2">=====数据表=====</td>
        </tr>
      <%
'/////////////////////////////////////
If SysDBType="Access" Then
'Const adSchemaTables = 20 'adSchemaColumns = 4
Set cnAcc = Server.CreateObject("ADODB.Connection")
cnAcc.Open conDB   
Set rsAcc = cnAcc.OpenSchema(20) 'adSchemaTablesA
i=0
Do while not rsAcc.EOF
  iTab = rsAcc("TABLE_NAME")
  iTyp = rsAcc("TABLE_TYPE") 'Replace(,"TABLE","")
  If inStr(iTab,">")<=0 And (iTyp="TABLE" Or iTyp="VIEW") Then
  i=i+1
  listTab = listTab&iTab&";"
  If i=1 And TabName="" Then
    TabName = iTab
  End If
  iCount = rs_Count(conDB," "&iTab&" ")
  If iCount>5000 Then iCount="<b>"&iCount&"</b>"
%>
        <tr>
          <td><a href="?TabName=<%=iTab%><%=strPars%>">
        <%If iTyp="VIEW" Then Response.Write("<font color=red>")%>
      <%=iTab%> </font></a></td>
          <td><%=iCount%></td>
        </tr>
      <%
  End If
 rsAcc.MoveNext
Loop
rsAcc.Close
cnAcc.Close
'/////////////////////////////////////
Else 'MsSql
sql = " SELECT [name],[uid] FROM [sysobjects] WHERE (xtype = 'U') ORDER BY [xtype] DESC,[name]  " 
rs.Open Sql,conDB,1,1
i=0
Do While NOT rs.EOF 
  i=i+1
  listTab = listTab&iTab&";"
  iTab = rs("name")
  iUid = rs("uid")
  If iUid>1 Then
    iUid = "<font color=red>&nbsp;(警告!可疑表!!!)</font>"
  Else
    iUid = ""
  End If
  If i=1 And TabName="" Then
    TabName = iTab
  End If
  If Left(Act,2)="DB" Then
   sAct = "&Act="&Act
  Else
   sAct = ""
  End If
  iCount = rs_Count(conDB," "&iTab&" ")
  If iCount>5000 Then iCount="<b>"&iCount&"</b>"
%>
        <tr>
          <td><a href="?TabName=<%=iTab%><%=sAct%><%=strPars%>"><%=iTab%></a> <%=iUid%></td>
          <td><%=iCount%></td>
        </tr>
      <%
rs.moveNext
Loop
rs.Close()
End If
'/////////////////////////////////////
%>
      </table>
    </td>
    <td width="150" rowspan="2" align="left" valign="top" nowrap="nowrap">====字段列表====
      <%
rs.Open " SELECT * FROM ["&TabName&"] ",conDB,1,1
tCols="" : fCols="" ': tColu=""
FOR i = 0 TO rs.Fields.Count-1
  fName = "["&rs.Fields(i).Name&"]"
  fType = rs.Fields(i).Type
  fSize = rs.Fields(i).DefinedSize
  If i=0 And KeyName="" Then
    KeyName = fName
  End If
  tColus=tColu&fName&", "
  If StrCols="" And fSize>24 Then
    tCols=tCols&fName&","
  End If
  fCols=fCols&fName&","
%>
      <div> <%=fName%> (<%=fType%>) <%=fSize%> </div>
    <%
NEXT
rs.Close()
If StrCols="" Then
  StrCols=Replace(tCols&"|",",|","")
End If
If inStr(StrCols,KeyName)<=0 Then
  StrCols=""&KeyName&","&StrCols
End If
  If Right(StrCols,2)=",|" Then
    StrCols=Left(StrCols,Len(StrCols)-2)
  End If
%>    </td>
    <td>
    <%If Session("SysMod")="DBOut" Then%>
    <table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
      <form id="fmDB2" name="fmDB" method="post" action="?">
        <tr>
          <td colspan="2" align="left" nowrap="nowrap"><textarea name="outTabs" cols="80" rows="5" id="outTabs"><%=Replace(listTab,";",vbcrlf)%></textarea></td>
        </tr>
        <tr>
          <td nowrap="nowrap">导出            
            <select name="Act" id="Act">
              <option value="DBStru" <%If Act="DBStru" Then Response.Write("selected")%>>[Stru]导出表结构 - 供PHP+MySql导入</option>
              <option value="DBData" <%If Act="DBData" Then Response.Write("selected")%>>[Data]导出表数据 - 供PHP+MySql导入</option>
            </select>
            <input type="submit" name="btnExecute2" id="btnExecute2" value="执行" /></td>
          <td nowrap="nowrap"><a href="?SysMod=DBClear">返回</a></td>
        </tr>
      </form>
    </table>
    <%Else%>
    <table border="0" align="center" cellpadding="0" cellspacing="0">
        <form id="fmDB" name="fmDB" method="post" action="?">
          <tr>
            <td align="left" nowrap="nowrap">表名:
              <input name="TabName" type="text" id="TabName" value="<%=TabName%>" size="12" maxlength="48" />
              &nbsp;&nbsp;&nbsp;主键
              <input name="KeyName" type="text" id="KeyName" value="<%=KeyName%>" size="12" maxlength="48" />
              &nbsp;&nbsp;&nbsp;分隔号:
              <input name="KeyChar" type="text" id="KeyChar" value="<%=KeyChar%>" size="12" maxlength="48" />
              &nbsp;&nbsp;&nbsp;&nbsp;<a href="?SysMod=DBOut">导出</a> | <a href="?SysMod=FileClear">文件</a> | <a href="../../member/ecard/admimp_code.asp?fstr=<%=fCols%>" target="_blank">生成</a> </td>
          </tr>
          <tr>
            <td align="left" nowrap="nowrap">字段:
              <textarea name="StrCols" cols="80" rows="2" id="StrCols"><%=StrCols%></textarea></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap="nowrap">木马:
              <textarea name="StrMuma" cols="80" rows="2" id="StrMuma"><%=StrMuma%></textarea></td>
          </tr>
          <tr>
            <td align="left" nowrap="nowrap">条件:
              <input name="StrWhr" type="text" id="StrWhr" value="<%=StrWhr%>" size="80" maxlength="255" />
              <select name="Act" id="Act">
                <option value="DBView">*浏览*</option>
                <option value="DBCheck" <%If Act="DBCheck" Then Response.Write("selected")%>>[检测]</option>
                <option value="DBClear" <%If Act="DBClear" Then Response.Write("selected")%>>清木马</option>
                <option value="DBCView" <%If Act="DBCView" Then Response.Write("selected")%>>预清除</option>
                <option value="DBExecu" <%If Act="DBExecu" Then Response.Write("selected")%>>[执行]</option>
              </select>
              <input type="submit" name="btnExecute" id="btnExecute" value="提交" /></td>
          </tr>
        </form>
    </table>
    <%End If%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="top"><div style="width:700px; height:470px; padding:3px; border:1px solid #CCCCCC; overflow:scroll; white-space:nowrap;">
        <li><span class="fnt00F">分隔号</span>: 主键分别为日期,字符,数字时分别使用#'和空; <span class="fnt00F">浏览</span>: N/M;浏览数据表中的数据;前N/后M个字符; <span class="fnt00F">检测</span>: &lt;script|&lt;/script&gt; ;检测木马代码;开始代码/结束代码;</li>
        <li><span class="fnt00F">清除</span>: &lt;script...&lt;/script&gt; ;清除木马代码;可用|分开多段代码; <span class="fnt00F">执行</span>: 执行SQL语句,如: DELETE FROM [<%=TabName%>] WHERE <%=KeyName%>=<%=KeyChar%>987123456<%=KeyChar%></li>
        <%

If Session("SysMod")="DBClear" Then
		
sql = " SELECT TOP 120 "&StrCols&" FROM ["&TabName&"] "
sql = sql& " "&StrWhr&" "
KeyNBak=KeyName :KeyName=Replace(KeyName,"[","") :KeyName=Replace(KeyName,"]","")
StrCBak=StrCols :StrCols=Replace(StrCols,"[","") :StrCols=Replace(StrCols,"]","")
Response.Write "<li><span class='fntF0F'>提示:</span> "&KeyName&":"&sql&"</li>"
Response.Write "<table border=0 cellspacing=0 cellpadding=0>"
  aCol = Split(StrCols,",")
  Response.Write "<tr>"
  for j=0 to uBound(aCol)
	 Response.Write "<th nowrap>&nbsp;"&aCol(j)&"&nbsp;</th>"
  next 
  Response.Write "</tr>"
  
rs.Open sql,conDB,1,1
i=0
Do While NOT rs.EOF
 Response.Write vbcrlf&"<tr>" 
 
 If Act="DBCheck" Then '///////////////////////////////////////////////////////////
  If StrMuma<>"" And inStr(StrMuma,"|")>0 Then
	aMa = Split(StrMuma,"|")
	sMa1 = aMa(0)
	sMa2 = aMa(1)
  Else
	sMa1 = "<script"
	sMa2 = "</script>" 
  End If 
  for j=0 to uBound(aCol)
     sVal = lCase(rs(aCol(j))&"")
	 p1 = inStr(sVal,sMa1)
	 p2 = InStrRev(sVal,sMa2)
	 If p1>0 And p2>0 And p2>p1 Then
	   sVal = Mid(sVal,p1,p2-p1+Len(sMa2))
	   sVal = clr_Text(sVal)
	   Response.Write "<td nowrap class=tdCode>"&sVal&"</td>"
	 Else 
	   Response.Write "<td nowrap>---</td>"
	 End If
  next 
 ElseIf Act="DBClear" Or Act="DBCView" Then '///////////////////////////////////////
  If Len(StrMuma&"")>12 Then
   cKey = rs(KeyName)
   aMa = Split(StrMuma,"|")
   sql = " UPDATE ["&TabName&"] SET "
   ' 0 为第一个KeyName
   for j=1 to uBound(aCol)
	 sVal = rs(aCol(j))&"" 
	 for k=0 to uBound(aMa) 
	   sVal = Replace(sVal,aMa(k),"")	
	   sVal = Replace(sVal,"'","''")
	 next
	 sql=sql&"["&aCol(j)&"]='"&sVal&"'"
	 if j< uBound(aCol) Then
	   sql=sql& ","
	 End If
   next 
   sql=sql& " WHERE "&KeyNBak&"="&KeyChar&""&cKey&""&KeyChar&""
   If Act="DBCView" Then
     Response.Write "<br>"&clr_Text(sql)
   Else
     Call rs_DoSql(conDB,sql) 
   End If
  End If
 ElseIf Act="DBExecu" Then  '///////////////////////////////////////////////////////
   If StrMuma<>"" Then
     Call rs_DoSql(conDB,StrMuma) 
   End If
 Else ' DBView '////////////////////////////////////////////////////////////////////
  If StrMuma<>"" And inStr(StrMuma,"|")>0 Then
	aMa = Split(StrMuma,"|")
	LenN = clr_Numb(aMa(0),8)
	LenM = clr_Numb(aMa(1),24)
  Else
	LenN = 8
	LenM = 24 
  End If 
  for j=0 to uBound(aCol)
     sVal = rs(aCol(j))&""
	 If Len(sVal)>Int(LenN)+Int(LenM) Then
	   sVal = Left(sVal,LenN)&"..."&Right(sVal,LenM)
	 End If
	 sVal = clr_Text(sVal)
	 Response.Write "<td nowrap class=tdCode>"&sVal&"</td>"
  next 
 End If  '//////////////////////////////////////////////////////////////////////////
 
 Response.Write "</tr>"
 i=i+1
 rs.MoveNext()
Loop
rs.close()
Response.Write "</table>"

'sql=sql& " WHERE "&KeyName&"="&KeyChar&""&cKey&""&KeyChar&""
'Response.Write "<br>"&sql

Else 'DBOut
  
	'act: DBStru,DBData
  If Act<>"" Then
	outTabs = Request("outTabs")
	aTab = Split(outTabs,vbcrlf)
    If Act="DBStru" Then
	  sTabs = "            set names utf8; "
	  For i=0 To uBound(aTab)-1
		iStr = dbMy_TabStru(aTab(i))
		sTabs = sTabs&iStr
	  Next
	  Call File_Add2(Config_Path&"upfile/#dbf#/dbMy_Table_Stru.txt",sTabs,"utf-8")
    Else
	  For i=0 To uBound(aTab)-1
		sData = dbMy_TabData(aTab(i))
		Call File_Add2(Config_Path&"upfile/#dbf#/dbMy_Data_"&aTab(i)&".txt",sData,"utf-8")
	  Next
	  'Call File_Add2(Config_Path&"upfile/#dbf#/dbMy_Stru.txt",sData,"utf-8")
    End If
    Response.Write "<br> "&aTab(i)
  End If
  
End If

Response.Write(vbcrlf&Act&":"&i&":完成！")

%>
      </div></td>
  </tr>
</table>
<%
Set rs = Nothing
End If
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
' DBClear End; FileClear Start images
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
If Session("SysMod")="FileClear" Then
'js.css,xml;gif,jpg,jpeg;asp,htm,html
FileType = Request("FileType") :If FileType="" Then FileType=".asp|.htm|.js|.xml|.css|.jpg|.gif"
FilePath = Request("FilePath") :If FilePath="" Then FilePath="../../" 
FileIgno = Request("FileIgno") :If FileIgno="" Then FileIgno="\images\|\inc\|img\|\temf\|\ext\|\tools\|\editor\|\edfck\|\upfile\" 'IgnorePath
FileCSet = Request("FileCSet") :If FileCSet="" Then FileCSet="gb2312"
StrMuma = Request("StrMuma") ':Response.Write "###"&StrMuma ' <script src=http://g.cn/x.js></script->| 要清除的字符串
FileName = Request("FileName") 
'strPars = "&KeyChar="&KeyChar&"&StrMuma="&Server.URLEncode(StrMuma)&"&StrWhr="&Server.URLEncode(StrWhr)&""
FilePRep = Server.MapPath(FilePath)
FileLStr = ""
FileLTim = ""
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="1" class="bgPub bgRBot">
  <tr>
    <td width="220" rowspan="2" align="left" valign="top" nowrap="nowrap">
    <div style="width:210px; height:610px; padding:3px; border:1px solid #CCCCCC; overflow:scroll; white-space:nowrap;">
    =====文件列表=====    
      <%Call clr_fList(FilePath)%>    
      </div>
    </td>
    <td valign="top"><table border="0" align="center" cellpadding="0" cellspacing="0">
      <form id="fmFile" name="fmFile" method="post" action="?">
        <tr>
          <td align="left" nowrap="nowrap">文件类型:
            <input name="FileType" type="text" id="FileType" value="<%=FileType%>" size="48" maxlength="120" />
            &nbsp;&nbsp;&nbsp;&nbsp;初始目录:
            <input name="FilePath" type="text" id="FilePath" value="<%=FilePath%>" size="12" maxlength="48" />
            &nbsp;            &nbsp;&nbsp;&nbsp;<a href="?SysMod=FileClear">刷新</a> | <a href="?SysMod=DBClear">数据库</a> | 登出 </td>
        </tr>
        <tr>
          <td align="left" valign="top" nowrap="nowrap">木马字符:
            <textarea name="StrMuma" cols="80" rows="2" id="StrMuma"><%=StrMuma%></textarea></td>
        </tr>
        <tr>
          <td align="left" nowrap="nowrap">忽略目录:
            <input name="FileIgno" type="text" id="FileIgno" value="<%=FileIgno%>" size="48" maxlength="120" />
&nbsp;&nbsp;&nbsp;&nbsp;文件编码:
            <input name="FileCSet" type="text" id="FileCSet" value="<%=FileCSet%>" size="12" maxlength="48" />
&nbsp;&nbsp;&nbsp;&nbsp;
            <select name="Act" id="Act">
                  <option value="FileView">*浏览*</option>
                  <option value="FileCheck" <%If Act="FileCheck" Then Response.Write("selected")%>>[检测]</option>
                  <option value="FileClear" <%If Act="FileClear" Then Response.Write("selected")%>>清木马</option>
                </select>
                <input type="submit" name="btnExecut2" id="btnExecut2" value="提交" /></td>
        </tr>
      </form>
    </table></td>
  </tr>
  <tr>
    <td align="left" valign="top"><div style="width:750px; height:515px; padding:3px; border:1px solid #CCCCCC; overflow:scroll; white-space:nowrap;">
      <li><span class="fnt00F">文件类型</span>: .asp|.aspx|.htm|.html|.js|.xml|.css|.jpg|.jpeg|.gif; <span class="fnt00F">初始目录</span>: /，../../等 <span class="fnt00F">忽略目录</span>: \img\|\tools\|\editor\; <span class="fnt00F">文件编码</span>: gb2312,big5,utf-8; <span class="fnt00F">提示：</span>: <%=Act%> - <%=clr_Text(StrMuma)%>;  </li> 
      <li><span class="fnt00F">浏览</span>: 前N/后M个字符; <span class="fnt00F">检测</span>: &lt;script|&lt;/script&gt;; <span class="fnt00F">清木马</span>: &lt;script...&lt;/script&gt; ;清除木马代码;可用|分开多段代码; <span class="fntF0F">提示</span>: 当前路径:(<%=FilePath%>)<%=Server.MapPath(FilePath)%></li>   
      <%
If Act="FileView" Then
  fArr = Split(FileLStr,"|")
  fArt = Split(FileLTim,"|")
  If StrMuma<>"" And inStr(StrMuma,"|")>0 Then
    nArr = Split(nArr,"|")
	nHead = nArr(0)
	nFoot = nArr(1)
  Else
    nHead = 8
	nFoot = 24
  End If
  For i=0 To uBound(fArr)-1
    Response.Write VBCRLF&"<li> *** "&fArt(i)&" "&fArr(i)&" <br>"
    iCont = File_Read(fArr(i),FileCSet)&""
    iCont = Left(iCont,nHead)&Right(iCont,nFoot)
    Response.Write clr_Text(iCont)
    Response.Write "</li>"
  Next
ElseIf Act="FileCheck" Then
  fArr = Split(FileLStr,"|")
  fArt = Split(FileLTim,"|")
  If StrMuma<>"" And inStr(StrMuma,"|")>0 Then
	aMa = Split(StrMuma,"|")
	sMa1 = aMa(0)
	sMa2 = aMa(1)
  Else
	sMa1 = "<script"
	sMa2 = "</script>" 
  End If 
  For i=0 To uBound(fArr)-1
    iCont = File_Read(fArr(i),FileCSet)&""
	p1 = inStr(iCont,sMa1)
	p2 = InStrRev(iCont,sMa2)
	If p1>0 And p2>0 And p2>p1 Then
	Response.Write VBCRLF&"<li> "&fArt(i)&" "&fArr(i)&" <br>"
	   iCont = Mid(iCont,p1,p2-p1+Len(sMa2))
	   If Len(iCont)>255 Then 
	   iCont = Left(iCont,80)&"..."&Right(iCont,30)
	   End If
	   iCont = clr_Text(iCont)
	   Response.Write " <font color=blue>"&iCont&"</font>"
	Response.Write "</li>"
	End If
  Next
ElseIf Act="FileClear" Then
  If Len(StrMuma&"")>12 Then
   aMa = Split(StrMuma,"|")
   fArr = Split(FileLStr,"|")
   fArt = Split(FileLTim,"|")
   Response.Write VBCRLF&"<li> 暂时不能用:<font color=red>字符比较[不正常!]</font> </li> "
   For i=0 To uBound(fArr)-1
     iCont = File_Read(fArr(i),FileCSet)&""
	 cntMa = 0
	 For k=0 to uBound(aMa) 
	   If Len(Trim(aMa(k)&""))>12 And inStr(iCont,aMa(k))>0 Then
	     iCont = Replace(iCont,aMa(k),"")	
		 cntMa = cntMa+1
	   Else
		 Response.Write VBCRLF&" <li> (No) </li> "
	   End If
	 Next
	 If cntMa>0 Then
	     ' Do something ...... <不正常!>
		 Response.Write VBCRLF&"<li> "&fArt(i)&" "&fArr(i)&" "
	     Response.Write "</li>"
	 End If
   Next 
  End If
End If
	  %>
    </div></td>
  </tr>
</table>



<%
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
' FileClear End; 
'///////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////
End If
%>
</body>
</html>
