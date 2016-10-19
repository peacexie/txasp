<!--#include file="config.asp"-->
<!--#include file="../../upfile/sys/para/tmp_File.asp" -->
<%

Set rs=Server.Createobject("Adodb.Recordset")
tMod = Trim(RequestS("tMod",3,24))
tTyp = Trim(RequestS("tTyp",3,255))
tFlg = Request("tFlg")
tLay = "" : fDis = ""
ID = Request("ID")
sqlK = " ParCode IN(SELECT 'tmp'+TypID FROM WebTyps WHERE TypMod='"&tMod&"' ) Or ParCode='tmp"&tMod&"' " '
If tTyp<>"" Then sqlK=sqlK& " Or ParCode='tmp"&tTyp&"' "

If tFlg="Sys" Then
  sqlK = " ParCode LIKE 'tst%' "
  fDis = " disabled "
ElseIf tFlg="File" Then
  fDis = " disabled "
ElseIf tFlg="Add" Then
    NM = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&ID&"'")
  If NM="" Then
    NM = rs_Val("","SELECT [SysName] FROM AdmSyst WHERE SysID='"&ID&"'")
  End If
  sql = " INSERT INTO [AdmPara] (" 
  sql = sql& "  ParName" 
  sql = sql& ", ParCode" 
  sql = sql& ", ParFlag" 
  sql = sql& ", ParRem" 
  sql = sql& ",LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & NM &"'" 
  sql = sql& ", 'tmp" & ID &"'" 
  sql = sql& ", 'ParTemp'" 
  sql = sql& ", '["&ID&"]"&NM&"(模版内容)'" 
  sql = sql& " ,'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"'"
  sql = sql& ")"
  Call rs_DoSql(conn,sql)
  Response.Redirect "../adupd/para_set1.asp?ID=tmp"&ID&"&Flag=Editor&nLen=12&fRet=Temp"
ElseIf tFlg="Del" Then
  Call rs_DoSql(conn,"DELETE FROM AdmPara WHERE ParCode='tmp"&ID&"'")
End If

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>模板选择，管理</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.bdTop {
	border-top:1px solid #CCC;
}
.bdBot {
	border-bottom:1px solid #CCC;
}
.bgCCC {
	background-color:#f0f0f0;
	padding:5px;
}
</style>
</head>
<body>
<div style="line-height:8px;">&nbsp;</div>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr>
    <td align="center" class="bgCCC"><a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=Sys">[系统模板]</a> 
    - <a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>">本模块模板</a> 
    <%If sTmpFile_Flag Then%>
    - <a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=File">[文件模板]</a> 
    <%End If%>
    - <a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=<%=tFlg%>">[重载本页]</a></td>
  </tr>
  <%If tFlg="File" Then%>
  <%

aTmpCode = Split(Eval("sTmpFile"&tMod)&";"&sTmpFile_Public,";")
aTmpName = Split(Eval("sTmpName"&tMod)&";"&sTmpName_Public,";")
's = sTmpFile_Public&";"&Eval("sTmpFile"&tMod)&tMod
'Response.Write s
For i=0 To uBound(aTmpCode)
  iCodf = aTmpCode(i) 
  If iCodf<>"" Then 
  iCode = Left(iCodf,Len(iCodf)-4)
  iName = aTmpName(i)
  iRem = File_Read("../../pfile/temf/"&iCodf,"utf-8")
  iRem = Replace(iRem,"[Site]",Config_Name)
  iRem = Replace(iRem,"[Date]",Date())
  iRem = Replace(iRem,"[Root]",Config_Path)
  iRem = Replace(iRem,"~"&iCode&"/",Config_Path&"pfile/temf/~"&iCode&"/")
  %>
  <tr>
    <td align="left">
    
      <fieldset style="padding:5px; border:1px solid #6CF">
        <legend> <span class="col00F"><%=i+1%></span> ------ <span <%=iCss%>><%=iCode%> - <%=iName%></span> ------
        <input type="button" name="Button" value="选择" onClick="owMsgUpd('<%=iCode%>');" >
        </legend>
        <div id="<%=iCode%>" style=" padding:5px;"> <%=iRem%> </div>
      </fieldset></td>
  </tr>
<%
  End If
Next  
  %>
  <%Else%>
  <tr>
    <td align="left" valign="top"><%

sql = " SELECT * FROM [AdmPara] WHERE "&sqlK&" "
sql =sql& " ORDER BY ParCode " 
rs.Open Sql,conn,1,1
i = 0
If NOT rs.EOF Then
Do While NOT rs.EOF
  i = i + 1
  iCode = rs("ParCode") : iTyp = Mid(iCode,4)
  iName = rs("ParName")
  iRem = rs("ParRem")
  iRem = Replace(iRem,"[Site]",Config_Name)
  iRem = Replace(iRem,"[Date]",Date())
  iRem = Replace(iRem,"[Root]",Config_Path)
  tLay = tLay&iTyp&";"
  iCss = "" : If inStr(tTyp,iTyp)>0 Then iCss = "class='col00F'"
%>
      <fieldset style="padding:5px; border:1px solid #6CF">
        <legend> <span class="col00F"><%=i%></span> ------ <span <%=iCss%>><%=iCode%> - <%=iName%></span> ------
        <input type="button" name="Button" value="选择" onClick="owMsgUpd('<%=iCode%>');" >
        ------ <a href="../adupd/para_set1.asp?ID=<%=iCode%>&Flag=Editor&nLen=12&fRet=Temp" target="_blank">[编辑]</a>
        <%If iTyp<>tMod And tFlg<>"Sys" Then%>
        <a onClick="Del_YN('?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=Del&ID=<%=iTyp%>','确认删除模版：\n[<%=iTyp%>]<%=iName%>？')" href="#" target="_self">[删除]</a>
        <%End If%>
        </legend>
        <div id="<%=iCode%>" style=" padding:5px; background-color:#F0F0F0;"> <%=iRem%> </div>
      </fieldset>
      <%
  rs.movenext
Loop
Else
%>
      <fieldset style="padding:5px; border:1px solid #6CF">
        <legend> 提示： </legend>
        无 模版资料 供选择，可在以下提示中 添加模版 ......
        <%If tFlg<>"Sys" Then%>
        或切换到 <A href="?tMod=InfA124&tTyp=&tFlg=Sys"><span class="colF0F">[系统模板]</span></A> 看看。
        <%End If%>
      </fieldset>
      <%
End If
rs.Close()
%></td>
  </tr>
  <%End If%>

  <tr>
    <td align="left" class="bgCCC">使用模版的优势：把<span class="colF00">经常使用的信息块或信息框架设置成模版</span>，可<span class="colF00">提高信息添加效率</span>，如把通知框架设置成模版；<br>
      如果点增加或编辑模版，在新窗口中设置好后直接关掉，然后在本窗口<a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=<%=tFlg%>">[<span class="colF0F">重载本页</span>]</a>即可；</td>
  </tr>
  
  <%If tFlg="File" Then%>
  <tr>
    <td align="left" class="bgCCC">文件模版说明：1. <span class="colF00">此部分需要FTP权限，自己做网页维护模版</span>；<br>
      2. 把模版做成网页，放在/pfile/temf/目录； <br>
    3. 一个模版对应一个文件和一个目录；文件名自定义,相关文件放在目录内，目录名为文件名前加~符号；<br>
    4. 如有css文件，也放在目录内，为避免重复，css里的定义项目前缀用tmf_；<br>
    5. 后台设置：系统与设置 &gt;&gt; 参数 &gt;&gt; 资料模版 &gt;&gt; tmp_File.asp文件模板 &gt;&gt; 设置显示哪些模版。</td>
  </tr>
  <%End If%>
  
  <form action="?" method="post" name="fmAdd" target="_blank" id="fmAdd">
    <tr>
      <td align="left" style="border:1px solid #6CF">可增加的<span class="bgCCC">模版</span>：
        <select name='ID' size=1 class="Input_Text" id="ID" style="width:240 " <%=fDis%>>
          <%
  If fDis="" Then
If inStr(tLay,tMod&";")<=0 Then
%>
          <OPTION value='<%=tMod%>'>[<%=tMod%>] --- [模块共用模板]</OPTION>
          <%
End If
		%>
          <%
sql = "SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&tMod&"'"
Response.Write sql
rs.Open sql,conn,1,1 
Do While Not rs.EOF 
sID = rs("TypID")
sName = rs("TypName")
If inStr(tLay,sID&";")<=0 Then
%>
          <OPTION value='<%=sID%>'>[<%=sID%>] --- <%=sName%></OPTION>
          <%
End If
rs.MoveNext
Loop
rs.Close()
  Else
    Response.Write "<OPTION value=''>[提示]不能增加系统模板和文件模板</OPTION>"
  End If
%>
        </select>
        <input type="submit" name="Button2" value="增加" <%=fDis%>>
        <input name="tFlg" type="hidden" id="tFlg" value="Add">
        &nbsp;<a href="?tMod=<%=tMod%>&tTyp=<%=tTyp%>&tFlg=<%=tFlg%>">[重载本页]</a></td>
    </tr>
  </form>
  
</table>
<script type="text/javascript">

function owMsgUpd(id) 
{
	window.opener.tmpCont = document.getElementById(id).innerHTML;
    window.opener.getTemp();
    window.close();		
}
</script>
<%
Set rs = Nothing
%>
</body>
</html>
