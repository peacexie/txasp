<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/cch_Class.asp"-->
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
Set rs=Server.Createobject("Adodb.Recordset")

send   = Request("send") 
TypID   = Trim(RequestS("TypID",3,24))
TypName = RequestS("TypName",3,255) 
TypFlag = Request("TypFlag")
ModID   = RequestS("ModID",3,24) 
ParID   = RequestS("ParID",3,255)
Deeps   = Len(ParID)-Len(Replace(ParID,";",""))+1 
DeMax   = RequestSafe(rs_Val("","SELECT ParNum FROM AdmPara WHERE ParCode='n"&ModID&"'"),"N",1)
If Deeps&"">DeMax&"" Then
 Response.Redirect("?ModID="&ModID)
End If 'Response.Write Desql&DeMax&rs_Val("",Desql)
If Deeps = 1 Then
 DName   = "大类别设置"
Else
 DName   = ""&Deeps&"级子类别设置"
End If

If send = "ins" Then
  Msg = ""
  sql = "SELECT * FROM [WebTyps] WHERE TypID='"&TypID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&TypID&"]已经存在！"
  Else
	TypLayer = ParID&TypID&";"
	TypDeep = Len(TypLayer)-Len(Replace(TypLayer,";",""))
	sql = "INSERT INTO [WebTyps] (TypID,TypMod,TypName,TypLayer,TypDeep,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&TypID&"','"&ModID&"','"&TypName&"','"&TypLayer&"',"&TypDeep&",'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&TypID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [WebTyps] WHERE TypID='"&TypID&"'"
	Call rs_DoSql(conn,sql)
    sql = " DELETE FROM [WebTyps] WHERE TypLayer LIKE '%"&TypID&";%'"
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [WebTyps] SET TypName='"&TypName&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TypID='"&TypID&"'"
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
ElseIf send="Swap" Then
   TypFlg = RequestS("TypFlg",3,24) 
   TypNew = RequestS("TypNew",3,24) 
   ' WebTyps,TypID,TypID,S110016,S11x016
   If Len(TypID)>5 AND Len(TypNew)>5 Then
	 If TypFlg="Edit" Then
		nRes1 = SysDBSwap("TypID","TypID",   "WebTyps"," WHERE TypID    LIKE '%"&TypID&"%' ",TypID,TypNew)
		nRes2 = SysDBSwap("TypID","TypLayer","WebTyps"," WHERE TypLayer LIKE '%"&TypID&"%' ",TypID,TypNew) 
	 Else 'Swap
		nRes1 = SysDBSwap("TypID","TypID",   "WebTyps"," WHERE TypID    LIKE '%"&TypID&"%' ",TypID,"(~@_@~)")
		nRes2 = SysDBSwap("TypID","TypLayer","WebTyps"," WHERE TypLayer LIKE '%"&TypID&"%' ",TypID,"(~@_@~)") 
		nRes1 = SysDBSwap("TypID","TypID",   "WebTyps"," WHERE TypID    LIKE '%"&TypNew&"%' ",TypNew,TypID)
		nRes2 = SysDBSwap("TypID","TypLayer","WebTyps"," WHERE TypLayer LIKE '%"&TypNew&"%' ",TypNew,TypID) 
		nRes1 = SysDBSwap("TypID","TypID",   "WebTyps"," WHERE TypID    LIKE '%(~@_@~)%' ","(~@_@~)",TypNew)
		nRes2 = SysDBSwap("TypID","TypLayer","WebTyps"," WHERE TypLayer LIKE '%(~@_@~)%' ","(~@_@~)",TypNew) 
	 End If
   End If
 
   Response.Write "<br>"&nRes1&"条记录,处理完成:"&Now()&"<a href='?ModID="&ModID&"'>返回</a><br>"

ElseIf send="upd" Then
    Call cchClear()
%>
    <!--#include file="../type/code_tupd.asp"-->
<%
ElseIf send="Import" Then

'sVers = "中文,Eng,日語" '中文,Eng,日語,Deu
'aVers = Split(sVers,",")
MBase = Left(ModID,4)&"1"&Mid(ModID,6) 'PicS224
idType = Mid(ModID,5,1)
  
yAct = Request("yAct") 
'Sql = "SELECT * FROM WebTyps WHERE TypMod LIKE 'Inf%' OR TypMod LIKE 'Pic%' ORDER BY TypMod,TypLayer"
Sql = "SELECT * FROM WebTyps WHERE TypMod='"&MBase&"' ORDER BY TypMod,TypLayer"
rs.Open Sql,conn,1,1
Do While not rs.eof 

  TypID = rs("TypID")
  TypMod = rs("TypMod")
  TypLayer = rs("TypLayer")
  TypFlag = rs("TypFlag")
  TypName = rs("TypName") :TypName=Replace(TypName&"","'","")
  TypNam2 = rs("TypNam2")
  TypDeep = rs("TypDeep")
  TypResume = rs("TypResume") :TypResume=Replace(TypResume&"","'","")
  ImgName = rs("ImgName")
  
    i = idType
	iChar = Left(TypID,1)
	iTypID = Left(TypID,1)&i&Mid(TypID,3) '' S110080 
	iTypLayer = Replace(TypLayer,iChar&"1",iChar&i) '' S110048;S120048;
	iTypName = TypName&"("&i&")"
    sql = "INSERT INTO [WebTyps] (TypID,TypMod,TypName,TypLayer,TypDeep,TypNam2,TypFlag,TypResume,ImgName,LogAddIP,LogAUser,LogATime)VALUES"
    sql =sql& "('"&iTypID&"','"&ModID&"','"&iTypName&"','"&iTypLayer&"',"&TypDeep&",'"&TypNam2&"','"&TypFlag&"','"&TypResume&"','"&ImgName&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	If rs_Exist(conn,"SELECT TypID FROM [WebTyps] WHERE TypID='"&iTypID&"'")="EOF" AND Len(TypMod)=7 AND Len(TypID)=7 Then
	  Call rs_DoSql(conn,sql)
	  'Response.Write vbcrlf&"<br>----"&sql
	Else
	  Response.Write vbcrlf&"<br> 已经存在: "&iTypID&"','"&ModID&"','"&iTypName&"','"&iTypLayer&"',"&TypDeep
	End If

rs.Movenext
Loop
rs.Close()

Response.Write vbcrlf&"<br><center><a href='?ModID="&ModID&"'>返回</a></center>"
Response.End()

ElseIf send="Clear" Then
  Call rs_DoSql(conn,"DELETE FROM WebTyps WHERE TypMod='"&ModID&"'")
	Msg = "清理成功!"
Else
    '
End If

TPName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")
TypID = ""
    sql = " SELECT * FROM [WebTyps] " 
	sql =sql& " WHERE TypMod='"&ModID&"' "
	If ParID<>"" Then
	sql =sql& " AND TypLayer LIKE '"&ParID&"%' " 
	End If
	sql =sql& " AND TypDeep="&Deeps&" " 
	sql =sql& " ORDER BY TypLayer " 
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1

'If Left(ModID,5)="InfN1" Then
  'Re2Mod = "Info"
'ElseIf Left(ModID,5)="InfN2" Then
  'Re2Mod = "Inf5"
'Else
  
'End If
Re2Mod = Left(ModID,3)

%>
<br>
<table width="620" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="6" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="1" cellspacing="0">
        <tr align="center">
          <td width="30%" align="center" background="/img/tool/MTop.jpg"><strong>[<%=TPName%>]</strong> <%=DName%></td>
          <td width="30%" align="left" background="/img/tool/MTop.jpg">&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="form1" method="post" action="?">
            <td align="right" nowrap background="/img/tool/MTop.jpg"> <a href="../type/type_center.asp?MD=<%=Re2Mod%>">返回</a> <a href="../type/type_set.asp">设置</a> <a href="?send=upd" target="_blank">刷新</a>
              <%If Deeps >1 Then%>
              | <a href="?ModID=<%=ModID%>">返回跟类别</a>&nbsp;
              <%End If%>
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center">NO</td>
    <td height="27" align="center">类别代码</td>
    <td height="27" align="center">类别名称</td>
    <td align="center">子类别</td>
    <td height="27" align="center">修改</td>
    <td height="27" align="center">删除</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="8" align="right"></td>
  </tr>
  <%
  if not rs.eof then
  i = 0
  Do While NOT rs.EOF
  i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		TypID = rs("TypID")
		TypName = Trim(rs("TypName")&"")
		TypLayer = rs("TypLayer")

	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <%=i%></td>
      <td><%=TypID%>
        <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="ParID" type="hidden" id="ParID" value="<%=ParID%>"></td>
      <td><input name="TypName" type="text" id="TypName" value="<%=TypName%>" size="24" maxlength="48">
      </td>
      <td align="center" nowrap><%If Deeps&"" >= DeMax&"" Then%>
        <font color="#CCCCCC">子类别</font>
        <%Else%>
        <a href="?ParID=<%=TypLayer%>&ModID=<%=ModID%>">子类别</a>
		<%=rs_Count(conn,"WebTyps WHERE TypLayer LIKE '%"&""&TypID&";%' AND TypDeep="&Deeps+1&"")%>
        <%End If%>
      </td>
      <td align="center"><input type="submit" name="Submit" value="修改" >
      </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TypID=<%=TypID%>&ModID=<%=ModID%>&ParID=<%=ParID%>','<%=TypName%>')">
      </td>
    </tr>
  </form>
  <%
  rs.movenext
  Loop
  end if
	  
	  rs.close()

FstID = Mid(ModID,4,2)&Deeps
DefID = rs_TypID(conn,ModID,TypID,Deeps)
	  
	  %>
  <tr bgcolor="#cccccc">
    <td colspan="8" align="right"></td>
  </tr>
  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right"><input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="ParID" type="hidden" id="ParID" value="<%=ParID%>">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="TypID" type="text" id="TypID" value="<%=DefID%>" size="12" maxlength="12"></td>
      <td><input name="TypName" type="text" id="TypName" size="24" maxlength="48"></td>
      <td align="center">&nbsp;</td>
      <td colspan="2"><input type="button" name="Button" value="新增类别" onClick="chkData()">
      </td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="8"><span class="colF03">注意</span>: 类别代码请用<%=FstID%>开头,建议使用自动编号; 代码按大小排列,不按顺序标号是为了方便在中间插入类别代码;
      <%
		If inStr(Config_PCopy,ModID)>0 Then
		  tChr = Mid(ModID,5,1)
		  tMod = Mid(ModID,1,4)&"1"&Mid(ModID,6)
		  If CStr(tChr)<>"1" AND inStr(Config_PCopy,tMod)>0 Then
		    tArr = Split(Config_Langs,";")
		%>
      <br>      
      <a href="?ModID=<%=ModID%>&send=Import"><font color="#0000FF">同步</font></a>: 导入类别使[<%=TPName%>]类别信息与[<%=tArr(1)%>]中第一个语言版本同步；建议先编辑好中文版本的类别(含子类别)，再执行本命令，把中文中的类别一一对应“导入到此版本”,注意修改名称，不要指望本程序自动翻译！<br>      
      <a href="?ModID=<%=ModID%>&send=Clear"><font color="#0000FF">清理</font></a>: 清理类别使[<%=TPName%>]类别信息与[<%=tArr(1)%>]中第一个语言版本中没有的信息；即把此版本的类别全部删除，以便重建。(<span class="colF00">警告:稳定运行的系统请慎用,否则后果自负</span>) 
      <%
		  End If
		End If
		%></td>
  </tr>
  <form name="ffswap" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right"><input name="send" type="hidden" id="send" value="Swap">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="ParID" type="hidden" id="ParID" value="<%=ParID%>"></td>
      <td align="center" bgcolor="#e8e8e8">更改ID</td>
      <td colspan="4" align="right"><label for=""></label>
        <select name="TypFlg" id="TypFlg">
          <option value='Edit'>更改ID号</option>
          <option value='Swap'>交换顺序</option>
        </select>
        旧
        <input name="TypID" type="text" id="TypID" value="<%=DefID%>" size="12" maxlength="12">
        新
        <input name="TypNew" type="text" id="TypNew" value="<%=DefID%>" size="12" maxlength="12">        
        <input type="submit" name="Button" value="提交" xxClick="chkData()">      </td>
    </tr>
  </form>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.TypID.value.length==0)
           { alert('[错误]\n 类别代码不能为空!');
             document.ff.TypID.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TypName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TypName.focus();
             errflag=0;
             break;
     }
	 var tID = document.ff.TypID.value; 
	 //if ( (tID.substring(0,3)!="<%=FstID%>") )
           //{ alert('[错误]\n 代码请用"<%=FstID%>"开头!');
             //document.ff.TypID.focus();
             //errflag=0;
             //break;
     //}
        }
          if (errflag==1)
          {    
		  document.ff.submit()
          }
}
  
                </script>

<%
Set rs = Nothing
%>

</body>
</html>
