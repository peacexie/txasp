<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../func1/func_file.asp"-->
<!--#include file="config.asp"-->
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<%

TypID     = RequestS("TypID",3,24)
TypLayer  = RequestS("TypLayer",3,255) 
TypName   = RequestS("TypName",3,255) 
send      = Request("send") 

If send = "ins" Then
  Msg = ""
  sql2 = "SELECT * FROM [WebTyps] WHERE TypMod='"&ModID&"' AND TypID='"&TypID&"'"
  exF2 = rs_Exist(conn,sql2)
  If exF2="YES" Then
    Msg = "新增失败！["&TypName&"] [ID号]或[标识] 已经存在！"
  Else
    TypLayer = TypLayer&TypID&";"
	TypDeep = Len(TypLayer)-Len(Replace(TypLayer,";",""))
	sql = "INSERT INTO [WebTyps] (TypID,TypLayer,TypName,TypDeep,TypMod,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&TypID&"','"&TypLayer&"','"&TypName&"',"&TypDeep&",'"&ModID&"','"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&TypName&"] 新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [WebTyps] WHERE TypID='"&TypID&"' OR TypLayer LIKE '%"&TypID&"%' AND TypMod='"&ModID&"'  "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [WebTyps] SET TypName='"&TypName&"' "
	sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
	sql = sql & "WHERE TypID='"&TypID&"' AND TypMod='"&ModID&"'  "
	Call Chk_Perm1(IDPerm,"upd") 
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
ElseIf send = "ed2" Then
  sql2 = "SELECT * FROM [WebTyps] WHERE TypMod='"&ModID&"' AND TypID='"&TypName&"'"
  exF2 = rs_Exist(conn,sql2)
  If exF2="YES" Then
    Msg = "修改ID失败！["&TypName&"] [ID号] 已经存在！"
  Else
  oTypID = TypID
    sql = " SELECT * FROM [WebTyps] "
    sql =sql& " WHERE TypMod='"&ModID&"' AND TypLayer LIKE '%"&oTypID&";%'  "
    sql =sql& " ORDER BY TypLayer,TypID " 'LEN(TypLayer)
    Set rs=Server.Createobject("Adodb.Recordset")
    rs.Open Sql,conn,1,1 '3
	Do While NOT rs.EOF 
		TypID = rs("TypID")
		TypLayer = rs("TypLayer")
		TypLayer = Replace(TypLayer,oTypID,TypName)  
		Call rs_DoSql(conn,"UPDATE WebTyps SET TypLayer='"&TypLayer&"' WHERE TypID='"&TypID&"'")
	rs.Movenext
	Loop 
		Call rs_DoSql(conn,"UPDATE WebTyps SET TypID='"&TypName&"' WHERE TypID='"&oTypID&"'")
	rs.Close()
	Set rs = Nothing
	Msg = "["&oTypID&"] 修改ID成功！"
  End If
ElseIf send="upd" Then
 s=""
    sql = " SELECT * FROM [WebTyps] "
    sql =sql& " WHERE TypMod='"&ModID&"' " ' AND TypLayer LIKE '%"&oTypID&";%' 
    sql =sql& " ORDER BY TypLayer,TypID " 'LEN(TypLayer)
    Set rs=Server.Createobject("Adodb.Recordset")
    rs.Open Sql,conn,1,1 '3
	Do While NOT rs.EOF 
		TypID = rs("TypID")
		TypLayer = rs("TypLayer")
		TypName = rs("TypName")
		s=s&"<tr>"
		s=s&"<td width='22%' height='32' align='right'><img src='/images/index_118.gif' width='15' height='10'></td>"
		s=s&"<td><a href='/page/products.asp?TP="&TypID&"'><span class='STYLE9'>"&TypName&"</span></a></td>"
		s=s&"</tr>"
		'Call rs_DoSql(conn,"UPDATE WebTyps SET TypLayer='"&TypLayer&"' WHERE TypID='"&TypID&"'")
	rs.Movenext
	Loop 
		
	rs.Close()
	Set rs = Nothing
	Msg = " 修改ID成功！"

  set fso = CreateObject("Scripting.FileSystemObject") 
  set fil = fso.CreateTextFile(Server.MapPath("../../web/pics/sys/plist_.htm"),True)
    fil.Write(s)
    fil.Close
  set fso = Nothing

'<tr>
  '<td width="22%" height="36" align="right"><img src="/images/index_118.gif" width="15" height="10"></td>
  '<td><a href="Products.asp?mid=9"><span class="STYLE9">Punch line</span></a></td>
'</tr>

Else

End If

NowID = Get_AutoID(12)
    sql = " SELECT * FROM [WebTyps] "
	sql =sql& " WHERE TypMod='"&ModID&"'  "
	sql =sql& " ORDER BY TypLayer,TypID " 'LEN(TypLayer)
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="f0f0f0">
  <tr bgcolor="#e8e8e8">
    <td colspan="8"><FIELDSET>
      <LEGEND>
      <table width="100%" border="1" cellpadding="1" cellspacing="0">
        <tr bgcolor="#FFFFFF">
          <td width="20%" align="center" nowrap><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>]类别设定</strong></td>
          <td width="20%" align="center" nowrap><span id="ActTitle">请点击操作类别</span></td>
          <td nowrap><font color="#FF0000"><%=msg%>&nbsp;</font></td>
          <td width="15%" align="right" nowrap>&nbsp;</td>
        </tr>
      </table>
      </LEGEND>
      <table border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="f0f0f0">
        <form name="ff" method="post" action="?">
          <tr bgcolor="#e8e8e8">
            <td><span id="lbCode">标识</span>
              <input name="TypID" type="text" id="TypID" value="<%=NowID%>" size="18" maxlength="12" >
            </td>
            <td><span id="lbName">名称</span>
              <input name="TypName" type="text" id="TypName" size="24" maxlength="120"></td>
            <td>&nbsp;</td>
            <td><input type="button" name="DoAct" value="执行" onClick="chkData()" disabled>
              <a href="?send=upd">刷新</a>|<a href="type_set.asp?">设置</a> </td>
            <td>              <input name="send" type="hidden" id="send">
                <input name="TypLayer" type="hidden" id="TypLayer"></td></tr>
        </form>
      </table>
      </FIELDSET></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td align="center">NO</td>
    <td align="center">类别标识</td>
    <td align="center">类别名称</td>
    <td align="center">层级</td>
    <td align="center">增加</td>
    <td align="center">修改</td>
    <td align="center">改ID</td>
    <td align="center">删除</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="8" align="right"></td>
  </tr>
  <tr bgcolor="#F8F8F8" OnMouseOver='this.bgColor="#FFFFEE"' onMouseOut='this.bgColor="#F8F8F8"'>
    <td align="right">[0]</td>
    <td>[ID]</td>
    <td>┯ [ROOT]</td>
    <td align="center">0</td>
    <td align="center"><a href="#" onClick="SetData('增加类别','ins','','<%=NowID%>','','')">增加</a></td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
  </tr>
  <%
	  i = 0
	  Do While Not rs.eof 
	    i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#F8F8F8"
		end if
		TypID = rs("TypID")
		TypLayer = rs("TypLayer")
		TypName = Trim(rs("TypName")&"")
		ImgName = Trim(rs("ImgName")&"")
		TypDeep = rs("TypDeep")
		nLayStr = ""
		If TypDeep > 0 Then
		  For j = 1 To TypDeep-1
		    nLayStr = nLayStr & "│ "
		  Next
		End If
		nLayStr = nLayStr & "├─"
	  %>
  <tr bgcolor="<%=col%>" OnMouseOver='this.bgColor="#FFFFCC"' onMouseOut='this.bgColor="<%=col%>"'>
    <td align="right"><%=i%></td>
    <td><%=TypID%> </td>
    <td><%=nLayStr%> <%=TypName%></td>
    <td align="center"><%=TypDeep%></td>
    <td align="center"><a href="#" onClick="SetData('增加子类','ins','','<%=NowID%>','<%=TypID%>','<%=TypLayer%>')">增加</a></td>
    <%If nLayer < 9 Then%>
    <%
	Else
      Response.Write "<td align='center'><font color='#CCCCCC'>增加</font></td>"
	End If
	%>
    <td align="center"><a href="#" onClick="SetData('修改类别','edt','<%=TypName%>','<%=NowID%>','<%=TypID%>','<%=TypLayer%>')">修改</a> </td>
    <td align="center"><a href="#" onClick="SetData('修改ID号','ed2','<%=TypName%>','<%=NowID%>','<%=TypID%>','<%=TypLayer%>')">改ID</a></td>
    <%If nLayer < 9 Then%>
    <%
	Else
      Response.Write "<td align='center'><font color='#CCCCCC'>刷新</font></td>"
	End If
	%>
    <td align="center"><a href="#" onClick="SetData('删除类别','del','<%=TypName%>','<%=NowID%>','<%=TypID%>','<%=TypLayer%>')">删除</a> </td>
  </tr>
  <%
	    rs.movenext
	  loop
	  
	  rs.close()
	  set rs = nothing
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="8" align="right"></td>
  </tr>
</table>
<script type="text/javascript">
function SetData(xAct2,xAct,xName,xNowID,xThisID,xLayID)
{
  ActTitle.innerHTML = xAct2;
  document.ff.DoAct.disabled = false;
  document.ff.send.value = xAct ;
  if(xAct=='ins'){
  document.ff.TypName.value = xName ;
  document.ff.TypID.value = xNowID ;
  document.ff.TypLayer.value = xLayID ;}
  if(xAct=='ed2'){
    lbCode.innerHTML = '旧ID';
    lbName.innerHTML = '新ID';
    document.ff.TypName.value = xNowID ;
    document.ff.TypID.value = xThisID ;}
  if(xAct=='edt'){
    document.ff.TypName.value = xName ;
    document.ff.TypID.value = xThisID ;}
  if(xAct=='del'){
    document.ff.TypName.value = xName ;
    document.ff.TypID.value = xThisID ;}
}
<%if send="ins" then%>
SetData('增加子类','ins','<%=bTypName%>','<%=bTypID%>','','<%=bTypLayer%>');
<%end if%>
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.TypID.value.length==0)
           { alert('[错误]\n 请输入 类别标识!');
             errflag=0;
             break;
     }
	 if (document.ff.TypName.value.length==0)
           { alert('[错误]\n 请输入 类别名称!');
             document.ff.TypName.focus();
             errflag=0;
             break;
     }
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
  
                </script>
</body>
</html>
