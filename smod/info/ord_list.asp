<!--#include file="config.asp"-->
<!--#include file="../../pfile/lang/vmemb.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
	If KW&"" <> "" Then
	  sqlK = sqlK & " AND (KeyCode LIKE '%"&KW&"%') " 
	End If

cID = 0
sID = ""
If yAct="del_sel" then
  If Chk_URL3(Config_Path&"smod/info/ord_list.asp")="eUrl" Then
    Response.End()
  End If
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"DELETE FROM OrdInfo WHERE KeyID='"&iID&"'")
	  Call rs_DoSql(conn,"DELETE FROM OrdItem WHERE KeyCode='"&iID&"'")
	End If
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_N" then
  If Chk_URL3(Config_Path&"smod/info/ord_list.asp")="eUrl" Then
    Response.End()
  End If
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  cID = cID + 1
	  Call rs_DoSql(conn,"DELETE FROM OrdInfo WHERE KeyID='"&iID&"' AND LogAUser='"&Session("MemID")&"'")
	  Call rs_DoSql(conn,"DELETE FROM OrdItem WHERE KeyCode='"&iID&"' AND LogAUser='"&Session("MemID")&"'")
	End If
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf Left(yAct,4)="set_" then
  If Chk_URL3(Config_Path&"smod/info/ord_list.asp")="eUrl" Then
    Response.End()
  End If
  v = Mid(yAct,5,1)
  For iy = 1 To Request.Form("yID").Count
    cID = cID + 1
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
	Call rs_DoSql(conn,"UPDATE OrdInfo SET SetState='"&v&"' WHERE KeyID='"&iID&"'")
  Next
    Msg = cID&" 条记录设置成功!"

ElseIf yAct="del_Clear" then
  If Chk_URL3(Config_Path&"smod/info/ord_list.asp")="eUrl" Then
    Response.End()
  End If
    Call rs_DoSql(conn,"DELETE FROM OrdItem WHERE KeyCode='---' AND LogATime<#"&DateAdd("d",-5,Now())&"#")
    Call rs_DoSql(conn,"DELETE FROM OrdItem WHERE KeyCode<>'---' AND KeyCode NOT IN(SELECT KeyID FROM OrdInfo)")
    Call rs_DoSql(conn,"DELETE FROM OrdItem WHERE InfProID NOT IN(SELECT KeyID FROM InfoPics)")
	'Response.Write DateAdd("d",-5,Now())
    Msg = " 清理成功!"
End If

If PrmFlag="(Mem)" Then 
  sqlU = " AND LogAUser='"&Session("MemID")&"' "
Else
  sqlU = " "
End If

    sql = " SELECT * FROM [OrdInfo] "
	sql =sql& " WHERE 1=1 "&sqlK&sqlU ''KeyMod='"&ModID&"'
	sql =sql& " ORDER BY KeyID DESC" 'SetTop,
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

'Call rs_DoSql(conn,"UPDATE GboSend SET SetShow='Y' ")
'Response.Write verMemb&verNow

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong><%=vMem_ODA1%></strong> 
          <%If Session("UsrID")&""<>"" Then%>
           | <a href="ord_rep.asp">报表&gt;&gt;</a>
          <%End If%>
          </td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
              <input name="TPU" type="hidden" id="TPU" value="<%=IDPerm%>">
              <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
<%=vMem_ODA2%>:
<input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="<%=vMem_ODA3%>"></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&verMemb="&verMemb&"&KW="&KW&"&PrmFlag="&PrmFlag&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center"><%=vMem_ODA4%></td>
    <td height="27" align="center" nowrap><%=vMem_ODA5%></td>
    <td align="center" nowrap><%=vMem_ODA6%></td>
    <td width="12%" align="center" nowrap><%=vMem_ODA7%></td>
    <td width="12%" align="center" nowrap><%=vMem_ODA8%></td>
    <td width="12%" align="center" nowrap>金额</td>
    <td width="12%" height="27" align="center" nowrap><%=vMem_ODA9%></td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="10" align="right" nowrap></td>
  </tr>
  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If

KeyID = rs("KeyID")
KeyCode = rs("KeyCode")
InfSubj = rs("InfSubj")

LnkName = rs("LnkName")
LnkSex  = rs("LnkSex")

LogATime = rs("LogATime")'FormatDateTime(,2)
LogAUser = rs("LogAUser")
InfSubj = Show_Text(InfSubj) 

SetState = rs("SetState") : If SetState="-" Then SetState="N"
If PrmFlag<>"(Mem)" Then
  dFlag = true
ElseIf SetState="N" And PrmFlag="(Mem)" Then
  dFlag = true
Else
  dFlag = false
End If
If verMemb="2" Then
  SetState = Get_State(SetState,"N;D;P;S","New;Doing;Pay;Send")
Else
  SetState = Get_State(SetState,"N;D;P;S","新建;处理中;已付款;已发货")
End If
InfNum = rs("InfNum")
'aSum = rs_Val(conn,"SELECT SUM(InfPrice*InfCount) AS aSum FROM OrdItem WHERE KeyCode='"&KeyID&"'")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
       <%If dFlag Then%>
        <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>">
        <%Else%>
        <input name="xID" type="checkbox" id="xID" value="---" disabled>
        <%End If%>
        
        </td>
      <td align="center"><a href="ord_view.asp?ID=<%=KeyID%>&PrmFlag=<%=PrmFlag%>&verMemb=<%=verMemb%>" target="_blank"><%=KeyCode%></a></td>
      <td align="center" nowrap><%=InfSubj%></td>
      <td align="center" nowrap><%=LnkName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=SetState%></td>
      <td align="center" nowrap><%=InfNum%></td>
      <td align="center" nowrap><a href="ord_view.asp?ID=<%=KeyID%>&PrmFlag=<%=PrmFlag%>&verMemb=<%=verMemb%>" target="_blank"><%=vMem_ODA9%></a></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap><span id="yFlag" style="visibility:hidden ">N</span><%=vMem_ODB1%>
        <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TPU" type="hidden" id="TPU" value="<%=IDPerm%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td colspan="3" align="right" nowrap><select name="yAct" id="yAct" >
         
         <%If PrmFlag<>"(Mem)" Then%>
         <option value="set_N">设置_新建</option>
         <option value="set_D">设置_处理中</option>
         <option value="set_P">设置_已付款</option>
         <option value="set_S">设置_已发货</option>
         <option value="del_sel">删除.所选</option>
         <option value="del_Clear">清理.垃圾</option>
         <%Else%>
         <option value="del_N"><%=vMem_ODB2%></option>
		 <%End If%>
        </select></td>
      <td colspan="3" align="left" nowrap><input type="submit" name="Submit" value="<%=vMem_ODB3%>"></td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="10">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing

	  %>
    <tr bgcolor="#999999">
      <td colspan="10" align="right"></td>
    </tr>
  </form>
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
