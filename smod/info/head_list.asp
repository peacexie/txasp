
<!--#include file="config.asp"-->
<!--#include file="head_config.asp"-->

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

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)
MD = RequestS("MD",3,48)
yVal = RequestS("yVal","C",24)
  
If MD&"" <> "" then
  sqlK = sqlK & " AND (KeyMod='"&MD&"') " 
End If
If TP&"" <> "" then
  sqlK = sqlK & " AND (InfTyp2 = '"&TP&"') " 
End If
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' "
  sqlK = sqlK & " OR InfCont LIKE '%"&KW&"%' "
  sqlK = sqlK & " ) " 
End If

If yAct="xSetShow" Then

Elseif yAct="del_sel" then
  cID = 0
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sImg = rs_Val("","SELECT ImgName FROM InfoHead WHERE KeyID='"&iID&"'")
	  If sImg<>"" Then
	    Call fil_del(Config_Path&sImg)
	  End If
	  Call rs_DoSql(conn,"DELETE FROM InfoHead WHERE KeyID='"&iID&"'")
	  cID = cID + 1
	End If
  Next
    Msg = cID&" 条记录删除成功!"
End If
    sql = " SELECT InfoHead.* FROM [InfoHead] "
	sql =sql& " WHERE 1=1 "&sqlK 'AND KeyMod='"&ModID&"' 
	sql =sql& " ORDER BY LogATime DESC " ':Response.Write sql
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

MDName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="9" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>头条设置</strong></td>
          <td width="30%" align="center" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><a href="../../sadm/admin/type_list.asp?ModID=InfHead&ModNM=新闻头条" target="_blank">类别</a>
              <input name="send" type="hidden" id="send" value="sch">
              <select class=form id=select name=MD >
			    <option value=''>[模块]</option>
				<%=Get_SOpt(OrgTCode,OrgTName,TP,"")%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="10">
              <select name="TP" id="TP">
                <option value=''>[位置]</option>
                <%=hOpt(TP2)%>
              </select>
              <input type="submit" name="Submit" value="搜索">
            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="9" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TP2="&TP2&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td width="40%" height="27" align="center">主题</td>
    <td align="center" nowrap>模块</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>位置</td>
    <td align="center" nowrap>发布时间</td>
    <td align="center" nowrap>发布人</td>
    <td height="27" align="center" nowrap>修改</td>
    <td align="center" nowrap>&nbsp;</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="11" align="right" nowrap></td>
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
KeyRE = rs("KeyRE") 
KeyMod = rs("KeyMod")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfSubj = Show_sTitle(InfSubj,SetSubj)
ModName = Get_SOpt(OrgTCode,OrgTName,KeyMod,"Val")
ImgFlag = ""
If ImgName<>"" Then
  ImgFlag = "<img src='../../img/tool/attfile.gif' border=0 align='absmiddle'>" 
End If
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
If TypName="" Then
  TypName = "<font color=gray>"&MDName&"</font>"
End If
TypNam2 = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfTyp2&"'")

	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input <%If tPerm="N" Then Response.Write("disabled") %> 
			name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="info_view.asp?ID=<%=KeyRE%>&ModID=<%=KeyMod%>" target="_blank"><%=InfSubj%></a><%=ImgFlag%>     </td>
      <td align="center" nowrap><%=ModName%></td> 
      <td align="center" nowrap><%=TypName%></td>
      <td align="center" nowrap><%=TypNam2%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap><%=LogAUser%></td>
      <td align="center" nowrap><a href="head_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>&MD=<%=MD%>">修改</a></td>
      <td align="center" nowrap>&nbsp;</td>
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
      <td colspan="11">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
      
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td>全选
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="MD" type="hidden" id="MD" value="<%=MD%>">
      <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
      <td colspan="4" align="left" nowrap><select class=form id=aModID name=aModID>
        <%=Get_SOpt(OrgTCode,OrgTName,TP,"")%>
      </select>
      <input name="aSubmit" type="button" id="aSubmit" value="增加" onClick="ActAdd()"></td>
      <td align="right" nowrap><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
        </select></td>
      <td colspan="2" align="left" nowrap><input type="submit" name="Submit" value="执行">      </td>
    </tr>
      
    <tr bgcolor="#999999">
      <td colspan="11" align="right"></td>
    </tr>
  </form>
</table>
<script type="text/javascript">

function ActAdd()
{
   var ObmModID = document.getElementById('aModID');
   location.href = "head_add.asp?MD="+ObmModID.value;
}  

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
