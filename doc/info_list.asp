<!--#include file="dinc/_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理公文 -<%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  <%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,255)
TP2= RequestS("TP2",3,48)
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" then
  sqlK = sqlK & " AND (InfType LIKE '"&TP&"%') " 
End If
If TP2&"" <> "" then
  sqlK = sqlK & " AND (InfTyp2='"&TP2&"') " 
End If
cID = 0
sID = ""
  yVal = RequestS("yVal","C",24)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
  
If yAct="SetShow" Or yAct="SetHot" Or yAct="SetUBB" Or yAct="SetTop" Then
  If yAct="SetTop" Then
    yVal=RequestSafe(yVal,"N",888)
  Else
    yVal=Left(yVal,1)
  End If
  If sID<>"" Then
    sID = Replace(sID&"''",",''","")
	sql = " UPDATE DocsNews SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" 
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
	Call del_sfDir(ModTab,sID) 
    Msg = cID&" 条记录删除成功!"
End If
    sql = " SELECT DocsNews.* FROM [DocsNews] "
	sql =sql& " WHERE LogAUser='"&Session("InnID")&"' "&sqlK 'KeyMod='"&ModID&"'
	sql =sql& " ORDER BY LogATime DESC "  ':Response.Write sql
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If


%>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
    <tr align="center" bgcolor="E0E0E0">
      <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
          <tr align="center" bgcolor="#FFFFFF">
            <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[内部公文]管理 </strong> | <a href="info_add.asp?TPU=<%=IDPerm%>">新增&gt;&gt;</a></td>
            <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
            <form name="fm01" method="post" action="?">
              <td align="right" nowrap><input name="send" type="hidden" id="send" value="sch">
                <select class=form id=select name=TP >
                  <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
                </select>
                <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
                <input type="submit" name="Submit" value="搜索"></td>
            </form>
          </tr>
        </table></td>
    </tr>
    <tr align="center" bgcolor="E0E0E0">
      <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&TP2="&TP2&"",1)%></td>
    </tr>
    <tr align="center" bgcolor="E0E0E0">
      <td width="5%" height="27" align="center" nowrap>NO</td>
      <td width="40%" height="27" align="center">主题</td>
      <td height="27" align="center" nowrap>类别</td>
      <td align="center" nowrap>发布</td>
      <td align="center" nowrap>阅读</td>
      <td align="center" nowrap>记录</td>
      <td height="27" align="center" nowrap>修改</td>
      <td align="center" nowrap>组别</td>
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
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
If TypName="" Then
  TypName = "<font color=gray>"&MDName&"</font>"
End If
InfSubj = rs("InfSubj")
SetSubj = rs("SetSubj")
InfTo = rs("InfTo")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime")
InfSubj = Show_sTitle(InfSubj,SetSubj)
InfTyp2 = rs("InfTyp2")
InfTyp2 = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&InfTyp2&"'")

fRead = rs_Count(conn,"DocsLogs WHERE KeyID='"&KeyID&"'")
If fRead&""="" Then
fRead = "<span class=fntF0F>未读</span>"
End If

	  %>
      <tr bgcolor="<%=col%>">
        <td align="right" nowrap><%=i%>
          <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
        <td align="left"><a href="doc_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a><%=xxxInfTo%></td>
        <td nowrap><%=TypName%><%=""&TypNam2%></td>
        <td align="center" nowrap><%=LogATime%></td>
        <td align="center" nowrap><%=fRead%></td>
        <td align="center" nowrap><a href="#" onclick="javascript:window.open('doc_logs.asp?ID=<%=KeyID%>','win_<%=Rnd_ID("",5)%>','width=640,height=480')">记录</a></td>
        <td align="center" nowrap><a href="info_edit.asp?ID=<%=KeyID%>&PG=<%=Page%>&KW=<%=KW%>&TP=<%=TP%>">修改</a></td>
        <td align="center" nowrap><%=InfTyp2%></td>
      </tr>
      <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
      <tr bgcolor="E0E0E0">
        <td height="21" align="right" nowrap><span id="yFlag" style="visibility:hidden ">N</span>
          <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
        <td align="left">全选
          <input name="Page" type="hidden" id="Page" value="<%=Page%>">
          <input name="TP" type="hidden" id="TP" value="<%=TP%>">
          <input name="TP2" type="hidden" id="TP2" value="<%=TP2%>">
          <input name="KW" type="hidden" id="KW" value="<%=KW%>"></td>
        <td nowrap>&nbsp;</td>
        <td colspan="2" align="center" nowrap><select name="yAct" id="yAct" >
            <option value="del_sel">删除.所选</option>
            <!--
          <option value="SetShow">设置_显示</option>
          <option value="SetHot" selected>设置_推荐</option>
		  <option value="SetTop">设置_顺序</option>
          -->
          </select>
          <select name="yVal" id="yVal">
            <option value="Y">Y</option>
            <option value="N">N</option>
            <option value="X">X</option>
          </select></td>
        <td colspan="3" align="left" nowrap><input type="submit" name="Submit" value="执行"></td>
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
</div>
<%jsFlag="mTopAdm"%>
<!--#include file="dinc/inc_bot.asp"-->
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
