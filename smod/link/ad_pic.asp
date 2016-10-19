<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="config.asp"-->
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
TP = Request("TP") 

send = Request("send")
If send="ins" Then
  InfType = RequestS("InfType","C",24)
  InfName = RequestS("InfName","C",24)
  InfPath = RequestS("InfPath","C",60)
  InfCont = "" :sUrl = "" :sPic = ""
  For i=1 To 8
   sUrl = sUrl&RequestS("InfUrl"&i,"C",240)&"|"
   sPic = sPic&RequestS("InfPic"&i,"C",240)&"|"
  Next
  InfCont = sUrl&"(^)"&sPic
  InfPara = RequestS("InfPara1","N",120)&"|"&RequestS("InfPara2","N",60)&"|"&RequestS("InfPara3","N",3)
sql = " INSERT INTO [WebAdvert] (" 
sql = sql& "  KeyID,KeyMod" 
sql = sql& ", InfType, InfName, InfCont" 
sql = sql& ", InfPath, InfPara" 
sql = sql& ", LogAddIP,LogAUser,LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & Get_AutoID(24) &"','AdRPic'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & InfName &"'" 
sql = sql& ", '" & InfCont &"'"  
sql = sql& ", '" & InfPath &"'" 
sql = sql& ", '" & InfPara &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")" ':Response.Write sql
  Call rs_DoSql(conn,sql)
  Msg = "增加成功！"

End If

sqlK = " KeyMod='AdRPic' "
If TP<>"" Then 
  sqlK=sqlK&" AND (InfType LIKE '%"&TP&"%' OR InfName LIKE '%"&TP&"%')"
End If 

If yAct="del" Then
 sql = "DELETE FROM WebAdvert WHERE KeyID='"&RequestS("ID","C",48)&"'"
 Call rs_DoSql(conn,sql)
 Msg = cID&" 条记录删除成功!"
End If

 sql = " SELECT * FROM [WebAdvert] WHERE "&sqlK
 sql =sql& " ORDER BY InfType,KeyID DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 rs.PageSize = 12 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="40%" align="center" bgcolor="#FFFFFF"><a href="ad_list.asp">浮动广告</a> - <a href="ad_pic.asp">图片广告</a> - <a href="ad_text.asp">文字广告</a><br>            <strong>图片广告管理</strong></td>
          <td width="30%" align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <td align="center" nowrap>&nbsp;</td>
          <form name="fSearch" action="?">
            <td align="right" nowrap><input name="TP" type="text" id="TP" size="12" maxlength="24">
              <input type="submit" name="Submit" value="搜索">
              <input type="submit" name="button" id="button" value="刷新" onClick="javascript:window.open('ad_pview.asp')"></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="7" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">编号</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">位置说明</td>
    <td align="center" nowrap bgcolor="E0E0E0">大小-速度</td>
    <td align="center" nowrap bgcolor="E0E0E0">设置</td>
    <td align="center" nowrap bgcolor="E0E0E0">删除</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">更新</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="9" align="right" nowrap></td>
  </tr>
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
InfType = rs("InfType")
InfName = rs("InfName")
InfPara = rs("InfPara")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
'InfType = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
	  %>
      
  <form name="flist" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><input type="hidden" name="hiddenField">
        <%=i%> </td>
      <td align="center"><%=InfType%></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=InfName%></td>
      <td align="center" nowrap><%=InfPara%></td>
      <td align="center" nowrap><a href="ad_pedit.asp?ID=<%=KeyID%>&TP=<%=TP%>">设置</a></td>
      <td align="center" nowrap><a onClick="Del_YN('?ID=<%=KeyID%>&Page=<%=Page%>&yAct=del&TP=<%=TP%>','确认删除?')" href="#" >删除</a></td>
      <td align="center" nowrap><a href="ad_pview.asp?ID=<%=KeyID%>&Act=View" target="_blank">更新</a></td>
    </tr>
  </form>
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
	  
DefID = "xad"&Rnd_ID("0",4) 
	  
	  %>
  <tr bgcolor="#999999">
    <td colspan="9" align="right"></td>
  </tr>
  
  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" nowrap><input name="send" type="hidden" id="send" value="ins">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="InfType" type="text" id="InfType" value="<%=DefID%>" size="12" maxlength="12"></td>
      <td><input name="InfName" type="text" id="InfName" size="18" maxlength="24">
      </td>
      <td colspan="2">
      <input name="InfPara1" type="text" id="InfPara1" value="120" size="4" maxlength="4">        
      <input name="InfPara2" type="text" id="InfPara2" value="60" size="4" maxlength="4">
      <input name="InfPara3" type="text" id="InfPara3" value="3" size="4" maxlength="3"></td>
      <td colspan="2"><input type="button" name="Button" value="新增" onClick="chkData()">
      <input name="Page" type="hidden" id="Page" value="<%=Page%>">      </td>
    </tr>
  </form>
  <tr bgcolor="#FFFFFF">
    <td colspan="9">注意：为避免某些浏览器(遨游)拦截广告，编号请不要用guanggao,softad单词，且不要单独使用小写ad字母；</td>
  </tr>
</table>

<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.InfType.value.length==0)
           { alert('[错误]\n 代码!');
             document.ff.InfType.focus();
             errflag=0;
             break;
     }
	 if (document.ff.InfName.value.length==0)
           { alert('[错误]\n 请输入 名称!');
             document.ff.InfName.focus();
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
