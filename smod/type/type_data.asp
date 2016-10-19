<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp"-->
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


yAct = Request("yAct")
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,60)

Group = RequestS("Group","C",24) : If Group="" Then Group="strNull"
Typ01 = RequestS("Typ01","C",240)
Typ02 = RequestS("Typ02","C",240)

'Response.Write Typ01 'PicS224:S210020;S220052;
aTypA = Split(Typ01,":")
If uBound(aTypA)=1 Then
  ModID = aTypA(0)
  TypID = aTypA(1)
  'Response.Write sModID&sTypID
Else
  ModID = ""
  TypID = ""
End If
ModTab = rel_ModTab(ModID)
'ModTab = "InfoNews" '


a0 = Split("Inf,Pic,Gbo,BBS,Tra,Mem",",")
s0 = "" : s1 = "" : s2 = "" : s3 = ""
For i=0 To 5
  s0 = s0&vbcrlf&"<a href='?Group="&a0(i)&"'>"&a0(i)&"</a> &nbsp; "
Next


' substring(SysID,4,3) 
'sqlin = "SELECT Mid(SysID,4) AS SysID FROM AdmSyst WHERE SysType='Module' AND SysID NOT IN('ModHome','ModSystem') " ''ModTrade',
'sql = " SELECT SysID,[SysName] FROM AdmSyst WHERE SysType IN("&sqlin&") ORDER BY SysType,SysID "
sql = " SELECT SysID,[SysName] FROM AdmSyst WHERE SysType LIKE '"&Group&"%' ORDER BY SysType,SysID "
Set rs = Server.Createobject("Adodb.Recordset")
rs.Open sql,conn,1,1 
Do While NOT rs.EOF
 sID = rs("SysID") 'Response.Write "<br>"&sID
 sName = rs("SysName")
 si = Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&sID&"' ORDER BY TypLayer",TypID,"Lay")
 si = vbcrlf&Replace(si,"<option value=''>┯ [ROOT]</option>","<option value=''>┯ "&sName&" ======== </option>")
 si = vbcrlf&Replace(si,"value='","value='"&sID&":") 
 If TypID="" And ModID=sID Then
   si = vbcrlf&Replace(si,"<option value='"&sID&":'>","<option value='"&sID&":' selected >")
 End If
 s1 = s1&si
 '//<optgroup label='"&sName&"'></optgroup>
rs.MoveNext
Loop
rs.Close()
SET rs = Nothing


'// 处理
aTypB = Split(Typ02,":")
If uBound(aTypB)=1 Then
  MDObj = aTypB(0)
  TPObj = aTypB(1)
Else
  MDObj = ""
  TPObj = ""
End If

If ModTab="GboInfo" Or ModTab="TradeInfo" Then
  TypID = Replace(TypID,";","")
  TPObj = Replace(TPObj,";","")
End If 'GboInfo,TradeInfo

If yAct="Move" Then
  If MDObj<>"" AND TypID<>TPObj Then
    Msg1 = "移动成功!"
    sql = "UPDATE "&ModTab&" SET KeyMod='"&MDObj&"',InfType='"&TPObj&"' "
	sql = sql & " WHERE KeyMod='"&ModID&"' AND InfType='"&TypID&"'"
    'Response.Write sql
    Call rs_Dosql(conn,sql)
  Else
    Msg1 = "未移动!"
  End If 
ElseIf yAct="List" Then
  yIDs = Request.Form("yID")
  If yIDs = "" Then
    Msg2 = ""
  ElseIf MDObj<>"" AND TypID<>TPObj Then
    Msg2 = "移动成功!"
	yIDs = Replace(Replace(yIDs," ",""),",","','")
    sql = "UPDATE "&ModTab&" SET KeyMod='"&MDObj&"',InfType='"&TPObj&"' "
	sql = sql & " WHERE KeyID IN('"&yIDs&"')"
    'Response.Write sql
    Call rs_Dosql(conn,sql)
  Else
    Msg2 = "未移动!"
  End If
End If




%>
<br>
<table width="640" border="1" align="center">
  <form name="fm02" method="post" action="?">
    <tr>
      <td colspan="4" align="center"><strong>资料移动</strong> - <span class="colF0F">正常运行的系统，请不要随便执行该操作！否则后果自负！！！</span></td>
    </tr>
    <tr>
      <td width="10%" align="center">组别</td>
      <td align="center"><%=s0%></td>
      <td width="10%" align="center">当前</td>
      <td align="center"><%=Group%>X999 ------<a href="?">&nbsp;刷新</a></td>
    </tr>
    <tr>
      <td align="center">移动</td>
      <td align="center"><select class=form id=select name=Typ01 style="width:180px; ">
          <option value="">请选择 ...</option>
		  <%=s1 %>
        </select></td>
      <td width="10%" align="center">到</td>
      <td align="center"><select class=form id=select name=Typ02 style="width:180px; ">
          <option value="">请选择 ...</option>
		  <%=s1 %>
        </select></td>
    </tr>
    <tr>
      <td align="center">操作</td>
      <td align="center"><select class=form id=yAct name=yAct style="width:180px; ">
        <option value="List">List_显示信息 (不用选目标类别)</option>
        <option value="Move" <%If yAct="Move" Then Response.Write("selected")%>>Move_移动到... (选择目标类别)</option>    
      </select></td>
      <td align="center">&nbsp;</td>
      <td align="left" style="padding-left:36px"><input type="submit" name="Submit" id="button" value="执行" xdisabled>
        &nbsp;
        <input name="Group" type="hidden" id="Group" value="<%=Group%>"> 
        <span class="colF00"><%=Msg1%></span></td>
    </tr>
  </form>
</table>
<br>
<%

If yAct="List" And ModID<>"" Then

sqlK = "" ': Response.Write ModTab&TypID
If TypID&"" <> "" Then
  sqlK = sqlK & " AND ( InfType LIKE '"&TypID&"%' ) " 
End If
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If

    sql = " SELECT * FROM ["&ModTab&"] "
	sql =sql& " WHERE KeyMod='"&ModID&"' "&sqlK
	sql =sql& " ORDER BY KeyID DESC " 
	'Response.Write sql
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 120 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>

<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="5" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="30%" align="center" bgcolor="#FFFFFF"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>] </strong> | 信息移动</td>
          <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<span class="colF00"><%=Msg2%></span></td>
          <form name="fm01" method="post" action="?">
            <td align="right" nowrap><span style="padding-left:36px">
              <input name="Group" type="hidden" id="Group" value="<%=Group%>">
              <input name="Typ01" type="hidden" id="Typ01" value="<%=Typ01%>">
              <input name="yAct" type="hidden" id="yAct" value="<%=yAct%>">
            </span>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="5" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&Typ01="&Typ01&"&KW="&KW&"&yAct="&yAct&"&Group="&Group&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="3%" height="27" align="center" nowrap>NO</td>
    <td width="60%" height="27" align="center">主题</td>
    <td height="27" align="center" nowrap>类别</td>
    <td align="center" nowrap>发布</td>
    <td height="27" align="center" nowrap>&nbsp;</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="7" align="right" nowrap></td>
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
InfType = rs("InfType")
If ModTab="GboInfo" Then
  TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypID='"&InfType&"'")
ElseIf ModTab="TradeInfo" Then
  TypName = rs_Val("","SELECT TypName FROM WebType WHERE TypID='"&InfType&"'")
Else
  TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
End If 'GboInfo,TradeInfo
If TypName="" Then TypName=InfType
InfSubj = rs("InfSubj")
LogATime = rs("LogATime")
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%>
      <input 
			name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="adm_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a></td>
      <td align="center" nowrap><%=TypName%></td>
      <td align="center" nowrap><%=LogATime%></td>
      <td align="center" nowrap>&nbsp;</td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>全选
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td align="right"><span style="padding-left:36px">
        <input name="Group" type="hidden" id="Group" value="<%=Group%>">
        <input name="Typ01" type="hidden" id="Typ01" value="<%=Typ01%>">
        <input name="yAct" type="hidden" id="yAct" value="<%=yAct%>">
      </span>
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        把选种的信息 移动到</td>
      <td colspan="2" align="right" nowrap><select name="Typ02" id="Typ02" xxxonChange="fAct()">
        <%=s1%>
      </select>
      <input type="submit" name="Submit2" value="执行"></td>
      <td align="left" nowrap>&nbsp;</td>
    </tr>
    <%  
  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="7">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="7" align="right"></td>
    </tr>
  </form>
</table>
<br>

<script type="text/javascript">

var fmObj = document.fm02;

function Send(xAct){
  fmObj.yAct.value = xAct;
  //fmObj.submit();
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
<%

Else '// Show Types Tree Lists


s2 = s1
s2 = vbcrlf&Replace(s2,"<option value='","") 
s2 = vbcrlf&Replace(s2,"</option>","") 
a2 = Split(s2,vbcrlf)
s2 = ""
For i=0 To uBound(a2)
  If inStr(a2(i),"'>")>0 Then
  a3 = Split(a2(i),"'>")
  s3 = "" 
  For j=1 To 30-Get_cLen(a3(1))
    s3 = s3&"."
  Next
  's3 = s3&Get_cLen(a3(1)&" "&s3) 
  s2 = s2&vbcrlf&a3(1)&" "&s3&" "&a3(0)
  End If
Next


sCss = "width:640px; height:240px; overflow-y:scroll;"
sCss = sCss& " line-height:130%; border:1px solid #CCC; padding:8px; "
Response.Write vbcrlf&"<center><textarea style="""&sCss&""">"
Response.Write vbcrlf&"12345678--------12345678--------12345678"
Response.Write s2&vbcrlf&"</textarea></center>"


Function Get_cLen(xStr)
Dim i,j,ch : xStr=xStr&""
 i=0 : j=0
 For i=1 To Len(xStr)
  ch = AscW(Mid(xStr,i,1))
  If ch>255 Then
    j=j+2 
  Else
    j=j+1
  End If
 Next
Get_cLen = j
End Function

End If

'///////////////////////////// for Test

' Asc,AscB,AscW 
' Len,LenB

Response.Write vbcrlf&"<!--"

Response.Write vbcrlf&"<br>"&Asc("中")
Response.Write vbcrlf&"<br>"
Response.Write vbcrlf&"<br>"&AscB("1")
Response.Write vbcrlf&"<br>"
Response.Write vbcrlf&"<br>"&AscW("1")
Response.Write vbcrlf&"<br>"
Response.Write vbcrlf&"<br>"&AscW("院")

Response.Write vbcrlf&"<br>"&LenB("商")
Response.Write vbcrlf&"<br>"&Len("商")
Response.Write vbcrlf&"<br>"&LenB("-")
Response.Write vbcrlf&"<br>"&Len("-")
'Response.Write vbcrlf&"<br>"&LenW("─")
Response.Write vbcrlf&"<br>"&Len("院")

Response.Write vbcrlf&"-->"

%>

</body>
</html>
