<!--#include file="binc/_config.asp"-->
<!--#include file="binc/bbsfunc.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../upfile/sys/para/keywords.asp"-->
<%

yAct = Request.Form("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,240)
MD = RequestS("MD",3,48) :If MD="" Then MD="(Self)"
 
 '  版主 管理的帖子
 mStr = ""
 lStr = ""
 sql = "SELECT TypID,TypName,TypLayer FROM [WebTyps] WHERE TypMod='"&ModID&"' AND TypNam2 LIKE '%"&Session("MemID")&"%' "
 rs.Open sql,conn,1,1 
 Do While Not rs.EOF 
  tID = rs("TypID")
  tLay = rs("TypLayer")
  tNM = rs("TypName")
  mStr = mStr&tID&","
  lStr = lStr&" | <a href='?MD="&tID&"&MN="&tNM&"'>"&tNM&"</a>"
 rs.MoveNext
 Loop
 rs.Close() 

'Response.Write "<BR>"&MD&mStr
If MD="(Self)" Then
  MName = "我的帖子"
  sqlK = " LogAUser='"&bbsUser&"' " 
ElseIf MD<>"" AND mStr<>"" And inStr(mStr,MD&",")>0 Then
  MN = Show_Text(Request("MN"))
  MName = ""&MN&" 帖子管理"
  sqlK = " InfType LIKE '%"&MD&"%' " 
Else
  MD="(Self)"
  MName = "我的帖子"
  sqlK = " LogAUser='"&bbsUser&"' " 
End If
'Response.Write "<BR>"&MD&mStr

If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" then
  sqlK = sqlK & " AND (InfType LIKE '%"&TP&"%') " 
End If
	
cID = 0
sID = ""
If yAct="SetShow" OR yAct="SetHot" Then
  yVal = RequestS("yVal","C",24)
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	End If
  Next
  sID = Replace(sID&"''",",''","")
  If sID<>"" Then
    sql = " UPDATE BBSInfo SET "&yAct&"='"&yVal&"' "
    sql = sql& " WHERE KeyID IN("&sID&")" ':Response.Write sql : Response.End()
	Call rs_DoSql(conn,sql)
  End If
    Msg = cID&" 条记录 设置成功!"
Elseif yAct="del_sel" then
  If Chk_URL3(Config_Path&"bbs/buser.asp")="eUrl" Then
    Response.End()
  End If
  For iy = 1 To Request.Form("yID").Count
    iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    If iID<>"" Then
	  sID = sID&"'"&iID&"',"
	  cID = cID + 1
	  Call bbsDelID(iID,"") 
	End If
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="del_now" Then
  If Chk_URL3(Config_Path&"bbs/buser.asp")="eUrl" Then
    Response.End()
  End If
  If KW&TP<>"" Then
	 sql = "SELECT KeyID FROM BBSInfo WHERE 1=1 AND "&sqlK&" "
	 rs.Open sql,conn,1,1
	 cID = rs.RecordCount
	 Do While NOT rs.EOF
 	   iID = rs("KeyID") 
	   Call del_sfDir(ModTab,iID)
	 rs.MoveNext
	 Loop
	 rs.Close()
    Msg = cID&" 条记录删除成功!"
  End If
  sqlK="" : TP="" : KW=""
  Page=1 ' 清楚条件,重设第一页
ElseIf yAct="Clear" Then 
	'sql = "DELETE FROM BBSInfo WHERE KeyMod='"&ModID&"' AND LEN(KeyRE)>12 AND KeyRe NOT IN(SELECT KeyID FROM BBSInfo) "
	'Call rs_DoSql(conn,sql)
    'Msg = " 清理成功!"
End If

 'Response.Write mStr

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=MName%>-<%=bbsName%></title>
<link rel="stylesheet" type="text/css" href="bimg/style.css">
</head>
<body>
<!--#include file="_itop.asp"-->
<div align="center" style="width:980px; height:auto; margin:auto; background-color:#F2F6FB">
  <table width="980" border="0" align="center" cellpadding="1" cellspacing="8">
    <tr>
      <td align="left" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; <%=MName%> &gt;&gt; </td>
      <td width="50%" align="right" bgcolor="#FFFFFF" class="fntFFF"><a href="../inc/mem_inc/mem_main.asp" target="_blank">我的后台</a> | <a href="<%=bbsLogin%>?send=out&amp;goPage=<%=bbsPath%>">登出系统</a> | <a href="bind.asp">返回首页</a>&nbsp;&nbsp;</td>
    </tr>
  </table>
  <%


    sql = " SELECT * FROM [BBSInfo] "
	sql =sql& " WHERE KeyMod='"&ModID&"' "
	If sqlK <>"" Then sql=sql&" AND "&sqlK 
	sql =sql& " ORDER BY KeyID DESC" 
   'Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
   rs.PageSize = 18 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="sysTabC1">
    <tr align="center">
      <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
          <tr align="center" bgcolor="#FFFFFF">
            <td width="30%" align="center" bgcolor="#FFFFFF"><strong><%=MName%></strong></td>
            <td width="30%" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
            <form name="fm01" method="post" action="?">
              <td align="right" nowrap>
                <a href="?">我的帖子</a> <%=lStr%>&nbsp;<br />
                <input name="send" type="hidden" id="send" value="sch">
                <input name="MD" type="hidden" id="MD" value="<%=MD%>">
                <input name="MN" type="hidden" id="MN" value="<%=MN%>" />
                <select class="form" id="select" name="TP" >
                  <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
                </select>
                <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
                <input type="submit" name="Submit" value="搜索"></td>
            </form>
          </tr>
        </table></td>
    </tr>
    <tr align="center">
      <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"&MD="&MD&"&MN="&MN&"",1)%></td>
    </tr>
    <tr align="center">
      <td width="5%" height="21" align="center" nowrap class="sysTopBG fntFFF">NO</td>
      <td align="left" class="sysTopBG fntFFF">帖子主题</td>
      <td width="8%" align="center" nowrap class="sysTopBG fntFFF"><%If MD="(Self)" Then%>
        类别
        <%Else%>
        发布人
        <%End If%>
      </td>
      <td width="5%" align="center" nowrap class="sysTopBG fntFFF">发布</td>
      <td width="8%" align="center" nowrap class="sysTopBG fntFFF">推荐</td>
      <td width="8%" align="center" nowrap class="sysTopBG fntFFF">审核</td>
      <td width="8%" align="center" nowrap class="sysTopBG fntFFF">回复</td>
      <td width="8%" align="center" nowrap class="sysTopBG fntFFF">修改</td>
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
KeyRE = rs("KeyRE")&"" 
KeyMod = rs("KeyMod")
KeyFlag = rs("KeyFlag")
InfType = Show_Text(rs("InfType"))
InfSubj = Show_Text(rs("InfSubj"))
InfCont = rs("InfCont")
InfReply = rs("InfReply")&""
LnkName = rs("LnkName")
LnkEmail = rs("LnkEmail")
SetRead = rs("SetRead")
SetShow = rs("SetShow")
SetSAdm = rs("SetSAdm")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
ImgName = rs("ImgName")
LogAddIP = rs("LogAddIP")
LogAUser = rs("LogAUser")
LogATime = rs("LogATime")
If KeyRe="" Then
 PID = KeyID
 InfReply = "主贴"
Else
 PID = KeyRe
 InfSubj = "<font color=gray>"&InfSubj&"</font>"
 InfReply = "<font color=gray>回复</font>"
End If
If MD="(Self)" Then
  TypName = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&InfType&"'")
Else
  TypName = LogAUser
End If
nReply = rs_Count(conn," BBSInfo WHERE KeyRE='"&KeyID&"' ")
nDays = DateDiff("d",LogATime,Date)
If SetShow="" Then SetShow="-"
SetShow = Get_State(SetShow,"Y;N;X;-","已审;未审;未过;未知")
If SetHot="" Then SetHot="-"
SetHot = Get_State(SetHot,"N;Y;-","一般;推荐;未知")
	  %>
      <tr bgcolor="<%=col%>">
        <td align="right" nowrap><%=i%>
            <input name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
        <td align="left"><a href="bview.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a>
            <%if SetHot="Y" then%>
            <img src="../img/tool/icon_jian.gif" width="15" height="15" align="absmiddle" />
            <%elseif cStr(nDays)="0" then%>
            <img src="../img/tool/new2.gif" width="30" height="10" align="absmiddle" />
            <%else%>
            <img src="../img/tool/email2.gif" width="16" height="16" align="absbottom" />
            <%end if%></td>
        <td nowrap><%=TypName%></td>
        <td align="center" nowrap><%=LogATime%></td>
        <td align="center" nowrap><%=SetHot%></td>
        <td align="center" nowrap><%=SetShow%></td>
        <td align="center" nowrap><%=InfReply%></td>
        <td align="center" nowrap><a href="bedit.asp?ID=<%=PID%>&amp;eID=<%=KeyID%>" target="_blank">修改</a></td>
      </tr>
      <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
      <tr class="sysTopBG fntFFF">
        <td align="right" nowrap bgcolor="#99B8DD"><span id="yFlag" style="visibility:hidden ">N</span>全选
          <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
        <td bgcolor="#99B8DD"><input name="Page" type="hidden" id="Page" value="<%=Page%>">
          <input name="TP" type="hidden" id="TP" value="<%=TP%>">
          <input name="KW" type="hidden" id="KW" value="<%=KW%>">
          <input name="MD" type="hidden" id="MD" value="<%=MD%>" />
          <input name="MN" type="hidden" id="MN" value="<%=MN%>" /></td>
        <td colspan="2" nowrap bgcolor="#99B8DD"><select name="yAct" id="yAct" onChange="fAct()">
            <option value="del_sel">删除.所选</option>
            <option value="del_now">删除.当前</option>
            <%If MD="(Self)" Then%>
            <%Else%>
            <option value="SetHot">设置_推荐</option>
            <option value="SetShow">设置_审核</option>
            <%End If%>
          </select>
          <select name="yVal" id="yVal">
          </select>
        </td>
        <td colspan="4" nowrap bgcolor="#99B8DD"><input type="submit" name="Submit" value="执行"></td>
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
	  'Set rs = Nothing
	  
	  %>
    </form>
  </table>
  <script type="text/javascript">

var sAct = document.getElementById("yAct");
var sVal = document.getElementById("yVal");

function fAct(){
  sVal.options.length = 0;
  if((sAct.value=="del_sel")||(sAct.value=="del_now")){
    sVal.options.add(new Option("删除","删除")); 
  }
  if(sAct.value=="SetShow"){
	sVal.options.add(new Option("已审","Y"));
	sVal.options.add(new Option("未审","N"));
	//sVal.options.add(new Option("未通过","X"));
  }
  if(sAct.value=="SetHot"){
	sVal.options.add(new Option("推荐","Y"));
	sVal.options.add(new Option("一般","N"));
  }
  if(sAct.value=="SetTop"){
    sVal.options.add(new Option("默认","888"));
    for(i=0;i<=9;i++){ sVal.options.add(new Option(i,i)); }
  }
  //alert(XAct);
}
<%
If yAct<>"" Then 
  Response.Write("sAct.value='"&yAct&"'")
End If
%>
fAct();

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
  <div style="line-height:8px;">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td height="24" align="center" bgcolor="#FFFFFF">&nbsp;<img src="../img/tool/email2.gif" width="16" height="16" align="absmiddle" /> 普通帖子&nbsp;&nbsp;&nbsp;<img src="../img/tool/icon_jian.gif" width="15" height="15" align="absmiddle" /> 推荐帖子&nbsp;&nbsp;&nbsp;<img src="../img/tool/new2.gif" width="30" height="10" align="absmiddle" /> 今日新帖&nbsp;</td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
</body>
</html>
