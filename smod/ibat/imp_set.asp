<!--#include file="config.asp"-->
<%

'Call Remote__Test()
'Response.Write Request.ServerVariables("HTTP_HOST")
ModName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")

yAct = Request("yAct")

If yAct="Add" Or yAct="Copy" Then
  ID = Get_AutoID(18)
  InfSubj = RequestS("InfSubj",3,48)
  sql = " INSERT INTO [InfoOuter] (" 
  sql = sql& "  KeyID,KeyMod,InfSubj,InfType" 
  sql = sql& ", InfTyp2, InfUrl, InfCSet" 
  sql = sql& ", InfP1Str,InfP1,InfP2Str,InfP2"  
  sql = sql& ", InfDel1,InfDel2" 
  sql = sql& ", InfQ1Str,InfQ1,InfQ2Str,InfQ2" 
  sql = sql& ", InfRep1,InfRep2,InfKill" 
  sql = sql& ", LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & ID &"','"&ModID&"','"&InfSubj&"','"&RequestS("InfType",3,255)&"'" 
  sql = sql& ", '" & RequestS("InfTyp2",3,12) &"','" & RequestS("InfUrl",3,255) &"','" & RequestS("InfCSet",3,24) &"'" 
  sql = sql& ", '" & RequestS("InfP1Str",3,255) &"'," & RequestS("InfP1","N",0) &",'" & RequestS("InfP2Str",3,255) &"'," & RequestS("InfP2","N",0) &"" 
  sql = sql& ", " & RequestS("InfDel1","N",0) &"," & RequestS("InfDel2","N",0) &"" 
  sql = sql& ", '" & RequestS("InfQ1Str",3,255) &"'," & RequestS("InfQ1","N",0) &",'" & RequestS("InfQ2Str",3,255) &"'," & RequestS("InfQ2","N",0) &"" 
  sql = sql& ", '" & RequestS("InfRep1",3,255) &"','" & RequestS("InfRep2",3,255) &"','" & RequestS("InfKill",3,255) &"'" 
  sql = sql& ", '" & Get_CIP() &"','"&Session("UsrID")&"','" & Now() &"'" 
  sql = sql& ")" ':Response.Write sql
  If rs_Exist(conn,"SELECT KeyID FROM InfoOuter WHERE KeyMod='"&ModID&"' AND InfSubj='"&InfSubj&"'")="YES" Then
    sMsg = "增加失败，["&InfSubj&"]已经存在!"
  Else
    Call rs_DoSql(conn,sql)
	sMsg = "增加成功!" : If yAct="Copy" Then sMsg="复制成功!"
  End If

ElseIf yAct="Edit" Then
  ID = Request("ID")
  sql = " UPDATE [InfoOuter] SET " 
  sql = sql& " InfSubj = '" & RequestS("InfSubj",3,48) &"'" 
  sql = sql& ",InfType = '" & RequestS("InfType",3,255) &"'" 
  
  sql = sql& ",InfTyp2 = '" & RequestS("InfTyp2",3,12) &"'"
  sql = sql& ",InfUrl = '" & RequestS("InfUrl",3,255) &"'"
  sql = sql& ",InfCSet = '" & RequestS("InfCSet",3,24) &"'"
  
  sql = sql& ",InfP1Str = '" & RequestS("InfP1Str",3,255) &"'"
  sql = sql& ",InfP1 = " & RequestS("InfP1","N",0) &""
  sql = sql& ",InfP2Str = '" & RequestS("InfP2Str",3,255) &"'"
  sql = sql& ",InfP2 = " & RequestS("InfP2","N",0) &""
  
  sql = sql& ",InfDel1 = " & RequestS("InfDel1","N",0) &""
  sql = sql& ",InfDel2 = " & RequestS("InfDel2","N",0) &""
  
  sql = sql& ",InfQ1Str = '" & RequestS("InfQ1Str",3,255) &"'"
  sql = sql& ",InfQ1 = " & RequestS("InfQ1","N",0) &""
  sql = sql& ",InfQ2Str = '" & RequestS("InfQ2Str",3,255) &"'"
  sql = sql& ",InfQ2 = " & RequestS("InfQ2","N",0) &""
  
  sql = sql& ",InfRep1 = '" & RequestS("InfRep1",3,255) &"'"
  sql = sql& ",InfRep2 = '" & RequestS("InfRep2",3,255) &"'"
  sql = sql& ",InfKill = '" & RequestS("InfKill",3,255) &"'"
   
  sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
  sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
  sql = sql& ",LogETime = '" & Now() &"'" 
  sql = sql& " WHERE KeyID='"&ID&"' " ':Response.Write sql
  Call rs_DoSql(conn,sql)
  sMsg = "修改成功!"
  fEdit = "selected"
ElseIf yAct="Eing" Then
  ID = Request("ID")
  fEdit = "selected"
  sMsg = "请修改保存!"
ElseIf yAct="Del" Then
  ID = Request("ID")
  Call rs_DoSql(conn,"DELETE FROM InfoOuter WHERE KeyID='"&ID&"'")
  sMsg = "删除成功!"
End If





%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" src="../../inc/home/jsInfo.js"></script>
<style type="text/css">
.iBox {
	border:1px solid #EEEEEE;
	padding:0px 0px 0px 0px;
	margin:0px 2px 5px 2px;
}
</style>
</head>
<body>
<%



%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:1px solid #CCC; margin-top:2px;">
  <tr>
    <th width="15%" nowrap bgcolor="#F0F0F0">采集方案</th>
    <th nowrap bgcolor="#F0F0F0">信息类别</th>
    <th nowrap bgcolor="#F0F0F0">Url</th>
    <th width="10%" nowrap bgcolor="#F0F0F0">图文</th>
    <th width="10%" nowrap bgcolor="#F0F0F0">编码</th>
    <th width="5%" nowrap bgcolor="#F0F0F0">采集</th>
    <th width="5%" nowrap bgcolor="#F0F0F0">Edit</th>
    <th width="5%" nowrap bgcolor="#F0F0F0">Del</th>
  </tr>
  <%
 sqlmod = "'"&ModID&"'" :If Request("ModPub")="Module" Then sqlmod = "'Module','"&ModID&"'" 
 sql = " SELECT * FROM [InfoOuter] WHERE KeyMod IN("&sqlmod&") "
 If ID<>"" Then sql =sql& " OR KeyID='"&ID&"'" 
 sql =sql& " ORDER BY KeyID DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod") 
InfSubj = rs("InfSubj")
InfType = rs("InfType") 
InfTyp2 = rs("InfTyp2")
InfUrl = rs("InfUrl")
InfCSet = rs("InfCSet")
InfP1Str = rs("InfP1Str")  
InfP1 = rs("InfP1")
InfP2Str = rs("InfP2Str")  
InfP2 = rs("InfP2")
InfDel1 = rs("InfDel1")
InfDel2 = rs("InfDel2")
InfQ1Str = rs("InfQ1Str") 
InfQ1 = rs("InfQ1")
InfQ2Str = rs("InfQ2Str") 
InfQ2 = rs("InfQ2")
InfRep1 = rs("InfRep1")
InfRep2 = rs("InfRep2")
InfKill = rs("InfKill")
If ID=KeyID Then
tMod = KeyMod
tSubj = InfSubj
tType = InfType 
tTyp2 = InfTyp2
tUrl = InfUrl
tCSet = InfCSet
tP1Str = InfP1Str :tP1Str = Show_Form(tP1Str)
tP1 = InfP1
tP2Str = InfP2Strr :tP2Str = Show_Form(tP2Str)
tP2 = InfP2
tDel1 = InfDel1
tDel2 = InfDel2
tQ1Str = InfQ1Str :tQ1Str = Show_Form(tQ1Str)
tQ1 = InfQ1
tQ2Str = InfQ2Str :tQ2Str = Show_Form(tQ2Str)
tQ2 = InfQ2
tRep1 = InfRep1 :tRep1 = Show_Form(tRep1)
tRep2 = InfRep2 :tRep2 = Show_Form(tRep2)
tKill = InfKill
End If
If InfP1Str="" Then InfP1Str="<body"
If InfP2Str="" Then InfP2Str="</body"
If InfQ1Str="" Then InfQ1Str="<body"
If InfQ2Str="" Then InfQ2Str="</body"
If KeyMod<>"Module" Then
TypName = Get_TypeName(ModID,InfType)
Else
TypName = "<span class='col999'>[公用规则]</span>"
End If
'InfPath = rs("InfPath")

%>
  <tr>
    <td nowrap bgcolor="#FFFFFF"><%=InfSubj%></td>
    <td nowrap bgcolor="#FFFFFF"><%=TypName%></td>
    <td nowrap bgcolor="#FFFFFF"><a href="<%=InfUrl%>" target="_blank"><%=Left(InfUrl,42)%></a></td>
    <td align="center" nowrap bgcolor="#FFFFFF"><%=InfTyp2%></td>
    <td align="center" nowrap bgcolor="#FFFFFF"><%=InfCSet%></td>
    <td align="center" nowrap bgcolor="#FFFFFF">
    <a href="imp_do.asp?ID=<%=KeyID%>&ModID=<%=ModID%>&yAct=(Def)" target="_blank">采集</a>
    </td>
    <td align="center" nowrap bgcolor="#FFFFFF"><a href="?ID=<%=KeyID%>&yAct=Eing">Edit</a></td>
    <td align="center" nowrap bgcolor="#FFFFFF">
    <%If KeyMod=ModID Then%>
    <a href="?ID=<%=KeyID%>&yAct=Del">Del</a>
    <%Else%>
    <span class="col999">Del</span>
    <%End If%>
    </td>
  </tr>
  <%
 rs.MoveNext()
 Loop
 Set rs = Nothing
%>
</table>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" style="border:1px solid #CCC; margin-top:2px;">
  <tr>
    <td style="padding-left:60px;"><div style="float:right"> 
    <a href="?ModID=<%=ModID%>">重载本页</a> 
    <a href="bat_up.asp?ModID=<%=ModID%>">批量上传</a> 
    <a href="../../tools/out/outinfo.asp">单页分析</a> 
    </div>
      <strong> <%=ModName%> 采集规则设置</strong> <%=bakID%></td>
    <td width="32%" rowspan="4" align="left" valign="top"><div style="width:260px; height:280px; overflow-y:scroll; padding:5px; background-color:#F0F0F0"> 说明： <br>
        <div id="msgBox" style="padding:1px; background-color:#FFFFCC"><%=sMsg%></div>
        <!--#include file="imp_read.asp"--> 
    </div></td>
  </tr>
  <tr>
    <td valign="top"><table width="550" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
        <form id="fm01" name="fm01" method="post" action="?">
          <tr>
            <td align="center" bgcolor="#FFFFFF">Name:</td>
            <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=tSubj%>" size="23" maxlength="24" onFocus="iClear()" />
              <select name="InfType" id="InfType" style="width:150px" onFocus="iClear()">
                <%=Get_TypeOpt(ModID,tType)%>
              </select>
              <select name="InfTyp2" id="InfTyp2" onFocus="iClear()">
                <%=Get_SOpt("TLink;PLink;1Link","文字连接;图片连接;单页采集",tTyp2,"")%>
              </select></td>
          </tr>
          <tr>
            <td width="15%" align="center" bgcolor="#FFFFFF">Url: </td>
            <td bgcolor="#FFFFFF"><input name="InfUrl" type="text" id="InfUrl" value="<%=tUrl%>" size="48" maxlength="120" onFocus="iMsg(51)" />
              <select name="InfCSet" id="InfCSet" onFocus="iClear()">
                <%=Get_SOpt("UTF-8;gb2312;big5;iso-8859-1","",tCSet,"")%>
              </select></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Pos1:</td>
            <td bgcolor="#FFFFFF"><input name="InfP1Str" type="text" id="InfP1Str" value="<%=tP1Str%>" size="48" maxlength="48" onFocus="iMsg(11)" />
              <input name="InfP1" type="text" id="InfP1" value="<%=tP1%>" size="6" onFocus="iMsg(12)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Pos2:</td>
            <td bgcolor="#FFFFFF"><input name="InfP2Str" type="text" id="InfP2Str" value="<%=tP2Str%>" size="48" maxlength="48" onFocus="iMsg(13)" />
              <input name="InfP2" type="text" id="InfP2" value="<%=tP2%>" size="6" onFocus="iMsg(14)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Replace.:</td>
            <td nowrap="nowrap" bgcolor="#FFFFFF">
              <input name="InfRep1" type="text" id="InfRep1" value="<%=tRep1%>" size="48" maxlength="120" onFocus="iMsg(21)" />
              <input name="InfDel1" type="text" id="InfDel1" value="<%=tDel1%>" size="6" onFocus="iMsg(22)" />
              <br />
              <input name="InfRep2" type="text" id="InfRep2" value="<%=tRep2%>" size="48" maxlength="120" onFocus="iMsg(23)" />
              <input name="InfDel2" type="text" id="InfDel2" value="<%=tDel2%>" size="6" onFocus="iMsg(24)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Kill Strs:</td>
            <td bgcolor="#FFFFFF"><input name="InfKill" type="text" id="InfKill" value="<%=tKill%>" size="48" maxlength="120" onFocus="iMsg(31)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Flag1:</td>
            <td bgcolor="#FFFFFF"><input name="InfQ1Str" type="text" id="InfQ1Str" value="<%=tQ1Str%>" size="48" maxlength="120" onFocus="iMsg(41)" />
              <input name="InfQ1" type="text" id="InfQ1" value="<%=tQ1%>" size="6" onFocus="iMsg(42)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">Flag2:</td>
            <td bgcolor="#FFFFFF"><input name="InfQ2Str" type="text" id="InfQ2Str" value="<%=tQ2Str%>" size="48" maxlength="120" onFocus="iMsg(43)" />
              <input name="InfQ2" type="text" id="InfQ2" value="<%=tQ2%>" size="6" onFocus="iMsg(44)" /></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF">保存</td>
            <td bgcolor="#FFFFFF"><div style="float:right">
            <a href="?ModID=<%=ModID%>">重载本页</a>
            <a href="?ModID=<%=ModID%>&ModPub=Module">公用规则</a>
                <input type="button" name="btmSend" id="btmSend" value="确认提交" onClick="chkData()">
              </div>
              <select name="yAct" id="yAct" onFocus="iClear()">
                <option value="Add"  >Add -新增</option>
                <%If KeyMod=ModID Then%>
                <option value="Edit" <%=fEdit%>>Edit-修改</option>
                <%End If%>
                <option value="Copy" >Copy-复制</option>
              </select>
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
              <!--a href="imp_demo.asp">设置示例</a--></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<script type="text/javascript">
var Si = "<%=Si%>"; 
var Ai = Si.split("|"); 
var Ni=0; 
var fm = document.fm01;

function iClear()
{
  getElmID("msgBox").innerHTML = "";
}
function iMsg(n)
{
  var msg = new Array();
  msg[11] = "连接列表信息块 开始标记字符, 注意唯一性; ";
  msg[12] = "连接列表信息块 开始标记字符 偏移量, 可以为负数; ";
  msg[13] = "连接列表信息块 结束标记字符, 注意唯一性; ";
  msg[14] = "连接列表信息块 结束标记字符 偏移量, 可以为负数; ";
  msg[21] = "被替换的目标字符串, 多个用 <span class='col000'>|</span> 分开; ";
  msg[22] = "连接列表 前面忽略的连接个数; ";
  msg[23] = "替换后显示的字符串, 多个用 <span class='col000'>|</span> 分开; ";
  msg[24] = "连接列表 后面忽略的连接个数; ";
  msg[31] = "要删除的目标字符串, 多个用 <span class='col000'>|</span> 分开; ";
  msg[41] = "详细内容 开始标记字符, 注意唯一性; ";
  msg[42] = "详细内容 开始标记字符 偏移量, 可以为负数; ";
  msg[43] = "详细内容 结束标记字符, 注意唯一性; ";
  msg[44] = "详细内容 结束标记字符 偏移量, 可以为负数; ";
  msg[51] = "连接列表信息块地址; 单页采集则为单页的域名和目录名; ";
  getElmID("msgBox").innerHTML = " &nbsp; "+msg[n]; 
}

function chkData()
{
       
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fm.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     fm.InfSubj.focus();
     eflag = 1; break;
   }
   
 if (fm.InfUrl.value.length==0) 
   {   
     alert(" Url 不能为空！"); 
	 fm.InfUrl.focus();
     eflag = 1; break;
   } 

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 fm.submit();
		 }
}

</script>
</body>
</html>
