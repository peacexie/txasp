<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->

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
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>

<%
 		
send = Request("send")
PG = RequestS("PG","N",1)
KW = RequestS("KW",3,48)
ID = RequestS("ID",3,48)
TP = RequestS("TP",3,48)
TP2 = RequestS("TP2",3,48)
TP3 = RequestS("TP3",3,48)
Img = RequestS("Img",3,48)
InfSubj = RequestS("InfSubj",C,96)
InfURL = RequestS("InfURL",C,255) 

If send="ins" Then 
sql = " UPDATE [GboLink] SET " 
sql = sql& " InfType = '" & RequestS("InfType",C,48) &"'" 
sql = sql& ",InfSubj = '" & RequestS("InfSubj",C,48) &"'" 
sql = sql& ",InfURL = '" & RequestS("InfURL",C,255) &"'" 
sql = sql& ",ImgName = '" & RequestS("ImgNam3",C,255) &"'" 
sql = sql& ",SetSubj = '" & RequestS("SetSubj",C,24) &"'" 
sql = sql& ",SetTop = '" & RequestS("SetTop","N",888) &"'" 
sql = sql& ",SetShow = '" & RequestS("SetShow",C,2) &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' " ':Response.Write sql : Response.End()
If Len(InfSubj)>0 And Len(InfURL)>0 Then
  Call rs_DoSql(conn,sql) 
End If
  jsMsg = "信息修改成功！" ': Response.Write sql
  Response.Write js_Alert(jsMsg,"Redir","info_list.asp?KW="&KW&"&PG="&PG&"&TP="&TP&"&TP2="&TP2&"&TP3="&TP3&"") 
End If


SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [GboLink] WHERE KeyID='"&ID&"'",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
InfType = rs("InfType")
InfSubj = rs("InfSubj")
InfURL = rs("InfURL")
ImgName = rs("ImgName")
SetShow = rs("SetShow")
SetTop = rs("SetTop")
SetSubj = rs("SetSubj")
end if 
rs.Close()
SET rs=Nothing 

%>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
      <td align="right" bgcolor="#FFFFFF"><strong>信息编辑</strong></td>
    </tr>
	<!--
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">行业</td>
      <td bgcolor="#FFFFFF"><select id=InfField name=InfField style="width:120; ">
  <option value='LnkTalent' <%If InfField="LnkTalent" Then Response.Write("selected")%>>人才类链接</option>
  <option value='LnkElse'   <%If InfField="LnkElse"   Then Response.Write("selected")%>>其它文字链接</option>
  <option value='SchBaidu'  <%If InfField="SchBaidu"  Then Response.Write("selected")%>>百度人才搜索</option>
  <option value='SchGoogle' <%If InfField="SchGoogle" Then Response.Write("selected")%>>谷歌人才搜索</option>
        </select>
&nbsp;（自定义分类）</td>
    </tr>
	-->
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">网站名称</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="24">
&nbsp;      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">类别属性</td>
      <td bgcolor="#FFFFFF"><select id=InfType name=InfType style="width:120px; ">
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomLnk1' ORDER BY TypID",InfType)%>
      </select>
&nbsp;&nbsp;颜色
<select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90px ">
  <OPTION style='COLOR:#<%=SetSubj%>; BACKGROUND-COLOR:#<%=SetSubj%>' value='<%=SetSubj%>'>#<%=SetSubj%></OPTION>
  <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>#000000</OPTION>
  <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
  <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
  <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
  <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
  <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
  <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
</select>
&nbsp;&nbsp;顺序
<input name="SetTop" type="text" id="SetTop" value="<%=SetTop%>" size="6" maxlength="24"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">目标网址</td>
      <td bgcolor="#FFFFFF"><input name="InfURL" type="text" id="InfURL" value="<%=InfURL%>" size="60" maxlength="120">      </td>
    </tr>
	<!--
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">图片地址</td>
      <td bgcolor="#FFFFFF"><input name="ImgName" type="text" id="ImgName" value="<%=ImgName%>" size="60" maxlength="120">
        <br>
        图片连接填写 此项,可先在<a href="/smod/adpic/adpic_add.asp">浮动图</a>增加本地图片；          </td>
    </tr>
	-->
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">图片地址</td>
      <td bgcolor="#FFFFFF"><input name="ImgNam3" type="text" id="ImgNam3" value="<%=ImgName%>" size="60" maxlength="120">
        <input name=view2 type=button id="Button2" value="选择" onClick="window.open('../file/file_view.asp?yPath=myfile/link/')">
        <input name="InfSpeci" type="hidden" id="InfSpeci" value="<%=InfSpeci%>" size="12" maxlength="12">
      <input name="InfPrice" type="hidden" id="InfPrice" value="<%=InfPrice%>" size="12" maxlength="10"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
          <input name=view type=reset id="Button1" value="重    写">
          <input name="send" type="hidden" id="send" value="ins">
          <input name="ID" type="hidden" id="ID" value="<%=ID%>">
          <input name="TP" type="hidden" id="TP" value="<%=TP%>">
          <input name="TP2" type="hidden" id="TP2" value="<%=TP2%>">
          <input name="TP3" type="hidden" id="TP3" value="<%=TP3%>">
          <input name="KW" type="hidden" id="KW" value="<%=KW%>">
          <input name="PG" type="hidden" id="PG" value="<%=PG%>"></td>
    </tr>
  </form>
</table>

<script type="text/javascript">

  var xFile,xSize;
    function owFileGet()
    {
	  var tFile = xFile; 
	  var nPos = tFile.indexOf('/upfile/');
	  var sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  document.fm01.ImgNam3.value = tFile;
	  nPos = tFile.indexOf('.');
	  sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  document.fm01.InfSpeci.value = tFile;
	  document.fm01.InfPrice.value = xSize;
    }

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 网站名称 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }

 if (document.fm01.InfURL.value.length<12) 
   {   
     //alert(" 目标网址 不正确！"); 
	 //document.fm01.InfURL.focus();
     //eflag = 1; break;
   }
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>

</body>
</html>
