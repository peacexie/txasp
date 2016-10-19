<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->

<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<!--#include file="config.asp"-->
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>

<%
 		 
send = Request("send")
ReEnd = Request("ReEnd")
TP = RequestS("InfType",C,48)

If send="ins" Then 

InfSubj = RequestS("InfSubj",C,96)
InfURL = RequestS("InfURL",C,255) 

SetSAdm = "Y"
sql = " INSERT INTO [GboLink] (" 
sql = sql& "  KeyID,ImgName" 
sql = sql& ", InfType,InfSubj,InfURL" 
sql = sql& ", SetSubj,SetShow,SetTop"
sql = sql& ", LogAddIP,LogAUser,LogATime" 
sql = sql& ", LogEditIP,LogEUser,LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & Get_AutoID(24) &"','" & RequestS("ImgNam3",C,255) &"'" 
sql = sql& ", '" & TP &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfURL &"'" 
sql = sql& ", '" & RequestS("SetSubj",C,48) &"', '" & RequestS("SetShow",C,2) & "', '" & RequestS("SetTop","N",888) &"'" 
sql = sql& ", '" & Get_CIP() &"', '" & Session("UsrID") &"', '" & Now() &"'" 
sql = sql& ", '-', '-', '1900-12-31'" 
sql = sql& ")"
If Len(InfSubj)>0 And Len(InfURL)>6 Then
  Call rs_DoSql(conn,sql) 
  jsMsg = "信息填写成功！"
Else
  jsMsg = "信息填写失败！"
End If
  If ReEnd="Y" Then
    Response.Write js_Alert(jsMsg,"Redir","info_add.asp?KW="&KW&"&PG="&PG&"&InfType="&TP&"&ReEnd="&ReEnd&"") 
  Else
    Response.Write js_Alert(jsMsg,"Redir","info_list.asp?KW="&KW&"&PG="&PG&"&InfType="&TP&"&ReEnd="&ReEnd&"")   
  End If
End If



%>
<table width="640" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td width="80" align="center" nowrap>&nbsp;</td>
      <td align="right"><strong>连接增加</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">网站名称</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="24">
&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">类别属性</td>
      <td bgcolor="#FFFFFF"><select id=InfType name=InfType style="width:120; ">
        <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomLnk1' ORDER BY TypID",TP)%>
      </select>
      &nbsp;&nbsp;颜色
      <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90px ">
        <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>-[默认]-</OPTION>
        <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
        <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
        <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
        <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
        <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
        <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
      </select>
      &nbsp;&nbsp;顺序
      <input name="SetTop" type="text" id="SetTop" value="888" size="6" maxlength="24"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">目标网址</td>
      <td bgcolor="#FFFFFF"><input name="InfURL" type="text" id="InfURL" value="http://" size="60" maxlength="120">      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">图片地址</td>
      <td bgcolor="#FFFFFF"><input name="ImgNam3" type="text" id="ImgNam3" value="" size="60" maxlength="120">
        <input name=view2 type=button id="Button2" value="选择" onClick="window.open('../file/file_view.asp?yPath=myfile/link/')">
        <input name="InfSpeci" type="hidden" id="InfSpeci" value="<%=InfSpeci%>" size="12" maxlength="12">
      <input name="InfPrice" type="hidden" id="InfPrice" value="<%=InfPrice%>" size="12" maxlength="10"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">返回</td>
      <td bgcolor="#FFFFFF">
      <input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
      添加资料后返回列表
      &nbsp;&nbsp;&nbsp;
      <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
      添加资料后继续</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
          <input name=view type=reset id="Button1" value="重    写">
          <input name="send" type="hidden" id="send" value="ins"></td>
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
