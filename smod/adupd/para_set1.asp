<!--#include file="config.asp"-->
<!doctype html>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../sadm/func1/WinFunc.js" type="text/javascript"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jsInfo.js"></script>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtAct=mainLoad"></script>

</head>
<body>
<%

send = Request("send")
ID = RequestS("ID","C",24) 
Flag = RequestS("Flag","C",24) 'Txt,Num,Rem,Editor
nLen = RequestS("nLen","N",1)
fRet = Request("fRet")
'Response.Write fRet

If send = "Edit" Then
  sql = " UPDATE [AdmPara] SET " 
  If Flag="Rem" Or Flag="Editor" Then
    sql = sql& " ParRem = '" & RequestS("ParVal",3,480123) &"'" 
  ElseIf Flag="Num" Then
    sql = sql& " ParNum = " & RequestS("ParVal","N",1) &"" 
  Else
    sql = sql& " ParText = '" & RequestS("ParVal",3,nLen*2) &"'" 
  End If
  sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("UsrID")&"',LogETime='"&Now()&"' "
  sql = sql& " WHERE ParCode='"&ID&"' "
  Call rs_DoSql(conn,sql)
  msg = ""&ID&" 修改成功,请刷新!"
  If fRet="List" Then
    Response.Write js_Alert(msg,"Redir","para_list.asp")
	'Response.Write fRet
  End If
End If

sql = "SELECT * FROM [AdmPara] WHERE ParCode='" & ID &"' "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
If NOT rs.EOF Then
  ParCode = rs("ParCode")
  ParName = rs("ParName")
  ParFlag = rs("ParFlag")
  ParText = Show_Form(rs("ParText"))
  ParNum  = rs("ParNum")
  ParRem  = rs("ParRem") 'Show_Form()
End If
rs.Close()
SET rs=Nothing


%>
<div style="line-height:8px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="50" align="center" valign="top" bgcolor="#FFFFFF">常用</td>
    <td align="left" valign="top" bgcolor="#FFFFFF"><A href="?ID=intSpace&Flag=Num&nLen=6">空间大小</A> - <A href="?ID=seo_Home&Flag=Rem&nLen=12">首页优化</A> - <A href="?ID=rCopy.htm&Flag=Editor&nLen=12">版权备注</A> - <A href="?ID=rRead.txt&Flag=Editor&nLen=12">免责声明</A> - <A href="?ID=FilKW024&Flag=Rem&nLen=12">过滤字符</A> - <A href="para_list.asp"><span class="colF00">更多&gt;&gt;</span></A></td>
  </tr>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <form name="fmedit0" method="post" action="?">
    <tr>
      <td width="50%" bgcolor="#FFFFFF">参数名称：<%=ParName%></td>
      <td width="50%" bgcolor="#FFFFFF">参数编码：<%=ParCode%></td>
    </tr>
    <%If Flag="Rem" Then%>
    <tr>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><fieldset style="padding:5px;">
          <legend> 参数值 </legend>
          <textarea name="ParVal" cols="72" rows="<%=nLen%>"  wrap='Off' id="ParVal" ><%=ParRem%></textarea>
        </fieldset></td>
    </tr>
    <%ElseIf Flag="Text" Then%>
    <tr>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><fieldset style="padding:5px;">
          <legend> 参数值 <font color="#FF0000"><%=Msg%></font> </legend>
          <input name="ParVal" type="text" id="ParVal" value="<%=ParText%>" size="80" maxlength="<%=nLen%>">
        </fieldset></td>
    </tr>
    <%ElseIf Flag="Editor" Then%>
    <tr>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><fieldset style="padding:5px;">
          <legend> 参数值 </legend>
<textarea id="ParVal" name="ParVal" style="width:600px;height:<%=20*nLen%>px;visibility:hidden;display:none"><%=Show_Form(ParRem)%></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID01&edtCont=ParVal&edtTool=Basic"></script>

        </fieldset></td>
    </tr>
    <%Else%>
    <tr>
      <td colspan="2" align="center" bgcolor="#FFFFFF"><fieldset style="padding:5px;">
          <legend> 参数值 <font color="#FF0000"><%=Msg%></font> </legend>
          <input name="ParVal" type="text" id="ParVal" value="<%=ParNum%>" size="80" maxlength="<%=nLen%>">
        </fieldset></td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" bgcolor="#FFFFFF"><input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="Flag" type="hidden" id="Flag" value="<%=Flag%>">
        <input name="nLen" type="hidden" id="nLen" value="<%=nLen%>">
        <input name="send" type="hidden" id="send" value="Edit">
        <input name="fRet" type="hidden" id="fRet" value="<%=fRet%>">
	    <%If Flag="Editor" Then%>
        <input type="button" name="Submit" value="保存" onClick="ChkRem('0')">
        <%Else%>
        <input type="submit" name="Submit" value="保存" >
        <%End If%>
      </td>
      <td align="center" bgcolor="#FFFFFF">
        <A href="../../sadm/system/upd_para.asp">刷新</A>
<%If fRet="List" Then%>
      | <A href="para_list.asp">返回</A>
<%ElseIf fRet="Home" Then%>    
      | <A href="../../sadm/user/index.asp">返回</A>
<%ElseIf fRet="Close" Then%>    
      | <A style="cursor:pointer;" onClick="javascript:window.close()">关闭</A>
<%Else%>    
      | <A class="col999">关闭</A>
<%End If%>
      </td>
    </tr>
  </form>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="50" align="center" valign="top" bgcolor="#FFFFFF"><strong>注意</strong></td>
    <td align="left" valign="top" bgcolor="#FFFFFF">*. 修改成功后，刷新生效; 
<%
If Flag="Num" Then
  Response.Write "<br>*. 此项填写数字; "
End If
If ParFlag="ParSEO" Then
  Response.Write "<br>*. a: 所有内容请填在括号内; <br> &nbsp; b: 括号内不能有半角的括号,引号,换行等特殊字符; <br> &nbsp; c: 不按要求填写,后果自负; "
End If
If ParFlag="ParFil" Then
  Response.Write "<br>*. a: 关键词用分号(;)分开; <br> &nbsp; b: 关键词至少两个字符; "
End If
%>

    </td>
  </tr>
</table>
<br>
<script type="text/javascript">

function ChkRem(fmID)
{
  var fmObj = eval("document.fmedit"+fmID);
  fmObj.ParVal.value = apiGetValEditID01(); 
  fmObj.submit();
  
}

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Srart For ////////////////
 if (document.fm01.ParCode.value.length==0) 
   {   
     alert(" 参数编码 不能为空！"); 
     document.fm01.ParCode.focus();
     eflag = 1; break;
   }
 if (document.fm01.ParName.value.length==0) 
   {   
     alert(" 参数名称 不能为空！"); 
     document.fm01.ParName.focus();
     eflag = 1; break;
   }
 if (document.fm01.ParText.value.length==0) 
   {   
     alert(" 参数值 不能为空！"); 
     document.fm01.ParText.focus();
     eflag = 1; break;
   }
   //tmv = chkF_Mail(document.fm1.XXXXXX,"");
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
