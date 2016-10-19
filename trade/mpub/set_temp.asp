<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../sadm/func1/WinFunc.js" type="text/javascript"></script>
<script type="text/javascript" src="../../sadm/edfck/fckeditor.js"></script>
</head>
<body>
<%

send = Request("send")
ID = RequestS("ID","C",24) 
Flag = RequestS("Flag","C",24) 'Txt,Num,Rem,Editor
nLen = RequestS("nLen","N",1)
fRet = Request("fRet")
'Response.Write fRet
If ID=""   Then ID="TraJ124"
If Flag="" Then Flag="Editor"
If Request("nLen")="" Then nLen="18"

If send = "Edit" Then
 If Request("isAdd")="OK" Then
  sql = " INSERT INTO [TradePara] (" 
  sql = sql& "  ParCode,ParMod,ParFlag"  
  If Flag="Rem" Or Flag="Editor" Then
    sql = sql& ", ParRem" 
  ElseIf Flag="Num" Then
    sql = sql& ", ParNum" 
  Else
    sql = sql& ", ParText" 
  End If
  sql = sql& ",LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & rs_TraTID() &"','UserTemp','"&ID&"'" 
  If Flag="Rem" Or Flag="Editor" Then
    sql = sql& ", '" & RequestS("ParVal",3,8000) &"'" 
  ElseIf Flag="Num" Then
    sql = sql& ", " & RequestS("ParVal","N",1) &"" 
  Else
    sql = sql& ", '" & RequestS("ParVal",3,8000) &"'" 
  End If
  sql = sql& " ,'"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"'"
  sql = sql& ")"
 Else
  sql = " UPDATE [TradePara] SET " 
  If Flag="Rem" Or Flag="Editor" Then
    sql = sql& " ParRem = '" & RequestS("ParVal",3,48000123) &"'" 
  ElseIf Flag="Num" Then
    sql = sql& " ParNum = " & RequestS("ParVal","N",1) &"" 
  Else
    sql = sql& " ParText = '" & RequestS("ParVal",3,nLen*2) &"'" 
  End If
  sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("MemID")&"',LogETime='"&Now()&"' "
  sql = sql& " WHERE ParFlag='"&ID&"' AND LogAUser='"&Session("MemID")&"' "  
 End If
  Call rs_DoSql(conn,sql)
  msg = ""&ID&" 修改成功,请刷新!"
End If

sql = "SELECT * FROM [TradePara] WHERE ParFlag='" & ID &"' AND LogAUser='"&Session("MemID")&"' "
SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open sql,conn,1,1 
If NOT rs.EOF Then
  ParCode = rs("ParCode")
  ParMod = rs("ParMod")
  ParFlag = rs("ParFlag")
  'ParText = Show_Form(rs("ParText"))
  'ParNum  = rs("ParNum")
  ParRem  = rs("ParRem") 'Show_Form()
  isAdd = "NO"
Else
  ParRem  = "无资料，请添加..."
  isAdd = "OK"
End If
rs.Close()
SET rs=Nothing

ParName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ID&"'")

%>
<div style="line-height:12px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="120" align="center" valign="top" bgcolor="#FFFFFF">设置项目</td>
    <td align="left" valign="top" bgcolor="#FFFFFF">
     <A href="?ID=TraT124&Flag=Editor&nLen=18">供求模版</A> - 
     <A href="?ID=TraJ124&Flag=Editor&nLen=18">职位模版</A> - 
     <!--A href="?ID=Tra_Kil&Flag=Rem&nLen=18">过滤字符</A-->  
     <!--A href="?ID=Tra_Typ&Flag=Rem&nLen=18">类别刷新</A> -->
    </td>
    <!--td align="left" valign="top" bgcolor="#FFFFFF"><A href="?ID=intSpace&Flag=Num&nLen=6">空间大小</A> - <A href="?ID=seo_Home&Flag=Rem&nLen=12">首页优化</A> - <A href="?ID=rCopy.htm&Flag=Editor&nLen=12">版权备注</A> - <A href="?ID=rRead.txt&Flag=Editor&nLen=12">免责声明</A> - <A href="?ID=FilKW024&Flag=Rem&nLen=12">过滤字符</A> - <A href="para_list.asp"><span class="colF00">更多&gt;&gt;</span></A></td-->

  </tr>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <form name="fmedit0" method="post" action="?">
    <tr>
      <td width="50%" bgcolor="#FFFFFF">参数名称：<%=ParName%>模版</td>
      <td width="50%" bgcolor="#FFFFFF">参数编码：<%=ID%></td>
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
          <script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID_0' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= <%=20*nLen%> ;
oFCKeditor.Value	= '<%=Show_jsStr(ParRem)%>' ;
oFCKeditor.Create() ;
</script>
          <input name="ParVal" type="hidden" value="">
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
      <input name="isAdd" type="hidden" id="isAdd" value="<%=isAdd%>">
      | <A class="col999">刷新</A>
    
      | <A class="col999">关闭</A>

      </td>
    </tr>
  </form>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="50" align="center" valign="top" bgcolor="#FFFFFF"><strong>注意</strong></td>
    <td align="left" valign="top" bgcolor="#FFFFFF">*. 把经常要输入的信息，录入到模版，提高资料录入效率！ 
      <%
If Flag="Num" Then
  Response.Write "<br>*. 此项填写数字; "
End If
If ParFlag="ParSEO" Then
  Response.Write "<br>*. a: 所有内容请填在括号内; <br> &nbsp; b: 括号内不能有括号,引号,换行等特殊字符; <br> &nbsp; c: 不按要求填写,后果自负; "
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
  fmObj.ParVal.value = FCKeditorAPI.GetInstance("EditID_"+fmID).GetXHTML(true);
  //alert(fmObj.ParVal.value);
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
