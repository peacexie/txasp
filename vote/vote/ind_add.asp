<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../sadm/edfck/fckeditor.js"></script>
</head>
<body>
<%

send = Request("send")

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="ind_add2.asp" enctype="multipart/form-data" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
      <td align="center" bgcolor="#FFFFFF"><strong>[问卷调查]增加</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">主题</td>
      <td bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" size="60" maxlength="120">      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">类别</td>
      <td bgcolor="#FFFFFF"><select name="InfType" id="InfType" style="width:165px;">
        <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay")%>
      </select>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;编号
      <input name="KeyCode" type="text" id="KeyCode" value="<%=rs_AutID(conn,"VoteInfo","KeyCode","Vot-","1","")%>" size="24" maxlength="24"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF">问卷须知<br>
      <input type="hidden" name="InfRem1" id="InfRem1"></td>
      <td bgcolor="#FFFFFF">
<script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID01' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 360 ;
oFCKeditor.Value	= '<%=Show_jsStr(InfRem1)%>' ;
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
	  </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">问卷日期</td>
      <td bgcolor="#FFFFFF"><input name="InfTime1" type="text" id="InfTime1" value="<%=Date()%>" size="14" maxlength="10">  
        ~ 
          <input name="InfTime2" type="text" id="InfTime2" value="<%=Date()%>" size="14" maxlength="10">        
        &nbsp;(格式:2008-12-31)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">验证规则</td>
      <td bgcolor="#FFFFFF"><select name="InfCard" id="InfCard" style="width:120px;">
        <option value='N'>无身份证号验证</option>
        <option value='Y'>有身份证号验证</option>
      </select>
        ~
        <select name="InfVNum" id="InfVNum" style="width:120px;">
          <option value='D1T1'>每IP每天可投1次</option>
          <option value='DNT1'>每IP共可问卷1次</option>
          <option value='DXTN'>[无限制] 问卷</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF">图片</td>
      <td bgcolor="#FFFFFF"><input type='file' name='ImgName' style="width:360px; ">      </td>
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
        <input name=view type=reset id="Button1" value="重    写">      </td>
    </tr>
  </form>
</table>
<script src="../../upfile/pics/htm/js<%=ModID%>.js" type="text/javascript"></script>

<script type="text/javascript">
document.fm01.InfType.options[1].selected=true;
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 
	document.fm01.InfRem1.value = FCKeditorAPI.GetInstance("EditID01").GetXHTML(true);
 if (document.fm01.InfRem1.value.length==0) 
   {   
     alert(" 内容 不能为空！"); 
     eflag = 1; break;
   }   

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}
//document.fm01.InfType.options[1].selected=true;
</script>
</body>
</html>
