<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../../sadm/edfck/fckeditor.js" type="text/javascript"></script>
<title>Depart Demo</title>
</head>
<body>

<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../upfile/sys/config/_depart.asp"-->

<%

 aCode = Split(strDepCode,"|")
 aName = Split(strDepName,"|")
 aFlag = Split(strDepFlag,"|")
 TreeOpt = "" : TreeArr = ""
 For i=0 To uBound(aCode)-1
    TreeOpt=TreeOpt&vbcrlf&"<option value='"&aName(i)&"'>"&aName(i)&"</option>"
    TreeArr=TreeArr&"new Array("&gdArr(aCode(i))&"),"&vbcrlf
 Next

Function gdArr(xID)
  Dim rs,s : s="'不限'" 
  Dim sql,KeyID,InfSubj
  sql = " SELECT * FROM [InfoNews] WHERE KeyMod='InfD124' AND InfType LIKE '%D110028%' AND InfTyp2='"&xID&"' " ' AND SetShow='Y'
  Set rs=Server.CreateObject("Adodb.Recordset")
  rs.Open Sql,conn,1,1 '("&rs.RecordCount&"人)
  If NOT rs.EOF Then
  Do While NOT rs.EOF
    KeyID = rs("KeyID")
    InfSubj = rs("InfSubj")
     s=s&",'"&InfSubj&"'"
  rs.MoveNext
  Loop
  End If
  Set rs=Nothing
  gdArr = s
End Function

%>


<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="right"><strong>信息增加</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>主题</td>
      <td><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120"></td>
    </tr>
    <tr>
      <td align="center">类别</td>
      <td><select name="InfType" style="width:120px; " id="InfType">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypID",InfType)%>
        </select>
        &nbsp; &nbsp;
        姓名
        <input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="24" maxlength="24" <%If ModID="Gbo936" Then Response.Write("readonly") %> ></td>
    </tr>
    <tr>
      <td align="center" nowrap>审核</td>
      <td><select name="SetShow" id="SetShow" style="width:120px;">
          <option value="Y"  <%if SetShow="Y" then Response.Write("selected")%>>已审核</option>
          <option value="N"  <%if SetShow="N" then Response.Write("selected")%>>未审核</option>
        </select>
        &nbsp; &nbsp; 邮箱
        <input name="LnkEmail" type="text" id="LnkEmail" value="<%=LnkEmail%>" size="24" maxlength="60"></td>
    </tr>

    <tr>
      <td align="center">部门</td>
      <td>
        <select name="TreeTypA" size=1 id="TreeTypA" style="width:120px;" onChange="TreeChange()">
          <option value="不限">选择部门</option>
          <%=TreeOpt%>
          <option value="不限">不限</option>
        </select>
        &nbsp; &nbsp; 医生
<select name="TreeTypB" size="1" id="TreeTypB" style="width:150px;">
          <option value="不限">选择医生</option>
        </select>
      </td>
    </tr>

    <tr>
      <td align="center">内容</td>
      <td><script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID01' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 240 ;
oFCKeditor.Value	= '<%=Show_jsStr(InfCont)%>' ;
oFCKeditor.ToolbarSet = 'Basic' 
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
        <input name="InfCont" type="hidden" value=""></td>
    </tr>

    <tr>
      <td align="center">回复</td>
      <td><script type="text/javascript">
var oFCKeditor = new FCKeditor( 'EditID02' ) ;
oFCKeditor.BasePath	= oFCKeditor.BasePath.replace("sadm/edfck/","<%=Config_Path%>sadm/edfck/"); 
oFCKeditor.Height	= 120 ;
oFCKeditor.Value	= '<%=Show_jsStr(InfReply)%>';
oFCKeditor.ToolbarSet = 'Basic' 
oFCKeditor.Create() ; //Server.HTMLEncode(InfCont)
</script>
        <input name="InfReply" type="hidden" value=""></td>
    </tr>

    <tr>
      <td align="center" nowrap><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
        注意:留言12K字内</td>
    </tr>
  </form>
</table>
<p>
  <script type="text/javascript">

var TreeForm = document.getElementById('fm01');
var TreeOptA = TreeForm.TreeTypA;
var TreeOptB = TreeForm.TreeTypB;
var TreeArr = new Array(
<%=TreeArr%>
new Array("不限")
);


document.getElementById('fm01').TreeTypA.value = "技术设计";
TreeChange();
document.getElementById('fm01').TreeTypB.value = "技术图片1";


function TreeChange(){
  index = TreeOptA.options.selectedIndex-1;
  if(index<0){
      for(var i=(TreeOptB.options.length-1);i>0;i--){TreeOptB.options.remove(i);}
	  return;
  };
  TreeOptB.length = TreeArr[index].length;
  for(var i = 0;i<TreeArr[index].length;i++)
  {
    var iVal = TreeArr[index][i];
    TreeOptB.options[i].text = iVal;
    TreeOptB.options[i].value = iVal;
  }
}

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

  document.fm01.InfCont.value = FCKeditorAPI.GetInstance("EditID01").GetXHTML(true);
  document.fm01.InfReply.value = FCKeditorAPI.GetInstance("EditID02").GetXHTML(true);

 if (document.fm01.InfCont.value.length>=12000) 
   {   
     alert(" 内容 不能超过 12K!"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }


         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script></p>
<hr />
<p>aDep = Split(KeyFlag&amp;&quot;^&quot;,&quot;^&quot;)<br />
  iDep = rs_Val("",&quot;SELECT SysID FROM AdmSyst WHERE SysType='Depart' AND SysName='&quot;&amp;aDep(0)&amp;&quot;'&quot;)<br />
  sDep = Replace(KeyFlag,&quot;^&quot;,&quot; / &quot;)<br />
  'If inStr(Session(UsrPStr),&quot;(&quot;&amp;iDep&amp;&quot;)&quot;)&gt;0 Then<br />
  If inStr(Session(UsrPStr),&quot;{Admin}&quot;)&gt;0 Or SwhDepSubs&lt;&gt;&quot;Y&quot; Or (SwhDepSubs=&quot;Y&quot; AND inStr(Session(UsrPStr),&quot;(&quot;&amp;iDep&amp;&quot;)&quot;)&gt;0) Then<br />
fDep = &quot;OK&quot;<br />
Else<br />
fDep = &quot;NG&quot;<br />
End If </p>
<p>dd</p>
</body>
</html>
