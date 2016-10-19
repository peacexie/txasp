<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->
<!--#include file="../../upfile/sys/config/MTypList.asp" -->
<%
Call Chk_Perm1("SysTools","")
Set rs=Server.Createobject("Adodb.Recordset")

tTyp = Trim(RequestS("tTyp",3,255)) :If tTyp="" Then tTyp="Index"
tKey = RequestS("tKey",3,48)
rDir = Request.Servervariables("HTTP_REFERER")

If tTyp="Index" Then
ElseIf tTyp="Search" Then
  s = RequestS("s",3,48) ': If s="" Then s="(Null)"
  sqlK = " WHERE KeyMod in('GboU124','GboU224') AND InfSubj LIKE '%"&s&"%' "
ElseIf tTyp="GboU124" Then
  sqlK = " WHERE KeyMod='GboU124' "
ElseIf tTyp<>"" Then
  sqlK = " WHERE KeyMod='GboU224' AND InfType='"&tTyp&"' AND LogAUser='"&Session("UsrID")&"' "
Else 
  'sqlK = " WHERE KeyMod='GboU124' "
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>站务笔记导出</title>
<style type="text/css">
body, td, th {
	font-size: 12px;
}
body, fieldset, div {
	margin: 5px;
}
.bgMenu {
	background-color:#DDF;
	padding:3px;
}
.bgRep {
	background-color:#DDD;
	padding:5px;
}
.bgCCC {
	background-color:#CCC;
	text-align:right;
	padding:5px;
}
.bgFFC {
	background-color:#FFC;
	padding:5px;
}
.highlight {
	background:green;
	font-weight:bold;
	color:white;
	padding:0px 8px;
}
</style>
</head>
<body>
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" class="bgMenu"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <form name="schdb" action="?">
            <td width="50%" align="left" bgcolor="#FFFFFF"><input name="s" type="text" id="s" onChange="n = 0;" size=8>
              <input type="submit" value="库表搜索">
              <input name="tTyp" type="hidden" id="tTyp" value="Search" /></td>
          </form>
          <form name="search" onSubmit="doSearch(this.s.value);return false;">
            <td width="50%" align="right" nowrap="nowrap" bgcolor="#FFFFFF"><input name="s" type="text" id="s" onChange="n = 0;" size=8>
              <input type="submit" value="页内查找"></td>
          </form>
        </tr>
        <tr>
          <td colspan="2" align="center"><a href="?">[Index]</a>
            <%    
aCode = Split(sGboU224Code,"|")
aName = Split(sGboU224Name,"|")
For i=0 To uBound(aCode)-1
  sID = aCode(i)
  sName = aName(i)
  If Left(sName,1)<>"~" Then
    Response.Write " - <a href='?tTyp="&sID&"'>"&sName&"</a>"
  End If
Next
%>
            - <a href="?tTyp=GboU124">[站务笔记]</a></td>
        </tr>
      </table></td>
  </tr>
</table>
<%If tKey<>"" Then%>
<%
 sql = " SELECT KeyID,InfSubj,InfCont,InfReply,LogATime FROM [GboInfo] WHERE KeyID='"&tKey&"' "
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 rs.Open Sql,conn,1,1
 If NOT rs.EOF Then
  KeyID = rs("KeyID")
  InfSubj = Show_Text(rs("InfSubj")) 
  xxxCont = Show_Html(rs("InfCont"))
  xxxReply = Show_Html(rs("InfReply"))
  LogATime = rs("LogATime")
 End If
 rs.Close
%>
<div align="center" style="width:700px; margin:auto; text-align:left">
  <fieldset style="border:1px solid #006">
    <legend> <span style="color:#00F; font-weight:bold"><%=InfSubj%></span> <%=LogATime%> <a href="<%=rDir%>">[返回]</a></legend>
    <div>
      <%Call Show_sfGbook(KeyID,".org.htm")%>
      <div class="bgFFC">
        <%Call Show_sfGbook(KeyID,".rep.htm")%>
      </div>
      <div class="bgCCC"> <a href="<%=rDir%>">[返回]</a> </div>
    </div>
  </fieldset>
</div>
<%ElseIf tTyp="Search" Then%>
<div align="center" style="width:700px; margin:auto; text-align:left">
  <fieldset style="border:1px solid #006">
    <legend> <span style="color:#00F; font-weight:bold">[数据库搜索]</span> </legend>
    <div>
      <%
 sql = " SELECT TOP 36 KeyID,InfSubj,LogATime FROM [GboInfo] "&sqlK
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 rs.Open Sql,conn,1,1
 j = 0
 Do While NOT rs.EOF
  j = j +1 
  KeyID = rs("KeyID")
  InfSubj = Show_Text(rs("InfSubj")) 
  LogATime = rs("LogATime")
  Response.Write vbcrlf&j&" - "&LogATime&" - <a href='?tKey="&KeyID&"'>"&InfSubj&"</a> <br>"
  rs.MoveNext()
 Loop
 rs.Close
%>
    </div>
  </fieldset>
</div>

<%ElseIf tTyp="Index" Then%>
<%    
aCode = Split(sGboU224Code,"|")
aName = Split(sGboU224Name,"|")
For i=0 To uBound(aCode)-1
  sID = aCode(i)
  sName = aName(i)
  If Left(sName,1)<>"~" Then
%>
<div align="center" style="width:700px; margin:auto; text-align:left">
  <fieldset style="border:1px solid #006">
    <legend> <span style="color:#00F; font-weight:bold"><%=i+1%> - <%=sName%></span> </legend>
    <div>
      <%
 sql = " SELECT KeyID,InfSubj,LogATime FROM [GboInfo] WHERE InfType='"&sID&"' "
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 rs.Open Sql,conn,1,1
 j = 0
 Do While NOT rs.EOF
  j = j +1 
  KeyID = rs("KeyID")
  InfSubj = Show_Text(rs("InfSubj")) 
  LogATime = rs("LogATime")
  Response.Write vbcrlf&j&" - "&LogATime&" - <a href='?tKey="&KeyID&"'>"&InfSubj&"</a> <br>"
  rs.MoveNext()
 Loop
 rs.Close
%>
    </div>
  </fieldset>
</div>
<%
  End If
Next
%>
<div align="center" style="width:700px; margin:auto; text-align:left">
  <fieldset style="border:1px solid #006">
    <legend> <span style="color:#00F; font-weight:bold"><%=i+1%> - [站务笔记]</span> </legend>
    <div>
      <%
 sql = " SELECT KeyID,InfSubj,LogATime FROM [GboInfo] WHERE KeyMod='GboU124' "
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 rs.Open Sql,conn,1,1
 j = 0
 Do While NOT rs.EOF
  j = j +1 
  KeyID = rs("KeyID")
  InfSubj = Show_Text(rs("InfSubj")) 
  LogATime = rs("LogATime")
  Response.Write vbcrlf&j&" - "&LogATime&" - <a href='?tKey="&KeyID&"'>"&InfSubj&"</a> <br>"
  rs.MoveNext()
 Loop
 rs.Close
%>
    </div>
  </fieldset>
</div>
<%Else%>
<%
 sql = " SELECT * FROM [GboInfo] "&sqlK
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 rs.Open Sql,conn,1,1
 i = 0
 Do While NOT rs.EOF
  i = i + 1
  KeyID = rs("KeyID")
  InfSubj = Show_Text(rs("InfSubj")) 
  xxxCont = Show_Html(rs("InfCont"))
  xxxReply = Show_Html(rs("InfReply"))
%>
<div align="center" style="width:700px; margin:auto; text-align:left">
  <fieldset style="border:1px solid #006">
    <legend> <span style="color:#00F; font-weight:bold"><%=i%> - <%=InfSubj%></span> </legend>
    <div>
      <%Call Show_sfGbook(KeyID,".org.htm")%>
      <div class="bgRep">
        <%Call Show_sfGbook(KeyID,".rep.htm")%>
      </div>
    </div>
  </fieldset>
</div>
<%
  rs.MoveNext()
 Loop
 rs.Close
%>
<%End If%>
<%
Set rs = Nothing
%>
<script type="text/javascript">
function encode(s){
  return s.replace(/&/g,"&").replace(/</g,"<").replace(/>/g,">").replace(/([\\\.\*\[\]\(\)\$\^])/g,"\\$1");
}
function decode(s){
  return s.replace(/\\([\\\.\*\[\]\(\)\$\^])/g,"$1").replace(/>/g,">").replace(/</g,"<").replace(/&/g,"&");
}
function doSearch(s){
  if (s.length==0){
    alert('搜索关键词未填写！');
    return false;
  }
  s=encode(s);
  var obj=document.getElementsByTagName("body")[0];
  var t=obj.innerHTML.replace(/<span\s+class=.?highlight.?>([^<>]*)<\/span>/gi,"$1");
  obj.innerHTML=t;
  var cnt=loopSearch(s,obj);
  t=obj.innerHTML
  var r=/{searchHL}(({(?!\/searchHL})|[^{])*){\/searchHL}/g
  t=t.replace(r,"<span class='highlight'>$1</span>");
  obj.innerHTML=t;
  alert("搜索到关键词"+cnt+"处")
}
function loopSearch(s,obj){
  var cnt=0;
  if (obj.nodeType==3){
    cnt=replace(s,obj);
    return cnt;
  }
  for (var i=0,c;c=obj.childNodes[i];i++){
    if (!c.className||c.className!="highlight")
      cnt+=loopSearch(s,c);
  }
  return cnt;
}
function replace(s,dest){
  var r=new RegExp(s,"g");
  var tm=null;
  var t=dest.nodeValue;
  var cnt=0;
  if (tm=t.match(r)){
    cnt=tm.length;
    t=t.replace(r,"{searchHL}"+decode(s)+"{/searchHL}")
    dest.nodeValue=t;
  }
  return cnt;
}
</script>
</body>
</html>
