<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../himg/tconfig.asp"-->
<!--#include file="cfunc.asp"-->
<!--#include file="safe.asp"-->
<%
Call Chk_Perm1(xPara,"")

Function Url2CAll(xStr)
Dim s : s = xStr

s = Replace(s,"%30","0")
s = Replace(s,"%31","1")
s = Replace(s,"%32","2")
s = Replace(s,"%33","3")
s = Replace(s,"%34","4")
s = Replace(s,"%35","5")
s = Replace(s,"%36","6")
s = Replace(s,"%37","7")
s = Replace(s,"%38","8")
s = Replace(s,"%39","9")

s = Replace(s,"%41","A")
s = Replace(s,"%42","B")
s = Replace(s,"%43","C")
s = Replace(s,"%44","D")
s = Replace(s,"%45","E")
s = Replace(s,"%46","F")
s = Replace(s,"%47","G")
s = Replace(s,"%48","H")
s = Replace(s,"%49","I")
s = Replace(s,"%4A","J")
s = Replace(s,"%4B","K")
s = Replace(s,"%4C","L")
s = Replace(s,"%4D","M")
s = Replace(s,"%4E","N")
s = Replace(s,"%4F","O")
s = Replace(s,"%50","P")
s = Replace(s,"%51","Q")
s = Replace(s,"%52","R")
s = Replace(s,"%53","S")
s = Replace(s,"%54","T")
s = Replace(s,"%55","U")
s = Replace(s,"%56","V")
s = Replace(s,"%57","W")
s = Replace(s,"%58","X")
s = Replace(s,"%59","Y")
s = Replace(s,"%5A","Z")

s = Replace(s,"%61","a")
s = Replace(s,"%62","b")
s = Replace(s,"%63","c")
s = Replace(s,"%64","d")
s = Replace(s,"%65","e")
s = Replace(s,"%66","f")
s = Replace(s,"%67","g")
s = Replace(s,"%68","h")
s = Replace(s,"%69","i")
s = Replace(s,"%6A","j")
s = Replace(s,"%6B","k")
s = Replace(s,"%6C","l")
s = Replace(s,"%6D","m")
s = Replace(s,"%6E","n")
s = Replace(s,"%6F","o")
s = Replace(s,"%70","p")
s = Replace(s,"%71","q")
s = Replace(s,"%72","r")
s = Replace(s,"%73","s")
s = Replace(s,"%74","t")
s = Replace(s,"%75","u")
s = Replace(s,"%76","v")
s = Replace(s,"%77","w")
s = Replace(s,"%78","x")
s = Replace(s,"%79","y")
s = Replace(s,"%7A","z")

Url2CAll = Url2Char(s)
End Function

Function Url2Char(xStr)
Dim s : s = xStr
s = Replace(s,"%20"," ")
s = Replace(s,"%21","!")
s = Replace(s,"%22",chr(34)) '"
s = Replace(s,"%23","#")
s = Replace(s,"%24","$")
s = Replace(s,"%25","%")
s = Replace(s,"%26","&")
s = Replace(s,"%27","'")
s = Replace(s,"%28","(")
s = Replace(s,"%29",")")
s = Replace(s,"%2A","*")
s = Replace(s,"%2B","+")
s = Replace(s,"%2C",",")
s = Replace(s,"%2D","-")
s = Replace(s,"%2E",".")
s = Replace(s,"%2F","/")
s = Replace(s,"%3A",":")
s = Replace(s,"%3B",";")
s = Replace(s,"%3C","<")
s = Replace(s,"%3D","=")
s = Replace(s,"%3E",">")
s = Replace(s,"%3F","?")
s = Replace(s,"%40","@")
s = Replace(s,"%5B","[")
s = Replace(s,"%5C","\")
s = Replace(s,"%5D","]")
s = Replace(s,"%5E","^")
s = Replace(s,"%5F","_")
s = Replace(s,"%60","`")
s = Replace(s,"%7B","{")
s = Replace(s,"%7C","|")
s = Replace(s,"%7D","}")
s = Replace(s,"%7E","~")
s = Replace(s,"%7F","")
Url2Char=s
End Function

Function Char2Code(xStr)
Dim s : s = xStr
s = Replace(s," ","&#32;")
s = Replace(s,"!","&#33;")
s = Replace(s,chr(34),"&#34;") '"
s = Replace(s,"#","&#35;")
s = Replace(s,"$","&#36;")
s = Replace(s,"%","&#37;")
s = Replace(s,"&","&#38;")
s = Replace(s,"'","&#39;")
s = Replace(s,"(","&#40;")
s = Replace(s,")","&#41;")
s = Replace(s,"*","&#42;")
s = Replace(s,"+","&#43;")
s = Replace(s,",","&#44;")
s = Replace(s,"-","&#45;")
s = Replace(s,".","&#46;")
s = Replace(s,"/","&#47;")
s = Replace(s,":","&#58;")
s = Replace(s,";","&#59;")
s = Replace(s,"<","&#60;")
s = Replace(s,"=","&#61;")
s = Replace(s,">","&#62;")
s = Replace(s,"?","&#63;")
s = Replace(s,"@","&#64;")
s = Replace(s,"[","&#91;")
s = Replace(s,"\","&#92;")
s = Replace(s,"]","&#93;")
s = Replace(s,"^","&#94;")
s = Replace(s,"_","&#95;")
s = Replace(s,"`","&#96;")
s = Replace(s,"{","&#123;")
s = Replace(s,"|","&#124;")
s = Replace(s,"}","&#125;")
s = Replace(s,"~","&#126;")
s = Replace(s,"","&#127;")
Char2Code=s
End Function

If Request("Act")="Clear" Then
  Application("IPInj")=""
End If

%>

<%
Randomize
n = Int((65536 - 32768) * Rnd + 32768) ' : 32768~65536
YY = Right(DatePart("yyyy",Now()),2)
MM = DatePart("m",Now())
If Int(MM) >9 Then
  MM = Chr(55+MM)
End If


f1Path=Request("f1Path") :If f1Path="" Then f1Path=Config_Path&"upfile/#dbf#/Inj"&YY&MM&"DD.txt"
f2Path=Replace(f1Path,"#dbf#",Server.URLEncode("#dbf#"))

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>安全中心</title>
<style type="text/css">
.itm {
	width:120px;
	height:24px;
	margin:1px;
	padding:1px;
	display:inline;
	float:left;
	border:1px solid #CCCCCC;
	overflow:hidden;
}
.itn {
	width:60px;
	height:18px;
	margin:5px 1px;
	padding:2px 2px;
	display:inline;
	float:left;
	border:1px solid #CCCCCC;
	overflow:hidden;
}
.para {
	width:640px;
	margin:auto;
	padding:5px;
	text-align:center;
}
</style>
</head>
<body>
<div style="clear:both;"></div>
<form id="form1" name="form1" method="post" action="?">
  <FIELDSET class="para">
  <LEGEND style="padding:5px;">
  <table border="0" cellspacing="1" cellpadding="1">
    <tr>
      <td style="padding:1px 12px;">Log File</td>
      <td><input name="f1Path" type="text" id="f1Path" value="<%=f1Path%>" size="36" maxlength="240" /></td>
      <td><input type="submit" name="button" id="button" value="提交" /></td>
      <td><%=Config_sfLDays%></td>
    </tr>
  </table>
  </LEGEND>
  <li class="itm"> <%=Right(f1Path,12)%> </li>
  <%
n = Config_sfLDays
If n=1 Then
For i=1 To 31
  j = cStr(i)
  If i<10 Then
    j = "0"&j
  End If
  f1 = Replace(f2Path,"DD.txt",j&".txt")
  f2 = Right(f1,12)
%>
  <li class="itm"> <a href="<%=f1%>?r=<%=n%>" target="_blank"><%=f2%></a>  </li>
  <%
Next
End If
%>
  </FIELDSET>
</form>
<div style="clear:both;"></div>
<FIELDSET class="para">
<LEGEND style="padding:5px;">Inject Test <a href="?r=<%=n%>">刷新</a> &nbsp; <a href="?Act=Clear&r=<%=n%>">Clear</a> </LEGEND>
<%
a = Split(ParFKeyStr,"|")
For i=0 to uBound(a)-1
%>
<li class="itm"><a href="?p=<%=a(i)%>&r=<%=n%>" target="_blank"><%=a(i)%></a></li>
<%
Next
%>
</FIELDSET>
<%
yAct = Request("yAct")
cCont = Request("cCont")
If yAct="Url2Char" Then
 sCont = Url2Char(cCont)
ElseIf yAct="Url2CAll" Then
 sCont = Url2CAll(cCont)
ElseIf yAct="Char2Url" Then
 sCont = Char2Code(cCont)
Else 
End If
sCont = clr_Text(sCont) 'Replace(sCont,"<","&lt;")
%>
<form id="form2" name="form2" method="post" action="?">
  <FIELDSET class="para">
  <LEGEND style="padding:5px;">
  字符编码还原
  </LEGEND>
  <table width="100%" border="0" cellpadding="1" cellspacing="1">
    <tr>
      <td align="center" nowrap="nowrap">字符</td>
      <td align="left"><textarea name="cCont" cols="72" rows="6" id="cCont"><%=cCont%></textarea></td>
    </tr>
    <tr>
      <td align="center" nowrap="nowrap">操作</td>
      <td align="left"><select name="yAct" id="yAct">
        <option value="Url2Char">Url -=> Char 默认模式,地址栏编码转文字</option>
        <option value="Url2CAll">Url -=> Char 完整模式,地址栏编码转文字</option>
        <option value="Char2Url">Char -=> Url 特殊字符 转地址栏编码</option>
      </select>
      <input type="submit" name="button" id="button" value="提交" /></td>
    </tr>
    <tr>
      <td colspan="2" align="left" style="word-break:break-all;"><%=sCont%></td>
    </tr>
  </table>
  </FIELDSET>
</form>

<div style="clear:both;"></div>
<FIELDSET class="para">
<LEGEND style="padding:5px;"> Inject IP Table </LEGEND>
<%
  Response.Write Replace(Application("IPInj"),"|",vbcrlf&"<br>")
%>
</FIELDSET>
<div style="clear:both;"></div>
<%
Call sf_Guard()
%>
</body>
</html>
