<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp"-->
<!--#include file="../../upfile/sys/para/tmstyp2.asp" -->
<!--#include file="../../upfile/sys/para/tmsPara.asp" -->

http://localhost:240/u/demo/pfile/temc/xx_func.asp<br />
http://localhost:240/u/data/upfile/dtinf/2010/78/B6WV.5EWMM/index.htm<br />

<%

'生成静态，

'函数 
  'func_sfile.asp(Sub add_sfFile),            OK 1/2
  'func_sfile.asp(Sub get_TmpLink),           OK
  'upd_data.asp(function fConv=CFile),        ??
  
'设置：Config_Cont = "Html"

'/pfile/temc/*.htm
  'tModID --- tInfN124.htm
  'tTypID --- tN110012.htm
 

ID3 = "dtpic-2010-6W-D3AV.01DXN"
upRoot = Config_Path&"upfile/"
upPath = upRoot&Replace(ID3,"-","/")&"/" 
Call add_TmpHtml("InfoPics",ID3,"xCont") ',""


'Response.Write vbcrlf&"<hr> temp 1News = "&Show_Text(get_TmpCont("1News"))
'Response.Write vbcrlf&"<hr> temp 1Pics = "&get_TmpCont("1Pics")

'Response.Write vbcrlf&"<hr> temp 6UD = "&get_TmpCont("6UD")
'Response.Write vbcrlf&"<hr> temp 6Left = "&get_TmpCont("6Left")

%>



