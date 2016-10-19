<!--#include file="../upfile/sys/config/MTypList.asp"-->

<%
'SubDirect()
Set rs=Server.CreateObject("Adodb.Recordset")
verNow = "2"
'Call sf_Guard()
'Call cchCheck("") 'cchRAM() 'Call cchFile()
%>
<!--#include file="../inc/home/func3.asp"-->
<!--#include file="../pfile/lang/vpage.asp"-->

<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!--#include file="../sadm/func1/func_file.asp"-->

<!--#include file="../sadm/func1/md5_func.asp"-->
<!--#include file="admin/mconfig.asp"-->
<!--#include file="../pfile/lang/vmemb.asp"-->
<!--#include file="../inc/form/form_app.asp"-->
<%

get30Time = Request(App30Code&"0")
get30TSN = Request(App30Code&"1") 
get30TStr = Request(App30Code&"2")
get30Tab = Split(get30TStr,"-")
  
app99Tim = Now() 'date("Y-m-d H:i:s");
app30Arr = App30Set(app99Tim,"","")
app30Tab = Split(app30Arr(2),"-")

urlEvent = Server.HTMLEncode("class='lgnBtn1 lgnInpt' tabindex=1 onClick='ChkAShow()'") 
url30Par = "&"&App30Code&"0="&Server.HTMLEncode(app30Arr(0))&"&"&App30Code&"1="&app30Arr(1)&"&"&App30Code&"2="&app30Arr(2)&"" '//&".$App30Code."2=$app30Arr[2]
'urlParas = "?act=setInput&n="&Get_IDEnc(app30Tab(3),0)&"&s=30&m=16&e="&urlEvent&url30Par&""
'tmpStr = "?act=showMessage&s="&Get_IDEnc("PEACE123abc45&#20320;6",0)&url30Par&""

%>
