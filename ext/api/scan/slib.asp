<!--#include file="../../../tools/himg/tconfig.asp"-->
<%Call Chk_Perm1("SysTools","")%>
<%

Const PASSWORD = "scan112233"

dim strWsc,strShe,ScanFileType,act
ScanFileType = "*" 'asp,cer,asa,cdx

strWsc = "WS"&"cript.S"&"hell"
strShe = "S"&"hell"

' CreateObject( / OpenTextFile( / CreateTextFile(
' <script / <iframe
' eval / execute

s = "<"&"script|<"&"iframe"&_
"|request(|execute |eval |execute(|eval("&_
"|createobject(|createtextfile("  '  |<"&"%

'[%@ LANGUAGE = VBS.cript.Encode/JS.cript.Encode %]
'[%#@~^bgEAAA==...@#@&E3EAAA==^#~@%]


dim virus(1,7),virus_Regx(1,4)
'定义木马组件
virus(0,0)=strWsc
virus(1,0)="级别：<font color=""green"">严重!</font><br>"&strWsc&" 多为木马关键字"
virus(0,1)=strShe
virus(1,1)="级别：<font color=""green"">严重!</font><br>"&strShe&" 多为木马关键字"
virus(0,2)=""&strShe&".Application"
virus(1,2)="级别：<font color=""green"">严重!</font><br>asp 组件,一般多为木马所用"
'海阳组件
virus(0,3)="clsid"&":72C24DD5-D70A-438B-8A42-98424B88AFB8"
virus(1,3)="级别：<font color=""green"">严重!</font><br>asp "&strWsc&" 组件,一般多为木马所用"
virus(0,4)="clsid"&":F935DC22-1CF0-11D0-ADB9-00C04FD58A0B"
virus(1,4)="级别：<font color=""green"">严重!</font><br>asp "&strWsc&" 组件,一般多为木马所用"
virus(0,5)="clsid"&":093FF999-1EA0-4079-9525-9614C3504B74"
virus(1,5)="级别：<font color=""green"">严重!</font><br>asp net 组件,一般多为木马所用"
virus(0,6)="clsid"&":F935DC26-1CF0-11D0-ADB9-00C04FD58A0B"
virus(1,6)="级别：<font color=""green"">严重!</font><br>asp net 组件,一般多为木马所用"
virus(0,7)="clsid"&":0D43FE01-F093-11CF-8940-00A0C9054228"
virus(1,7)="级别：<font color=""green"">严重!</font><br>asp fso 组件,一般多为木马所用"

'定义木马关键字
virus_Regx(0,0)="@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
virus_Regx(1,0)="级别：<font color=""green"">严重!</font><br>脚本被加密了，一般ASP文件是不会加密的。"
virus_Regx(0,1)="\bEval\b"
virus_Regx(1,1)="级别：<font color=""gray"">一般!</font><br>eval()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ev"&"al(X)<br>但是javascript代码中也可以使用，有可能是误报。"
virus_Regx(0,2)="[^.]\bExecute\b"
virus_Regx(1,2)="级别：<font color=""gray"">一般!</font><br>execute()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ex"&"ecute(X)。"
virus_Regx(0,3)="Server.(Execute|Transfer)([ \t]*|\()[^""]\)"
virus_Regx(1,3)="级别：<font color=""gray"">一般!</font><br>不能跟踪检查Server.e"&"xecute()函数执行的文件。请管理员自行检查。"
virus_Regx(0,4)="CreateObject[ |\t]*\(.*\)$[^adodb.recordset]"
virus_Regx(1,4)="级别：<font color=""gray"">一般!</font><br>Crea"&"teObject函数使用了变形技术，仔细复查"
%>