<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/funcfile.asp"-->

<%

Set rs=Server.Createobject("Adodb.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject") 
s0 = ""

fPath = "/upfile/ind10/"
FlgMod = Request("FlgMod")&""
FlgTim = Timer() 
FlgGap = 15   ' 5分钟更新一次,建议此值大于5,小于60(1小时)
  If Application("FlgTim")&""="" Then
    Application("FlgTim") = (-1)*36000000
  End If

If FlgMod = "Clear" Then 
	Application("FlgTim") = (-1)*36000000
	fil = Log_fName()
	Response.Write "<center><br>请点击以下连接刷新相应部分信息<br><br>"
	Response.Write "<a target='_blank' href='?FlgMod=xxxxx'>Page</a> | "
	Response.Write "<a target='_blank' href='?FlgMod=Home'>首页</a> | "
	Response.Write "<a target='_blank' href='"&fPath&""&fil&"'>Logs</a> "
	Response.Write "</center>" 
ElseIf Abs(FlgTim-Application("FlgTim"))>FlgGap*60 AND FlgMod="Home" Then
	Application("FlgTim") = FlgTim
	Call uDG_Res()
	Call uDG_Pos()
	Call uInd_Res()
	Call uInd_Pos()
	Call uXml_Pos()
	Call Add_Logs(FlgMod)
End If

Set rs = Nothing
Set fso = Nothing

PagChk = LCase("/inc/home/upd_out.asp")
PagRef = LCase(Request.Servervariables("HTTP_REFERER"))

If inStr(PagRef,PagChk)>0 And FlgMod<>"Clear" Then
	Response.Write "刷新完成!"&Application("FlgMod")
End If

Function Add_Logs(xAct)
  Dim fil,dat,fp
  fil = Log_fName()
  fp = ""&fPath&""&fil
  dat = Now()&" : "&xAct&" : "&CStr(Timer()-FlgTim)
  If fil_exist(""&fPath&"",fil) Then
	Set fil2 = fso.OpenTextFile(Server.MapPath(fp),8,True)
	fil2.WriteLine(dat) 
  Else
    Set fil2 = fso.CreateTextFile(Server.MapPath(fp),True)
	fil2.Write(dat&vbcrlf)
  End If
  fil2.Close
End Function

Function Log_fName()
Dim fName,fNO,fGap
 fGap = 11 ' 4(8),7(5),11(3)
 fName = Left(Get_yyyymmdd(""),8)
 fNO=Right(fName,2) : fName=Left(fName,6)
 fNO=Int(fNO/fGap)  : fName=fName&fNO&".txt"
Log_fName = fName 
End Function

'<-script language='javascript' src='/inc/home/upd_out.asp?FlgMod=Home'></script->
%>
