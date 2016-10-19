<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->

<%

upRoot = Config_Path&"upfile/"

Call Chk_Perm9("MemFUpload","3") 
'Call Chk_URL()

Function ImgfPath()
Dim sPath,mStr,tDate,tSTab
  If Request("upPath")<>"" Then
    sPath = Request("upPath")
	sPath = Replace(sPath,"-","/")
  Else
    tDate = Now()
	tSTab = "123456789ABCDEFGHJKMNPQRSTUVWXY"
	mStr = Mid(tSTab,DatePart("m",tDate),1)
	mStr = mStr&Mid(tSTab,DatePart("d",tDate),1)
    sPath = "defup/"&Year(Now)&"/"&mStr
  End If
  Call fold_add9(Config_Path&"upfile/",sPath,0)
ImgfPath = sPath
End Function

%>
