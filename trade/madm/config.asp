<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<%


If Request("ModID")<>"" Then
  ModID = RequestS("ModID","C",48)
  Session("ModID") = ModID
ElseIf Session("ModID")&""<>"" Then
  ModID = Session("ModID")
Else
  ModID = "TraA124"
End If

If ModID="UserCorp" Then
 PathImg = "../../upfile/"
 TypeFile = ""
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "TradeCorp"
ElseIf inStr(ModID,"TraT")>0 Then
 PathImg = "../../upfile/"
 TypeFile = ""
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "TradePics"
Else
 PathImg = "../../upfile/"
 TypeFile = ".DOC/.XLS/.TXT/.HTML/.HTM"
 TypeImg = ".JPG/.GIF/.JPEG"
 ModTab = "TradeNews"
End If

'Call Chk_URL()
Call Chk_Perm1("ModTrade","") 

%>

