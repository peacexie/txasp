<!--#include file="../../tools/base/safe.asp"-->
<%

'/////////////////////////////////////////////////////////////////////////////////
function gMemPerm(xconn,xGrade)
Dim gArr,sPerm,rsp,i,tGrade,sqlp,tPerm
xGrade = Replace(xGrade," ","")
xGrade = Replace(xGrade,";",",")
gArr = Split(xGrade,",")
sPerm = ""
Set rsp=Server.Createobject("Adodb.Recordset")
For i = 0 To UBound(gArr)
    tGrade = gArr(i)
	If Trim(tGrade&"")<>"" Then
	  sqlp = " SELECT * FROM [MemSyst] WHERE SysID='"&tGrade&"'"  
      rsp.Open Sqlp,xconn,1,1
	    If NOT rsp.EOF Then
	      tPerm = rsp("SysPerm")&""
		  If tPerm<>"" Then
		    sPerm = sPerm &"{"& tPerm &"}"
		  End If
	    End If
      rsp.Close()   
	End If
Next
set rsp = NoThing  ': Response.Write ""&sPerm : Response.End()
gMemPerm = sPerm 
end function 	


'' ------------------------ ------------------------ 

varDef = "1"
verMemb = Request("verMemb")
If verMemb="" Then
  verMemb = varDef
End If

If verMemb="2" Then
  mCfgCode = "Corp;Privy;Gov;Org"
  mCfgName = "Enterprise;Personal;Government;Organization"
Else
  mCfgCode = "Corp;Privy;Gov;Org"
  mCfgName = "公司;个人;政府;组织"
End If

'' ------------------------ ------------------------ 

sCTypTraN124 = "(TypN1241);(TypN1242)"
sNTypTraN124 = "超管新闻1;超管新闻2"

sCTypTraT124 = "(TypT1241);(TypT1242)"
sNTypTraT124 = "超管供求1;超管供求2"


%>