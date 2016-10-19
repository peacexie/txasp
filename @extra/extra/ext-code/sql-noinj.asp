<%
 'On Error Resume Next
 '对地址栏参数的检测
 SQL_injdata = "'| and | or |exec |insert into |select |delete |update |count | * |&%|chr(|mid(|master|truncate|char(|declare "
 SQL_inj = split(SQL_Injdata,"|") 
 Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx
 Fy_Cl = 1        '处理方式：1=提示信息,2=转向页面,3=先提示再转向
 Fy_Zx = "Error.Asp"    '出错时转向的页面
 
On Error Resume Next
Fy_Url=Request.ServerVariables("QUERY_STRING")
Fy_a=split(Fy_Url,"&")
redim Fy_Cs(ubound(Fy_a))

for Fy_x=0 to ubound(Fy_a)
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1)
Next
For Fy_x=0 to ubound(Fy_Cs)
If Fy_Cs(Fy_x)<>"" Then

 For SQL_Data=0 To Ubound(SQL_inj) 
  If Instr(LCase(Request(Fy_Cs(Fy_x))),Sql_Inj(Sql_DATA))<>0   Then
  Select Case Fy_Cl
  Case "1"
     Response.Write "<Script Language=javascript>alert('                   出现错误！参数 "&Fy_Cs(Fy_x)&" 的值中包含非法字符串！\n\n  请不要在参数中出现：;,and,select,update,insert,delete,chr, 等非法字符！\n');history.go(-1)</Script>"
  Case "2"
     Response.Write "<Script Language=javascript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
     Response.Write "<Script Language=javascript>alert('                   出现错误！参数 "&Fy_Cs(Fy_x)&"的值中包含非法字符串！\n\n  请不要在参数中出现：;,and,select,update,insert,delete,chr, 等非法字符！\n');location.href='"&Fy_Zx&"';</Script>"
 End Select
   Response.End
End If
next
 End If
Next
'对表单参数的检测，限post方式 对get方式不作检测，get方式服务器会保存记录
if request.form<>"" then
 
FOR EACH name IN Request.Form
for i=0 to ubound(SQL_inj)
If Instr(LCase(request.form(name)),SQL_inj(i))<>0 Then
Select Case  Fy_Cl
  Case "1"
Response.Write "<Script Language=JavaScript>alert(' 表单 "&name&" 的值中包含非法字符串！\n\n请不要在表单中出现：;,and,select,update,insert,delete,chr,  等非法字符！');history.go(-1)</Script>"
  Case "2"
Response.Write "<Script Language=JavaScript>location.href='"&Err_Web&"'</Script>"
  Case "3"
Response.Write "<Script Language=JavaScript>alert(' 参数 "&name&"的值中包含非法字符串！\n\n请不要在表单中出现：;,and,select,update,insert,delete,chr,  等非法字符！');location.href='"&Err_Web&"';</Script>"
End Select
Response.End
End If
NEXT
NEXT
end if
%>