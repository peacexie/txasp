<%
 'On Error Resume Next
 '�Ե�ַ�������ļ��
 SQL_injdata = "'| and | or |exec |insert into |select |delete |update |count | * |&%|chr(|mid(|master|truncate|char(|declare "
 SQL_inj = split(SQL_Injdata,"|") 
 Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx
 Fy_Cl = 1        '����ʽ��1=��ʾ��Ϣ,2=ת��ҳ��,3=����ʾ��ת��
 Fy_Zx = "Error.Asp"    '����ʱת���ҳ��
 
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
     Response.Write "<Script Language=javascript>alert('                   ���ִ��󣡲��� "&Fy_Cs(Fy_x)&" ��ֵ�а����Ƿ��ַ�����\n\n  �벻Ҫ�ڲ����г��֣�;,and,select,update,insert,delete,chr, �ȷǷ��ַ���\n');history.go(-1)</Script>"
  Case "2"
     Response.Write "<Script Language=javascript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
     Response.Write "<Script Language=javascript>alert('                   ���ִ��󣡲��� "&Fy_Cs(Fy_x)&"��ֵ�а����Ƿ��ַ�����\n\n  �벻Ҫ�ڲ����г��֣�;,and,select,update,insert,delete,chr, �ȷǷ��ַ���\n');location.href='"&Fy_Zx&"';</Script>"
 End Select
   Response.End
End If
next
 End If
Next
'�Ա������ļ�⣬��post��ʽ ��get��ʽ������⣬get��ʽ�������ᱣ���¼
if request.form<>"" then
 
FOR EACH name IN Request.Form
for i=0 to ubound(SQL_inj)
If Instr(LCase(request.form(name)),SQL_inj(i))<>0 Then
Select Case  Fy_Cl
  Case "1"
Response.Write "<Script Language=JavaScript>alert(' �� "&name&" ��ֵ�а����Ƿ��ַ�����\n\n�벻Ҫ�ڱ��г��֣�;,and,select,update,insert,delete,chr,  �ȷǷ��ַ���');history.go(-1)</Script>"
  Case "2"
Response.Write "<Script Language=JavaScript>location.href='"&Err_Web&"'</Script>"
  Case "3"
Response.Write "<Script Language=JavaScript>alert(' ���� "&name&"��ֵ�а����Ƿ��ַ�����\n\n�벻Ҫ�ڱ��г��֣�;,and,select,update,insert,delete,chr,  �ȷǷ��ַ���');location.href='"&Err_Web&"';</Script>"
End Select
Response.End
End If
NEXT
NEXT
end if
%>