
<%

'http://caijujuan.blog.changanedu.com/
'http://www.changanedu.com/blog/blog.asp?name=zengwenyan
'http://www.changanedu.com/blog/u/caijujuan/index.htm
'Http_Host=Server_Name:Server_Port��

url = Request.ServerVariables("SERVER_NAME")

If inStr(url,"blog.changanedu.com")>0 Then
  If inStr(url,".blog.changanedu.com")>0 Then
    id = Replace(url,".blog.changanedu.com","")
	Response.Redirect "http://www.changanedu.com/blog/blog.asp?name="&id&""
  Else
    Response.Redirect "http://www.changanedu.com/blog/"
  End If
Else
  If inStr(url,"changanchs.com")>0 Then
	Response.Redirect "http://www.changanchs.com/"
  Else
    Response.Redirect "http://www.changanedu.com/"
  End If
End If

%>

'////////////////////////////////////////////////////// 

<%

id = LCase(Request.ServerVariables("SERVER_NAME")) 
idBak = id
id = Replace(id,".96327.com","")
id = Replace(id,".dg.gd.cn","")

Dim addr,idArr(8,4)

idArr(1,0) = "sys"                  ' ����:��˾,trade;����,supply;��Ʒ,product;
idArr(1,1) = "http://www.dg.gd.cn/sys/sadm/login.asp" 
idArr(1,2) = "http://www.dg.gd.cn/sys/sadm/login.asp"     
idArr(1,3) = "��ҵ��ҳ - ����,��Ʒ,��ҵ��ҳ"    

idArr(2,0) = "dgcity"                   ' ��̳,bbs;����,love;
idArr(2,1) = "http://www.dg.gd.cn/dgcity/" 
idArr(2,2) = "http://www.dg.gd.cn/dgcity/"    
idArr(2,3) = "����BBS��ҳ - ����ר������"   
idArr(3,0) = "member"                    
idArr(3,1) = "http://www.dg.gd.cn/sys/member/" 
idArr(3,2) = "http://www.dg.gd.cn/sys/member/"   
idArr(3,3) = "���˽�����ҳ - ���齻��"     

idArr(4,0) = "blog"                  ' ����,blog;���,photo;   
idArr(4,1) = "http://blog.dg.gd.cn/blog.asp?id=" 
idArr(4,2) = "http://blog.dg.gd.cn/"  
idArr(4,3) = "���˲�����ҳ - �������"
idArr(5,0) = "photo"                    
idArr(5,1) = "http://blog.dg.gd.cn/photo.asp?id=" 
idArr(5,2) = "http://blog.dg.gd.cn/"  
idArr(5,3) = "���������ҳ - �������"

idArr(6,0) = "trade"                  ' ����:��˾,trade;����,supply;��Ʒ,product;
idArr(6,1) = "http://trade.dg.gd.cn/co_index.asp?COID=" 
idArr(6,2) = "http://trade.dg.gd.cn/"     
idArr(6,3) = "��ҵ��ҳ - ������Ϣ,��Ʒ����,��Ʒ�ƹ�,��ҵ��ҳ-��ݸ��"  
idArr(7,0) = "job"                     ' �˲���:ְλ,job;����,resume;
idArr(7,1) = "http://www.jobease.com.cn/com_info/user.asp?id=" 
idArr(7,2) = "http://www.jobease.com.cn/" 
idArr(7,3) = "��ҵְλ��ҳ - �˲���"

addr = "(???)" 'addr = "http://www.dg.gd.cn/"
addf = "Frame"
For i = 1 To 7
  If id = idArr(i,0) Then
	  addr = idArr(i,2)
	  addf = "Redirect"
	  exit for
  ElseIf InStr(id,"."&idArr(i,0))>0 Then 
	  addr = idArr(i,1)&Replace(id,"."&idArr(i,0),"")	
	  TitName = idArr(i,3)
	  exit for
  End If
Next

'moodboker.trade.dg.gd.cn/
If inStr(idBak,".trade.dg.gd.cn")>0 Then
  t1 = inStr(idBak,".trade.dg.gd.cn")
  t2 = Left(idBak,t1-1)
  'Response.Write t2
  'Response.End
  Response.Redirect "http://trade.dg.gd.cn/co_index.asp?COID="&t2
  Response.End
ElseIf inStr(idBak,"218.16.118.235")>0 Then
  'Response.Redirect "http://www.dg.gd.cn/?"&Now()
   'Response.Write("��������������Ϊ������û�б�����ֹͣ�������ڱ����У����Ժ���ʣ��������ʣ�����ϵ����,�绰:0769-22028805 ��С��")
    response.write("")
     Response.End
End If

'Response.Write inStr(id,".trade.dg.gd.cn")
'Response.End

if addr = "(???)" then
  'Response.CodePage=936
  'Response.Charset="gb2312"
  'Response.Write "<p align='center'>��������������Ϊ������û�б�����ֹͣ�������ڱ����У����Ժ���ʣ��������ʣ�����ϵ����,�绰:0769-22028805 ��С�� </p>"
  response.write("  ")
 ' Response.Write "<p align='center'>&nbsp;</p>"
  'Response.Write "<p align='center' class='style1'><a href='http://www.dg.gd.cn' target='_blank'>��ݸ��</a>"
  Response.End()
elseif addf = "Redirect" then
  Response.Redirect addr
else
  Response.Redirect "http://www.dg.gd.cn/"
end if
%>

