<%

'/* Peace ˵�� ***************************************

'Ĭ��, ÿ300�� ���»���һ�Σ�
'������Ҫ���ò�����/dghome.php:$cache_exp

'������� ���Բ���
'?is_upd=Clear ɾ������,ǿ�Ƹ���;  
'?is_upd=noUpd ������,ֱ�Ӷ������ļ�; ��
'?is_upd=(������Ϊ��) ��������,ǿ�Ƹ��»���; ��
 
'����ʱ�� ���Բ���
'?is_tst=(������Ϊ��)����ʾ�������ļ���ʱ��; һ���С
'�����ã�����ʾ; 
'������ʹ�û����£���ʾ�������ļ���ʱ�䣩

'* **************************************************/

tTimer = timer() '; //���ڲ���

cache_exp = 300 '; //����ʱ��,��λ:�� (86400=24Сʱ) 
cache_file = "./data/cache/dghome.htm" '; //�����ļ�,[./]��ͷ
real_file = "/peace/dgind.php" '; //Ҫִ�е���ʵ�ļ�,[/]��ͷ
cache_flag = "Read" ';

Response.Write urlRead("http://b.dg.gd.cn/dghome.php") 

Function fileRead(pathName,xCSet)
 Dim str
 if fileExist(pathName) then
  set stm=server.CreateObject("Ado"&"db.Str"&"eam")
  stm.Type=2 
  stm.Mode=3 
  stm.Charset=xCSet
  stm.Open
  stm.loadfromfile Server.MapPath(pathName)
  str=stm.ReadText
  stm.Close
  set stm=nothing
 else
  str = "" '(File Read Error!)
 end if
 fileRead = str
End Function

Function urlRead(strUrl) 
  Dim objHttp 
  'strUrl = urlPath(strUrl)
  On Error Resume Next 
  Set objHttp = Server.CreateObject("Micro"&"soft.XML"&"HTTP") 
  With objHttp 
    .Open "Get", strUrl, False, "", "" 
    .Send 
  End With 
  if objHttp.Readystate <> 4 then 
    Set objHttp = Nothing 
    urlRead = False 
    Exit Function 
  end if 
  'objHttp.Charset = "utf-8"
  urlRead = objHttp.responseText  
  Set objHttp = Nothing 
End Function

%>

