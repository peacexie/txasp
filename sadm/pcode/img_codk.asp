<%
Option Explicit

Const CodeType = 2 'ע��1,4,7,10,13,16Ϊ�ڰ��� 2,5,8,11,14,17Ϊ��ɫ������ 3,6,9,12,15,18Ϊ�����
Const listcode = "0123456789abcdefghijklmnopqrstuvwxyz"

Response.buffer = True
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-ctrol", "no-cache"

Dim zNum, rNum, i, j, listnum
Dim Ados, Ados1, Config_PSess

'�õ���֤����ַ���
Dim zimg(6), NStr
Randomize Timer
For i = 0 To 5
    rNum = Fix(35 * Rnd) '��35��Ϊ9��Ϊʹ�ô���������
    zimg(i) = rNum
    listnum = listnum & Mid(listcode, rNum + 1, 1)
Next
'Session("ChkCode") = listnum
Config_PSess = Request("Config_PSess")
If Config_PSess="" Then '��¼��Session
  Session("ChkCode") = listnum
Else
  Session(Config_PSess) = listnum
End If
'*********************

Dim Pos
Set Ados = Server.CreateObject("Adodb.Stream")
Ados.Mode = 3
Ados.Type = 1
Ados.Open
Set Ados1 = Server.CreateObject("Adodb.Stream")
Ados1.Mode = 3
Ados1.Type = 1
Ados1.Open

'�õ���֤��ͼ��ʵ�岿��
Ados.LoadFromFile Server.Mappath("img_codk/body" & CodeType & ".Fix")
Ados1.write Ados.Read(2880)
For i = 0 To 5
    Ados.Position = (35 - zimg(i)) * 480
    Ados1.Position = i * 480
    Ados1.write Ados.Read(480)
Next

'�õ�ͼ��ͷ����Ϣ
Ados.LoadFromFile Server.mappath("img_codk/head.fix")
Pos = LenB(Ados.Read())
Ados.Position = Pos

'��ͷ����Ϣ��ʵ�岿�ֺϲ��ɺ�������
On Error Resume Next
For i = 0 To 15
    For j = 0 To 5
        Ados1.Position = i * 32 + j * 480
        Ados.Position = Pos + 30 * j + i * 270
        Ados.write Ados1.Read(30)
    Next
Next

'���ͼ��
Ados.Position = 0
Response.ContentType = "image/BMP"
Response.BinaryWrite Ados.Read()

Ados.Close
Set Ados = Nothing
Ados1.Close
Set Ados1 = Nothing
%>