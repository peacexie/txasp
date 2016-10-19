<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>Encode_Big5</title>
</head>
<body>
<%

ch = "試"
cu = AscW(ch) :Response.Write "<br>1."&cu
cr = ChrW(cu) :Response.Write "<br>2."&cr
c16 = Hex(cu) :Response.Write "<br>3."&c16
'c17 = c2to10(c16to2(c16)) :Response.Write "<br>4."&c17

Function vbsEscape(str) 
    dim i,s,c,a 
    s="" 
    For i=1 to Len(str) 
        c=Mid(str,i,1) 
        a=ASCW(c) 
        If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then 
            s = s & c 
        ElseIf InStr("@*_+-./",c)>0 Then 
            s = s & c 
        ElseIf a>0 and a<16 Then 
            s = s & "%0" & Hex(a) 
        ElseIf a>=16 and a<256 Then 
            s = s & "%" & Hex(a) 
        Else 
            s = s & "%u" & Hex(a) 
        End If 
    Next 
    vbsEscape = s 
End Function

Function vbsUnEscape(str) 
    dim i,s,c 
    s="" 
    For i=1 to Len(str) 
        c=Mid(str,i,1) 
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then 
            If IsNumeric("&H" & Mid(str,i+2,4)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4))) 
                i = i+5 
            Else 
                s = s & c 
            End If 
        ElseIf c="%" and i<=Len(str)-2 Then 
            If IsNumeric("&H" & Mid(str,i+1,2)) Then 
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2))) 
                i = i+2 
            Else 
                s = s & c 
            End If 
        Else 
            s = s & c 
        End If 
    Next 
    vbsUnEscape = s 
End Function 

Response.Write "<pre>"
org = "Test--測試--?#&%abc@_"
enc = vbsEscape(org)
un = vbsUnEscape(enc) 
url = Server.URLEncode(org)
q = Request("q")
Response.Write vbcrlf&"<br>org:"&org&"-----<a href='?'>"&org&"</a>"
Response.Write vbcrlf&"<br>enc:"&xxx&"-----<a href='?q="&enc&"'>"&enc&"</a>"
Response.Write vbcrlf&"<br>un-:"&un
Response.Write vbcrlf&"<br>url:"&xxx&"-----<a href='?q="&url&"'>"&url&"</a>"
Response.Write vbcrlf&"<br>q--:"&q

Response.Write vbcrlf&"<br>"
Response.Write vbcrlf&"<br>[XX]-測試-望眼:"&vbsEscape("[XX]-測試-望眼")
Response.Write vbcrlf&"<br><a href='encode.asp?q="&enc&"'>en_enc</a> | <a href='enc_gb.asp?q="&url&"'>gb_url</a>"
Response.Write vbcrlf&"<br><a href='encode.asp?q="&enc&"'>gb_enc</a> | <a href='enc_gb.asp?q="&url&"'>en_url</a>"
Response.Write vbcrlf&"<br><a href='?q="&q&"'>q_org</a> | <a href='?q="&vbsEscape(q)&"'>q_enc</a> "

Response.Write vbcrlf&"</pre>"

' js/Ajax escape()  unescape()。
' //32(255~=6x40)  ' //64(512~=6x80) 1024


%>
</body>
</html>
