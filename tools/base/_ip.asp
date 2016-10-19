
<% 

'ip = IP_User()
Function IP_User() 
Dim cIP
  cIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
  If cIP = "" or isnull(cIP) or isempty(cIP) Then 
    cIP = Request.ServerVariables("REMOTE_ADDR") 
  End If 
IP_User = cIP 
End function 

'n  = IP_Conv("219.16.73.120") -> 3675277688
Function IP_Conv(xIP) 
Dim a,n
  a = Split(xIP,".")
  n = (256^3)*a(0)+(256^2)*a(1)+256*a(2)+a(3)
IP_Conv = n
End Function 

'n  = IP_Addr(1032881152)
Function IP_Addr(xN) 
Dim sql
  sql = "SELECT addr FROM IPAddress WHERE ip1<="&xN&" and ip2>="&xN&""
IP_Addr = rs_Val("",sql)
End Function 

's  = IP_CBack(3675277688) -> 219.16.73.120
Function IP_CBack(xN) 
Dim a,m0,M : m0 = xN 
  M = 256^3  'm1 = m0 mod (256^3)
  m1 = (m0/M-Int(m0/M))*M : a1 = Int(m0/(256^3))
  m2 = m1 mod (256^2)     : a2 = Int(m1/(256^2))
  m3 = m2 mod 256         : a3 = Int(m2/256)
IP_CBack = a1&"."&a2&"."&a3&"."&m3
End Function 

%>


