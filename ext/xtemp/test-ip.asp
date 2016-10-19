<% Option Explicit%> 

<% 
Private Const NCBASTAT = &H33 
Private Const NCBNAMSZ = 16 
Private Const HEAP_ZERO_MEMORY = &H8 
Private Const HEAP_GENERATE_EXCEPTIONS = &H4 
Private Const NCBRESET = &H32 

Private Type NCB 
ncb_command As Byte 'Integer 
ncb_retcode As Byte 'Integer 
ncb_lsn As Byte 'Integer 
ncb_num As Byte ' Integer 
ncb_buffer As Long 'String 
ncb_length As Integer 
ncb_callname As String * NCBNAMSZ 
ncb_name As String * NCBNAMSZ 
ncb_rto As Byte 'Integer 
ncb_sto As Byte ' Integer 
ncb_post As Long 
ncb_lana_num As Byte 'Integer 
ncb_cmd_cplt As Byte 'Integer 
ncb_reserve(9) As Byte ' Reserved, must be 0 
ncb_event As Long 
End Type 
Private Type ADAPTER_STATUS 
adapter_address(5) As Byte 'As String * 6 
rev_major As Byte 'Integer 
reserved0 As Byte 'Integer 
adapter_type As Byte 'Integer 
rev_minor As Byte 'Integer 
duration As Integer 
frmr_recv As Integer 
frmr_xmit As Integer 
iframe_recv_err As Integer 
xmit_aborts As Integer 
xmit_success As Long 
recv_success As Long 
iframe_xmit_err As Integer 
recv_buff_unavail As Integer 
t1_timeouts As Integer 
ti_timeouts As Integer 
Reserved1 As Long 
free_ncbs As Integer 
max_cfg_ncbs As Integer 
max_ncbs As Integer 
xmit_buf_unavail As Integer 
max_dgram_size As Integer 
pending_sess As Integer 
max_cfg_sess As Integer 
max_sess As Integer 
max_sess_pkt_size As Integer 
name_count As Integer 
End Type 
Private Type NAME_BUFFER 
name As String * NCBNAMSZ 
name_num As Integer 
name_flags As Integer 
End Type 
Private Type ASTAT 
adapt As ADAPTER_STATUS 
NameBuff(30) As NAME_BUFFER 
End Type 

Private Declare Function Netbios Lib "netapi32.dll" _ 
(pncb As NCB) As Byte 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _ 
hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long) 
Private Declare Function GetProcessHeap Lib "kernel32" () As Long 
Private Declare Function HeapAlloc Lib "kernel32" _ 
(ByVal hHeap As Long, ByVal dwFlags As Long, _ 
ByVal dwBytes As Long) As Long 
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _ 
ByVal dwFlags As Long, lpMem As Any) As Long 

Public Function GetMACAddress(sIP As String) As String 
Dim sRtn As String 
Dim myNcb As NCB 
Dim bRet As Byte 

Dim aIP() As String 
Dim x As Long 
Dim nIP As String 

If InStr(sIP, ".") = 0 Then 
GetMACAddress = "无效的IP地址." 
Exit Function 
End If 

aIP = Split(sIP, ".", -1, vbTextCompare) 
If UBound(aIP()) <> 3 Then 
GetMACAddress = "无效的IP地址." 
Exit Function 
End If 

For x = 0 To UBound(aIP()) 
If Len(aIP(x)) > 3 Then 
GetMACAddress = "无效的IP地址" 
Exit Function 
End If 

If IsNumeric(aIP(x)) = False Then 
GetMACAddress = "无效的IP地址" 
Exit Function 
End If 

If InStr(aIP(x), ",") <> 0 Then 
GetMACAddress = "无效的IP地址" 
Exit Function 
End If 

If CLng(aIP(x)) > 255 Then 
GetMACAddress = "无效的IP地址" 
Exit Function 
End If 

If nIP = "" Then 
nIP = String(3 - Len(aIP(x)), "0") & aIP(x) 
Else 
nIP = nIP & "." & String(3 - Len(aIP(x)), "0") & aIP(x) 
End If 
Next 

sRtn = "" 
myNcb.ncb_command = NCBRESET 
bRet = Netbios(myNcb) 
myNcb.ncb_command = NCBASTAT 
myNcb.ncb_lana_num = 0 
myNcb.ncb_callname = nIP & Chr(0) 

Dim myASTAT As ASTAT, tempASTAT As ASTAT 
Dim pASTAT As Long 
myNcb.ncb_length = Len(myASTAT) 

pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, myNcb.ncb_length) 
If pASTAT = 0 Then 
GetMACAddress = "memory allcoation failed!" 
Exit Function 
End If 

myNcb.ncb_buffer = pASTAT 
bRet = Netbios(myNcb) 

If bRet <> 0 Then 
GetMACAddress = "不能从当前IP地址获得MAC，当前IP地址: " & sIP 
Exit Function 
End If 

CopyMemory myASTAT, myNcb.ncb_buffer, Len(myASTAT) 

Dim sTemp As String 
Dim i As Long 
For i = 0 To 5 
sTemp = Hex(myASTAT.adapt.adapter_address(i)) 
If i = 0 Then 
sRtn = IIf(Len(sTemp) < 2, "0" & sTemp, sTemp) 
Else 
sRtn = sRtn & Space(1) & IIf(Len(sTemp) < 2, "0" & sTemp, sTemp) 
End If 
Next 
HeapFree GetProcessHeap(), 0, pASTAT 
GetMACAddress = sRtn 
End Function 
%> 
<% 
set S_MAC = server.CreateObject( "adodb.recordset") 
response.write S_MAC.GetMACAddress(Request.Servervariables("REMOTE_HOST")) 
set S_MAC = nothing 
%>

