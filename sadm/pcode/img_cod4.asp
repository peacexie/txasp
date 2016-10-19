<%
'//===================================================
' 动网论坛验证码
' 作者	Dv.HxyMan
' 更新	2008-1-4
' 说明	前半部分为字库，后半部分为验证码生成程序。
'		可适当自行增加字库。默认为300个一级常用字。
'//===================================================
'----自动成生配置区_开始（请不要改动下面区块的内容，否则后台验证码设定程序将不能识别。）----
'<root><captcha_chartype_chinese>1</captcha_chartype_chinese><captcha_chartype_english>0</captcha_chartype_english><captcha_chartype_number>0</captcha_chartype_number><captcha_size>2</captcha_size><captcha_width_lbound>25</captcha_width_lbound><captcha_width_ubound>30</captcha_width_ubound><captcha_height_lbound>23</captcha_height_lbound><captcha_height_ubound>28</captcha_height_ubound><captcha_spacing_lbound>-6</captcha_spacing_lbound><captcha_spacing_ubound>0</captcha_spacing_ubound><captcha_angle_lbound>-15</captcha_angle_lbound><captcha_angle_ubound>15</captcha_angle_ubound><captcha_weight>1</captcha_weight><captcha_charshow>0</captcha_charshow><captcha_charshow_stepbystep_r>231</captcha_charshow_stepbystep_r><captcha_charshow_stepbystep_g>-1</captcha_charshow_stepbystep_g><captcha_charshow_stepbystep_b>-1</captcha_charshow_stepbystep_b><captcha_charshow_simple_r>0</captcha_charshow_simple_r><captcha_charshow_simple_g>0</captcha_charshow_simple_g><captcha_charshow_simple_b>8</captcha_charshow_simple_b><captcha_backshow>2</captcha_backshow><captcha_backshow_stepbystep_r>-1</captcha_backshow_stepbystep_r><captcha_backshow_stepbystep_g>-1</captcha_backshow_stepbystep_g><captcha_backshow_stepbystep_b>-1</captcha_backshow_stepbystep_b><captcha_backshow_simple_r>255</captcha_backshow_simple_r><captcha_backshow_simple_g>255</captcha_backshow_simple_g><captcha_backshow_simple_b>255</captcha_backshow_simple_b><captcha_charshow_mix_percent>30</captcha_charshow_mix_percent><captcha_backshow_mix_percent>10</captcha_backshow_mix_percent><captcha_pic_width>60</captcha_pic_width><captcha_pic_height>30</captcha_pic_height></root>
'----自动成生配置区_结束--------------------------------------------------------
'Option Explicit
Dim server_v1,server_v2,Chkpost
Chkpost=False
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
'If Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True
'If Not Chkpost Then Response.End
Dim f(350,4),u:u = 0
'If 1=0 Then
'f(u,0)="0":f(u,1)="000000000000000011100010001001000100100010010001001000100100010001110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="1":f(u,1)="000000000000000011000010100000010000001000000100000010000001000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="2":f(u,1)="000000000000000111100000001000000100000100000100000100000100000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="3":f(u,1)="000000000000000111100000001000000100011100000001000000100000010011110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="4":f(u,1)="000000000000000000100000110000101000010100010010001111100000100000010000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="5":f(u,1)="000000000000000111100010000001000000111000000010000001000000100011100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="6":f(u,1)="000000000000000011110010000001000000101100011001001000100100010001110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="7":f(u,1)="000000000000000111110000001000001000001000001000000100000100000010000000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'f(u,0)="8":f(u,1)="000000000000000011100010001001000100011100010011001000100100010001110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="9":f(u,1)="000000000000000011100010001001000100100110001101000000100000010011110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'End If
'If 1=0 Then
f(u,0)="A":f(u,1)="000000000000000001000001010000101000010100010001001111100100010100000100000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'f(u,0)="B":f(u,1)="000000000000000111100010001001000100111100010010001000100100010011110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="C":f(u,1)="000000000000000001110001000001000000100000010000001000000010000000111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="D":f(u,1)="000000000000000111000010010001000100100010010001001000100100100011100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="E":f(u,1)="000000000000000111110010000001000000111100010000001000000100000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="F":f(u,1)="000000000000000111110010000001000000111100010000001000000100000010000000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="G":f(u,1)="000000000000000001110001000001000000100000010001001000100010010000111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="H":f(u,1)="000000000000000100010010001001000100111110010001001000100100010010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'f(u,0)="I":f(u,1)="000000000000000111110000100000010000001000000100000010000001000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="J":f(u,1)="000000000000000011100000010000001000000100000010000001000000100011100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="K":f(u,1)="000000000000000100010010010001010000110000011000001010000100100010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="L":f(u,1)="000000000000000100000010000001000000100000010000001000000100000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="M":f(u,1)="000000000000001000010100001011001101100110101101010110101000010100001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="N":f(u,1)="000000000000000100010011001001100100101010010101001001100100110010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'f(u,0)="O":f(u,1)="000000000000000001000001010001000100100010010001001000100010100000100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="P":f(u,1)="000000000000000111100010001001000100100100011100001000000100000010000000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="Q":f(u,1)="000000000000000001000001010001000100100010010001001000100010100000110000000100000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="R":f(u,1)="000000000000000111100010001001000100100010011110001001000100010010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="S":f(u,1)="000000000000000011110010000001000000011000000010000000100000010011110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="T":f(u,1)="000000000000001111111000100000010000001000000100000010000001000000100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="U":f(u,1)="000000000000000100010010001001000100100010010001001000100100010001110000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="V":f(u,1)="000000000000001000001010001001000100100010001010000101000010100000100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="W":f(u,1)="000000000000001000001100000110010011001001011011001101100100010010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="X":f(u,1)="000000000000000100010001010000101000001000000100000101000100010010001000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
f(u,0)="Y":f(u,1)="000000000000000100010010001000101000010100000100000010000001000000100000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'f(u,0)="Z":f(u,1)="000000000000000111110000001000001000001000000100000100000100000011111000000000000000":f(u,2)=7:f(u,3)=12:u=u+1
'End If

Class CDvCode4
	Public mBuff(), mWidth, mHeight, mCodeTotal, mCode
	Public mMaxMargin, mMinMargin, mMaxWidth, mMinWidth, mMaxHeight, mMinHeight, mAngleMin, mAngleMax
	Public mUsedWidth, mPID180
	Private Sub Class_Initialize
		Randomize
		mCodeTotal	= GetRnd(2,4)	'生成的验证码个数2
		mMaxWidth	= 30			'可取的一个字符的最大宽度
		mMinWidth	= 25			'可取的一个字符的最小宽度
		mMaxHeight	= 28			'可取的一个字符的最大高度
		mMinHeight	= 23			'可取的一个字符的最小高度
		mMaxMargin	= 4				'可取的两个字符间的最大间距0
		mMinMargin	= 2				'可取的两个字符间的最小间距-6
		mWidth		= 140			'生成的图片宽度60
		mHeight		= 28			'生成的图片高度
		mAngleMin	= -5			'最小角度-15
		mAngleMax	= 5				'最大角度15
		mUsedWidth	= GetRnd(5,10)
		mPID180		= 0.01745329                '3.1415926/180
		ReDim mBuff(mWidth, mHeight)
	End Sub
	Public Function GetRnd(iMin, iMax)
		GetRnd = Int((iMax - iMin + 1) * Rnd + iMin)
	End Function
	Public Sub CreateCode
		Dim i, n, iLeft, iTop, iWidth, iHeight
		For i=1 To mCodeTotal
			n			= GetRnd(0, u-1)
			mCode		= mCode & f(n, 0)
			iWidth		= GetRnd(mMinWidth, mMaxWidth)
			iHeight		= GetRnd(mMinHeight, mMaxHeight)
			iLeft		= mUsedWidth+GetRnd(mMinMargin, mMaxMargin)
			iTop		= GetRnd(0, mHeight-iHeight)
			DrawChar	n, iLeft, iTop, iWidth, iHeight
		Next
	End Sub
	Public Sub SetPiex(x, y, c, b)
		If x<0 Or x>mWidth-1 Or y<0 Or y>mHeight-1 Then Exit Sub
		If 1=b Then
			mBuff(x, y)=c
		Else
			Dim xB, xE, yB, yE, t
			t=b/2:xB=x-t:xE=x+t-1:yB=y-t:yE=y+t-1
			For x=xB To xE
				For y=yB To yE
					SetPiex x, y, c, 1
				Next
			Next
		End If
	End Sub
	Public Sub WriteRGB(iR,iG,iB)
		Response.BinaryWrite ChrB(iB) & (ChrB(iG) & ChrB(iR))
	End Sub
	Public Sub DrawChar(iIndex, iLeft, iTop, iWidth, iHeight)
		Dim x, y, iRateX, iRateY, iRealX, iRealY, iRealWidth, sFont, cBit
		sFont		= f(iIndex,1)
		iRealWidth	= f(iIndex,2)
		iRateX		= iRealWidth/iWidth
		iRateY		= f(iIndex,3)/iHeight
		Dim a,cosa,sina,b
		a=GetRnd(mAngleMin,mAngleMax)*mPID180:cosa=Cos(a):sina=Sin(a):b=1
		For x=iWidth-1 To 0 Step -1
			For y=iHeight-1 To 0 Step -1
				iRealX	= Int(iRateX * x)
				iRealY	= Int(iRateY * y)
				cBit	= Mid(sFont, Int(iRealX+iRealWidth*iRealY)+1, 1)
				If "0"<>cBit Then
					SetPiex iLeft+x*cosa-y*sina,iTop+x*sina+y*cosa,1,b
				End If
			Next
		Next
		mUsedWidth	= iLeft+iWidth
	End Sub
	Public Sub DrawPicHead
		Response.Expires = -9999
		Response.AddHeader "pragma", "no-cache"
		Response.AddHeader "cache-ctrol", "no-cache"
		Response.ContentType = "image/bmp"
		Dim iBmpFileSize, iBmpSize
		iBmpSize = mWidth * mHeight * 3
		iBmpFileSize = iBmpSize + 54
		Response.BinaryWrite ChrB(66) & (ChrB(77) & ChrB(iBmpFileSize Mod 256) & ChrB((iBmpFileSize \ 256) Mod 256) & ChrB((iBmpFileSize \ 256 \ 256) Mod 256) & ChrB(iBmpFileSize \ 256 \ 256 \ 256) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0))
		Response.BinaryWrite ChrB(40) & (ChrB(0) & ChrB(0) & ChrB(0) & ChrB(mWidth Mod 256) & ChrB((mWidth \ 256) Mod 256) & ChrB((mWidth \ 256 \ 256) Mod 256) & ChrB(mWidth \ 256 \ 256 \ 256) & ChrB(mHeight Mod 256) & ChrB((mHeight \ 256) Mod 256) & ChrB((mHeight \ 256 \ 256) Mod 256) & ChrB(mHeight \ 256 \ 256 \ 256) & ChrB(1) & ChrB(0) & ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(iBmpSize Mod 256) & ChrB((iBmpSize \ 256) Mod 256) & ChrB((iBmpSize \ 256 \ 256) Mod 256) & ChrB(iBmpSize \ 256 \ 256 \ 256) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0))
	End Sub
	Public Sub CreatePic
		DrawPicHead
		Dim x, y, w, i, iAdd, beginC, add
		w	= mWidth-1
        add = GetRnd(1,20)
        iAdd = add
        beginC = GetRnd(100+add,255-add)
        dim ccr,ccg,ccb
ccr=0
ccg=0
ccb=8


		For y=mHeight-1 To 0 Step -1
            i = beginC
			For x=0 To w
                If i>255-add Then
					iAdd=abs(iAdd)*-1
				ElseIf i<100+add Then
					iAdd=abs(iAdd)
				End If
                i = i + iAdd
				If 1=mBuff(x, y) Then
                    If 1<>0 Then
					    WriteRGB ccr,ccg,ccb
                    Else
                        If GetRnd(0,100)<=30 Then
                            WriteRGB GetRnd(0,255),GetRnd(0,255),GetRnd(0,255)
                        Else
                            WriteRGB 255,255,255
                        End If
                    End If
				Else
                    If 1<>2 Then
					    WriteRGB i,i,i
                    Else
                        If GetRnd(0,100)<=10 Then
                            WriteRGB GetRnd(0,255),GetRnd(0,255),GetRnd(0,255)
                        Else
                            WriteRGB 255,255,255
                        End If
                    End If
				End If
			Next
		Next
	End Sub
End Class
Dim DvCode
Set DvCode	= new CDvCode4
DvCode.CreateCode
DvCode.CreatePic
Config_PSess = Request("Config_PSess")
If Config_PSess="" Then '记录入Session
  Session("ChkCode") = uCase(DvCode.mCode)
Else
  Session(Config_PSess) = uCase(DvCode.mCode)
End If

%>