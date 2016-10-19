
<%


Function ShowText(xStr)
  Dim sHtml
  sHtml = xStr
  sHtml = Replace(sHtml,"<","&lt;")
  sHtml = Replace(sHtml,">","&gt;")
  ShowText = sHtml
End Function


' 提取连接函数
Function OutLinks(xStr,xPat)
  Dim regEx, iMatch, Matches,rStr ' 建立变量。
  rStr = ""
  Set regEx = New RegExp          ' 建立正则表达式。
  regEx.IgnoreCase = True         ' 设置是否区分字符大小写。
  regEx.Global = True             ' 设置全局可用性。 <img name="s" ... />
  regEx.Pattern = xPat '"<a[^>]*(href=[^>]*)[^>]*>([^<]*|<img[^<]*>)</a>" ' 设置模式。
  Set Matches = regEx.Execute(xStr)       ' 执行搜索。
  For Each iMatch In Matches              ' 遍历匹配集合。
    rStr = rStr & iMatch.Value & "||" 
  Next
  'regEx.Pattern = "<a[^<]*href=['|""]([^<]*)['|""][^<]*>([^<]*)</a>"
  'rStr = regEx.replace(rStr,"<a href='$1' target='_blank'>$2</a>")
  Set regEx = nothing

  OutLinks = rStr
End Function

Function OutFAdd(xPName,xCont,xCSet)
 Dim st 
 If instr(xPName,":")<=0 Then xPName=server.MapPath(xPName)  
 Set st=Server.CreateObject("ADODB.Stream")   
  st.Type=2   
  st.Mode=3   
  st.Charset=xCSet  
  st.Open()   
  st.WriteText xCont   
  st.SaveToFile xPName,2   
  st.Close()  
 Set st=Nothing   
End Function

Function OutSPos(xCont,xStr)
  Dim vCont,vStr
  vCont=uCase(xCont)
  vStr=uCase(xStr)
  OutSPos=inStr(vCont,vStr)
End Function

Function OutPage(Path,xCharSet)
Dim tData,tLen
  tData = OutBody(Path) 
  tLen = Len(tData) ':Response.Write tLen
  if tLen =< 720 then 
    OutPage = OutToStr(tData,xCharSet) 'tData
  else
    OutPage = OutToStr(tData,xCharSet)
  end if 
End function

Function OutBody(url) 
  on error resume next
  Set Retrieval = CreateObject("Micro"&"soft.XML"&"HTTP") 
  With Retrieval 
  .Open "Get", url, False, "", "" 
  .Send 
  OutBody = .ResponseBody
  End With 
  Set Retrieval = Nothing 
End Function

Function OutToStr(body,Cset)
  dim objstream
  set objstream = Server.CreateObject("ado"&"db.stream")
  objstream.Type = 1
  objstream.Mode = 3
  objstream.Open
  objstream.Write body
  objstream.Position = 0
  objstream.Type = 2
  objstream.Charset = Cset
  OutToStr = objstream.ReadText 
  objstream.Close
  set objstream = nothing
End Function

%>