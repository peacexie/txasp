<%

Function rs_SPPage(xConn,xTabCols,xTabWhere,xPagNow,xPagSize,xKeyID,xKeySort,xRSPage)
  Dim cm
  Set cm = Server.CreateObject("ADODB.Command")
  With cm
    .ActiveConnection = xConn 
    .CommandText = "SPGetPage" '指定存储过程名
    .CommandType = 4 '表明这是一个存储过程
    .Prepared = true '要求将SQL命令先行编译 true false
    .Parameters.Append .CreateParameter("@TabCols" ,200,1,3200,xTabCols)
    .Parameters.Append .CreateParameter("@TabWhere",200,1,3200,xTabWhere)
    .Parameters.Append .CreateParameter("@PagNow"  ,3  ,1,4   ,xPagNow)
    .Parameters.Append .CreateParameter("@PagSize" ,3  ,1,4   ,xPagSize)
    .Parameters.Append .CreateParameter("@KeyID"   ,200,1,255 ,xKeyID)
    .Parameters.Append .CreateParameter("@KeySort" ,200,1,255 ,xKeySort)
    .Parameters.Append .CreateParameter("@PagCount",3  ,4)
    .Parameters.Append .CreateParameter("@RecCount",3  ,3,4   ,1200)
    Set xRSPage(0) = .Execute()
  End With
  If xRSPage(0).State = 0 Then '未取到数据，rs关闭
	xRSPage(1) = -1
	xRSPage(2) = -1
  Else
    xRSPage(0).Close '注意：若要取得参数值，需先关闭记录集对象
	xRSPage(1) = cm("@RecCount")
	xRSPage(2) = cm("@PagCount")
  End If
  Set cm = Nothing
End Function
Function Add_Log(xconn,xUsrID,xAct,xSys,xNote)
  Dim ip,tm,vp,sql,rp
  ip = Get_CIP() 'Request.ServerVariables("REMOTE_HOST")
  tm = Get_yyyymmdd("")&Get_hhmmss()&Get_mSec()&Rnd_ID("KEY",5)
  vp = LeftB(Request.ServerVariables("URL")&"",255)
  xAct  = RequestSafe(xAct&"",3,48)
  sNote = RequestSafe(xNote&" ("&Request.Servervariables("HTTP_USER_AGENT")&")",3,512)
  rp = Request.Servervariables("HTTP_REFERER")
  If inStr(rp,"?")>0 Then rp = Left(rp,inStr(rp,"?"))
  sql = " INSERT INTO AdmLogs (LogTime,LogUser,LogAct,LogPag1,LogPag2,LogSyst,LogNote,LogIP) values "
  sql = sql& "('"&tm&"','"&xUsrID&"','"&xAct&"','"&vp&"','"&rp&"','"&xSys&"','"&sNote&"','"&ip&"')"
  Call rs_DoSql(xconn,sql) 
End Function

Function rs_ETab(xConn,xTab)
    Dim rse
	On Error Resume Next
	Err = 0 
	Set rse = Server.Createobject("Adodb.Recordset")
	rse.open "select * from ["&xTab&"] ",xConn,1,1
	'echo Err 'eNum = cStr(err.number)
	If cStr(Err)="0" Then
	  rs_ETab = true
	else
	  rs_ETab = false ':echo "f.ETab"&xTab
	end if
    rse.Close()
    SET rse = Nothing 
	Err = 0
End Function
Function rs_Row(xcon,xSql) 
   Dim rsv,aVal,i : aVal = ""
   If xcon = "" Then xcon=conn
   Set rsv = Server.Createobject("Adodb.Recordset")
   rsv.Open xSql,xcon,1,1
   IF NOT rsv.EOF THEN
     Redim aVal(rsv.Fields.Count)
	 For i=0 To rsv.Fields.Count-1
	  aVal(i) = rsv(i) ': Response.Write aVal(i)
	 Next
   End IF
   rsv.Close()
   SET rsv = Nothing 
   rs_Row = aVal
End Function
Function rs_Tab(xcon,xSql) 
   Dim rsv,aVal,i : aVal = ""
   If xcon = "" Then xcon=conn
   Set rsv = Server.Createobject("Adodb.Recordset")
   rsv.Open xSql,xcon,1,1
   IF NOT rsv.EOF THEN
     aVal = rsv.GetRows() '记录数:UBound(aVal,2)+1，栏位数:UBound(aVal, 1)+1。
   End IF
   rsv.Close()
   SET rsv = Nothing 
   rs_Tab = aVal
End Function

Function rs_Exist(xconn,xSql) 
   Dim rsExt,ObjExt
   Set rsExt=Server.Createobject("Adodb.Recordset")
   rsExt.Open xSql,xconn,1,1
   IF rsExt.EOF THEN
      ObjExt = "EOF"
   ELSE
      ObjExt = "YES"
   End IF
   rsExt.Close()
   SET rsExt = Nothing
   rs_Exist = ObjExt   
End Function
Function rs_Count(xconn,xTab)   
   Dim rsCnt,ObjCnt,sql 
   If inStr(xTab ,"")<0 Then xTab="["&Trim(xTab)&"]"
   sql  = " SELECT COUNT(*) AS Exprs_Count FROM "&xTab&" " 
   Set rsCnt = Server.Createobject("Adodb.Recordset")
   rsCnt.open sql,xconn
   ObjCnt = rsCnt("Exprs_Count")
   rsCnt.Close()
   SET rsCnt = Nothing
   rs_Count = ObjCnt
End Function 
Function rs_Val(xcon,xSql) 
   Dim rsv,sVal : sVal = ""
   If xcon = "" Then xcon=conn
   Set rsv = Server.Createobject("Adodb.Recordset")
   'Response.Write xSql&"))"
   rsv.Open xSql,xcon,1,1
   IF NOT rsv.EOF THEN
      sVal = rsv(0)
   End IF
   rsv.Close()
   SET rsv = Nothing 
   rs_Val = sVal
End Function
Function rs_DoSql(xconn,xSql) 
   Dim cnFPubcnDo 
   Set cnDo = Server.CreateObject("Adodb.Connection")
   cnDo.Open xconn 
   cnDo.Execute(xSql)
   cnDo.Close()
   Set cnDo = Nothing
End Function

Function rs_OrdID(xconn,xType,xPrev,xExt) 
  Dim tKey,tID,tSql,dbKey,tOld,tMin,uExt,uKey
  tKey = Get_FmtID(xType,"")
  tID = Left(tKey,xPrev) 
  tSql = "SELECT TOP 1 KeyCode FROM OrdInfo WHERE KeyCode LIKE '"&tID&"%' ORDER BY LogATime DESC"
  dbKey = rs_Val(xconn,tSql) ':echo dbKey
  If dbKey="" Then
	rs_OrdID = tKey&Get_9999ID(1,xExt)
  Else
	tOld = Right(dbKey,xExt)    :tMin = Get_9999ID(1,xExt)
	uExt = Next_ID(tOld,tMin,1) :uKey = tKey&uExt
	rs_OrdID = uKey ':echo tOld&"-"&tMin&"-"&uExt
  End If
  '重复检查???
End Function

Function rs_AutID(xconn,xTab,xCol,xModPath,xType,xTime) 
Dim sYYYY,sMD,tSql
  If xModPath&"" = "" Then ':Inf,Pic,job,trade,love,bbs,
    xModPath = "dat12" 
  Else
    xModPath = Replace(xModPath,"-","")
  End If
  If xTime&"" = "" Then xTime = Now()
  sYYYY = Left(Year(xTime),4) 
  If xType="" Or xType="1" Then
    sMD = Get_FmtID("md-hnsx",xTime) 
  Else
    sMD = Get_FmtID(lCase(xType),xTime)
  End If
  tKey = xModPath&"-"&sYYYY&"-"&sMD&"."&Base_32(Session.SessionID(),32,3,"Right") 
  tKey = tKey&Rnd_ID("A",24-Len(tKey))
  tSql = "SELECT TOP 1 "&xCol&" FROM "&xTab&" WHERE "&xCol&" LIKE '%"&tKey&"%' "
  If rs_Val(xconn,tSql)="" Then
    rs_AutID = tKey
  Else
    Call rs_AutID(xconn,xTab,xCol,xModPath,xType,xTime)
  End If 
End Function

Function rs_TypID(xconn,xMod,xNowID,xDeep)
  Dim FstID,DefID,ExtID,MaxID
  FstID = Mid(xMod,4,2)&xDeep 
  DefID = Next_ID(xNowID,FstID&"0012",4)
  If rs_Exist(xconn,"SELECT TypID FROM WebTyps WHERE TypID='"&DefID&"'")="YES" Then
	MaxID = rs_Val(xconn,"SELECT TOP 1 TypID FROM WebTyps WHERE TypID LIKE '"&FstID&"%' ORDER BY TypID DESC ")
	DefID = Next_ID(MaxID,FstID&"0012",4)
  End If
  rs_TypID = DefID
End Function




%>
