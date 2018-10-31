<!--#include file = "crypt.asp"-->
<!--#include file = "CheckPwdHandler.asp"-->
<%
	Response.Expires = -1					
	Response.ExpiresAbsolute = Now() -1
	Response.cachecontrol = "no-cache"

'定义常量
g_LinesPerPage = 0
g_DecimalDigits = 2

YDay = 0	'说明：凯路计算预收汇日期的方法：付款条件内的X天 + YDay(天) + 提单日期 

'通用函数
'函数说明
'  货币翻译	TranCurrency(值，"国家")
'  日期翻译	TranDate(日期)
'函数体
Function ThreeNum(n,H)
  Dim NumName(30)
  NumName(0)=""
  NumName(1)="ONE"
  NumName(2)="TWO"
  NumName(3)="THREE"
  NumName(4)="FOUR"
  NumName(5)="FIVE"
  NumName(6)="SIX"
  NumName(7)="SEVEN"
  NumName(8)="EIGHT"
  NumName(9)="NINE"
  NumName(10)="TEN"
  NumName(11)="ELEVEN"
  NumName(12)="TWELVE"
  NumName(13)="THIRTEEN"
  NumName(14)="FOURTEEN"
  NumName(15)="FIFTEEN"
  NumName(16)="SIXTEEN"
  NumName(17)="SEVENTEEN"
  NumName(18)="EIGHTEEN"
  NumName(19)="NINETEEN"
  NumName(22)="TWENTY"
  NumName(23)="THIRTY"
  NumName(24)="FORTY"
  NumName(25)="FIFTY"
  NumName(26)="SIXTY"
  NumName(27)="SEVENTY"
  NumName(28)="EIGHTY"
  NumName(29)="NINETY"
  if Len(n)>2 then
    c1=Left(n,1)
  end if
  if Len(n)>1 then
    c2=Left(Right(n,2),1)
  end if
  c3=Right(n,1)
  if c1>0 then
    ThreeStr=NumName(c1)&" HUNDRED"
  end if
  if (c2>0 or c3>0) and (H=1 or c1>0) then
    ThreeStr=ThreeStr&" AND "
  end if
  if c2>1 then
    ThreeStr=ThreeStr&NumName(c2+20)
    if c3>0 then
      ThreeStr=ThreeStr&" "&NumName(c3)
    end if
  else
    if c2>0 or c3>0 then
      ThreeStr=ThreeStr&NumName(c2*10+c3)
    end if
  end if
  ThreeNum=ThreeStr
End Function
Function usInteger(n)
 Dim NumStrw(30)
  NumStrw(1)="ONE"
  NumStrw(2)="TWO"
  NumStrw(3)="THREE"
  NumStrw(4)="FOUR"
  NumStrw(5)="FIVE"
  NumStrw(6)="SIX"
  NumStrw(7)="SEVEN"
  NumStrw(8)="EIGHT"
  NumStrw(9)="NINE"
  Pos=3
  PStr=""
  ThrNum=""
  NumStr=""
  if InStr(n,".") then
	XP=InStr(n,".")
	IntN=Left(n,XP-1)
	PointN=Right(n,Len(n)-Xp)
    PStr=" POINT"
	for i=1 to Len(PointN)
		SS=Mid(PointN,i,1)
		PStr=PStr&" "& NumStrw(Cint(ss))
	next
	n=IntN
  end if
  if n<>"" then
  Do While Pos<Len(n)
    ThrNum=Mid(n,Len(n)-Pos+1,3)
    ThrNum=ThreeNum(ThrNum,1)
    Select case Pos
      case 6
        ThrNum=ThrNum&" THOUSAND"
      case 9
        ThrNum=ThrNum&" MILLION"
      case 12
        ThrNum=ThrNum&" BILLION"
      case 15
      ThrNum=ThrNum&" TRILLION"
    end select
    if Len(NumStr)>0 then
      NumStr=ThrNum&" "&NumStr
    else
      NumStr=ThrNum
    end if
    Pos=Pos+3
  Loop
  ThrNum=Left(n,Len(n)-Pos+3)
  ThrNum=ThreeNum(ThrNum,0)
  Select case Pos
    case 6
      ThrNum=ThrNum&" THOUSAND"
    case 9
      ThrNum=ThrNum&" MILLION"
    case 12
      ThrNum=ThrNum&" BILLION"
    case 15
      ThrNum=ThrNum&" TRILLION"
  end select
  if Len(NumStr)>0 then
    NumStr=ThrNum&" "&NumStr
  else
    NumStr=ThrNum
  end if
end if
  usInteger=NumStr&PStr
End Function
Function usSuffix(n)
  Dim NumSuffix(9)
  NumSuffix(0)="th"
  NumSuffix(1)="st"
  NumSuffix(2)="th"
  NumSuffix(3)="rd"
   if Len(n)>0 then
    x=Right(n,1)
    if x<4 then
      usSuffix=NumSuffix(x)
    else
      usSuffix="th"
    end if
	y=Right(n,2)
	if y=11 or y=12 or y=13 then
		usSuffix="th"
	end if
  end if
End Function
Function cnInteger(n)
  Dim NumName(10)
  NumName(0)="零"
  NumName(1)="壹"
  NumName(2)="贰"
  NumName(3)="叁"
  NumName(4)="肆"
  NumName(5)="伍"
  NumName(6)="陆"
  NumName(7)="柒"
  NumName(8)="捌"
  NumName(9)="玖"
  Dim PosName(11)
  PosName(0)=""
  PosName(1)="拾"
  PosName(2)="佰"
  PosName(3)="仟"
  PosName(4)=""
  PosName(5)="拾"
  PosName(6)="佰"
  PosName(7)="仟"
  PosName(8)=""
  PosName(9)="拾"
  PosName(10)="佰"
  PosName(11)="仟"

  Num=Int(n)
  l=Len(Num)
  o=0
  For i=1 To l
    c=Mid(Num,i,1)
    if c=0 then
      o=o+1
    else
      if o>0 then
        NumStr=NumStr&"零"
      end if
      o=0
      NumStr=NumStr&NumName(c)&PosName(l-i)
    end if
    if l-i=4 then
      NumStr=NumStr&"万"
    end if
    if l-i=8 then
      NumStr=NumStr&"亿"
    end if
  Next
  cnInteger=NumStr
End Function
Function TranCurrency(n,country)
  Select Case country
    case "us"
      NumStr=""
      Pos=InStr(n,".")
      if Pos=0 then
        Pos=Len(n)+1
      end if
      if Pos>1 then
        NumStr=usInteger(Left(n,Pos-1))
      end if
      if InStr(n,".")<Len(n) and InStr(n,".")>0 then
        c=Mid(CStr(n),InStr(n,".")+1,2)
        if Len(c)=1 then
           c=c&"0"
        end if
        if c="01" then
          NumStr=NumStr&" AND CENT ONE"
        else
          NumStr=NumStr&" AND CENTS "&usInteger(c)
       end if
      end if
      TranCurrency=NumStr&" ONLY"
    case "cn"
      NumStr=""
      Pos=InStr(n,".")
      if Pos=0 then
        Pos=Len(n)+1
      end if
      if Pos>1 then
        NumStr=cnInteger(Left(n,Pos-1))&"元"
      end if
      if Pos<Len(n) and Pos>0 then
        c=Mid(n,Pos+1,1)
        if c>0 then
          NumStr=NumStr&cnInteger(c)&"角"
        end if
        if Pos+1<Len(n) then
          c=Mid(n,Pos+2,1)
        else
          c=0
        end if
        if c>0 then
          NumStr=NumStr&cnInteger(c)&"分"
        end if
      end if
      TranCurrency=NumStr&"整"
  End Select
End Function
Function TranDate(D)
  Dim MonthName(12)
  MonthName(1)="Jan."
  MonthName(2)="Feb."
  MonthName(3)="Mar."
  MonthName(4)="Apr."
  MonthName(5)="May."
  MonthName(6)="Jun."
  MonthName(7)="Jul."
  MonthName(8)="Aug."
  MonthName(9)="Sep."
  MonthName(10)="Oct."
  MonthName(11)="Nov."
  MonthName(12)="Dec."
  DateStr=D
  if IsDate(D) then
    DateStr=MonthName(Month(D))&Day(D)&","&Year(D)
  end if
  TranDate=UCase(DateStr)
End Function
Function GetItemValue(RSDB,ItemField,ValueField,Item)
  i=0
  GetItemValue=""
  do while not RSDB.eof
    if Item=RSDB(ItemField) then
      GetItemValue=RSDB(ValueField)
      RSDB.MoveFirst
      Exit Function
    end if
    RSDB.MoveNext
    i=i+1
  loop
  if i>0 then
    RSDB.MoveFirst
  end if
End Function
	Function PutItemValue(RSDB,ItemField,ValueField,Item,ItemValue)
	  PutItemValue=1
	  i=0
	  if not RSDB.Eof then
		RSDB.MoveFirst
	  end if
	  do while not RSDB.eof
			if Item=RSDB(ItemField) then
				  RSDB(ValueField)=ItemValue
				  RSDB.UpDate
				  PutItemValue=0
				  RSDB.MoveFirst
				  Exit Function
			end if
		RSDB.MoveNext
		i=i+1
	  loop
	  if i>0 then
		RSDB.MoveFirst
	  end if
	End Function

	Function FUnitChs(NameID)
		Set RSunitchs=Server.CreateObject("ADODB.RecordSet")
		RSunitchs.Open "select UnitNameChs from UnitName where pUnitName='"&NameID&"'",G_DBConn,2,3,1
			if not RSunitchs.eof then
				FUnitChs=RSunitchs("UnitNameChs")
			else
				FUnitChs=""
			end if
		RSunitchs.close
		Set RSunitchs=Nothing
	End Function
'程序
svrName=Request.ServerVariables("SERVER_NAME")
Set g_DBConn = Server.CreateObject("ADODB.Connection")
g_DBConn.Open Application("ConnectionString")
FromUrl=Request.ServerVariables("LOCAL_ADDR")
If FromUrl="192.168.6.201" Then
	'=====国外用户
	SignInName=Trim(Request("SignInName"))
	SignInPw=Trim(Request("SignInPw"))
	SignInConPw=Trim(Request("SignInConPw"))
	If SignInName="" or SignInPw="" or SignInConPw="" Then
		If Session("SignInName")="" Then
			Response.Redirect("../index.asp")
			Response.End()
		End If	
	Else
			Set RSTemp=g_DBConn.Execute("Select CustCode,CustId From Customer Where SignInName='"&SignInName&"' And SignInPw='"&SignInPw&"' And SignInConPw='"&SignInConPw&"'")
			If RSTemp.Eof Then
				response.write "<body background='../img/gback.jpg' onDBLclick='history.back()'>"
				response.write "<div align='center'> "
				response.write "<p>&nbsp;</p><table width='83%' border=1>"
				response.write "<tr> "
				response.write "<td>"
				response.write "<p align='center'><b><font size=6 Color=red>Error UserName or Password Or Confirmed P.W!</font></b></p>"
				response.write "<p align='center'><font size=6><b><font face='楷体'>The User Name Or Password Or Confirmed P.W Is Wrong!</font></b></font></p>"
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
				response.write "</div>"
				response.write "</body>"
	
				RSTemp.Close
				Response.end
			Else
				Session("SignInName")=SignInName
				Session("SignInCustId")=RSTemp("CustId")
				Session("SignInCustCode")=RSTemp("CustCode")
			End If
			RSTemp.Close
			Set RSTemp=Nothing
	End If
SignInCustId=Session("SignInCustId")
SignInCustCode=Session("SignInCustCode")		
Else
	'=====国内用户
    UserName=Trim(Request("name"))
    Password=Request("pwd")
    if UserName="" or Password="" then
		if Session("UserId")="" then
%>
<head> 
<meta http-equiv="windows-Target" contect="_top"> 
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<SCRIPT LANGUAGE="JavaScript"> 
if (top.frames.length!=0) 
	if (window.name == "topFrame")		//如果页面显示在顶部的窗口内(TopManu.asp)，将于Default.asp在同一文件夹内
	{
		top.location = "Default.asp"
	}
	else
	{
		top.location="../Default.asp"; 
	}
else
	document.location="../Default.asp";
</SCRIPT> 
</head> 
<%
			'response.redirect "/default.asp"
   			Response.end
		end if 
  	else
		Set RSVian = g_DBConn.Execute("SELECT EmpId, EmpCode,Grade,DeptCode FROM employee WHERE EmpCode='"&UserName&"' AND PassWord='"&md5(Password)&"'")
		if RSVian.eof then
			response.write "<body background='../img/gback.jpg' onDBLclick='history.back()'>"
			response.write "<div align='center'> "
		    response.write "<p>&nbsp;</p><table width='83%' border=1>"
			response.write "<tr> "
			response.write "<td>"
			response.write "<p align='center'><b><font size=6>Error username or password!</font></b></p>"
			response.write "<p align='center'><font size=6><b><font face='楷体'>错误的用户名或口令！</font></b></font></p>"
			response.write "</td>"
			response.write "</tr>"
		    response.write "</table>"
			response.write "</div>"
			response.write "</body>"

			RSVian.Close
	  		Response.end
		else
			Session("UId") = RSVian("EmpId")		'备忘录模块需要
	  		Session("UserID")=UCase(UserName)
	  		Session("UserGrade")=RSVian("Grade")
	  		Session("UserDept")=RSVian("DeptCode")
		end if
		RSVian.Close
		Set RSVian=Nothing
  	end if

'=======================================
	'根据菜单中检测数据表Sys_Module(业务管理菜单条目)中是否有排产信息。五金类的需要有排产条目，化工类不需要。
	'AFFAIRTYPE是全局变量，标志业务类型。五金类(METAL)、化工类(CHEM)、综合类(UNITE)
	'如果业务类型是五金类，则把 AFFAIRTYPE 赋值“METAL”，在新建预算单时合同类型默认“五金类”。
		'Session("AffairType")="UNITE"
		'Session("AffairType")="METAL"
		Session("AffairType")="CHEM"
		
		AffairType=Session("AffairType")
'=======================================

'读取初使值
UserID=Session("UserID")
UDept=Session("UserDept")
UGrade=Session("UserGrade")
End If
	
	'判断使用权限
	Set gRS = Server.CreateObject("ADODB.Recordset")
'==============================================================================================================================
'	计算当前页使用的模块代码
'==============================================================================================================================
	call GetModuleCode()
	
'==============================================================================================================================
'	判断使用权限
'==============================================================================================================================
	Set gRS = Server.CreateObject("ADODB.Recordset")

	if ModuleCode <> "" then
		gRS.Open "SELECT A.EmpCode FROM Employee A  "&_
				 "WHERE NOT EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
				 "			WHERE R.EmpCode = A.EmpCode AND ISNULL(G.DenyModuleRight, '') LIKE '%"& Left(ModuleCode, 2) &",%') "&_
				 "		AND ISNULL(A.DenyModuleRight, '') NOT LIKE '%"& Left(ModuleCode, 2) &",%' "&_
				 "		AND (EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
				 "			WHERE R.EmpCode = A.EmpCode AND ISNULL(G.ModuleRight, '') LIKE '%"& Left(ModuleCode, 2) &",%') "&_
				 "			OR ISNULL(A.ModuleRight, '') LIKE '%"& Left(ModuleCode, 2) &",%') "&_
				 "		AND A.EmpCode = '"& UserId &"'", G_DBConn, 0, 1, 1
		If gRS.EOF Then
			response.write "<body background='../img/gback.jpg' onDBLclick='history.back()'>"
			response.write "<div align='center'> "
		    response.write "<p>&nbsp;</p><table width='83%' border=1>"
			response.write "<tr> "
			response.write "<td>"
			response.write "<p align='center'><font size=6><b><font face='楷体'>对不起，您未被授权使用这个功能！</font></b></font></p>"
			response.write "</td>"
			response.write "</tr>"
		    response.write "</table>"
			response.write "</div>"
			response.write "</body>"
			Response.end
		End If
		gRS.Close
	end if
	set gRS = nothing
	
	'错误信息函数
	function ErrMsg(Msg)
		Response.Write "<body onclick='history.back()'>"
		Response.Write "<center><font color=red>"&Msg&"</font></center>"
		response.write "</body>"
		Response.End
	end function

	''报运部分所用到函数*******************
	Function AddSpace(Source,Number)	'左对齐函数
		if IsNull(Source) then
			Exit Function
		end if
		If Len(Source)>=Number then
			AddSpace=Source&" "
		else
			AddSpace=Source&Space(Number-Len(Source))
		End if
	End Function

	Function AddSpaceR(Source,Number)	'右对齐函数
		if IsNull(Source) then
			Exit Function
		end if
		If Len(Source)>=Number then
			AddSpaceR=Source&" "
		else
			AddSpaceR=Space(Number-Len(Source))&Source
		End if
	End Function

	Sub WriteProdDesc(aDocuID,aBillMark,aProdDesc)
		If aBillMark<>"BillCust" and aBillMark<>"BillPack" and aBillMark<>"BillPackF" and aBillMark<>"BillInv" and aBillMark<>"BillInvF" and aDocuId="" then
			Response.Write "WriteProdDesc传入参数错误!"
			Response.End
		end if
		Set RSVian=Server.CreateObject("ADODB.RecordSet")
			RSVian.Open "Select * From DocuProdDesc Where DocuID='"&aDocuID&"' and BillMark='"&aBillMark&"'",g_DBConn,2,3,1
				If Not RSVian.Eof then
					RSVian("ProdDesc")=aProdDesc
				else
					RSVian.AddNew
					RSVian("DocuId")=aDocuId
					RSVian("BillMark")=aBillMark
					RSVian("ProdDesc")=aProdDesc
				end if
				RSVian.UpDate
			RSVian.Close
		Set RSVian=Nothing
	End Sub

	Function ReadProdDesc(aDocuID,aBillMark)
		Set RSVian=Server.CreateObject("ADODB.RecordSet")
			RSVian.Open "Select ProdDesc From DocuProdDesc Where DocuID='"&aDocuID&"' and BillMark='"&aBillMark&"'",g_DBConn,2,3,1
				If Not RSVian.Eof then
					ReadProdDesc=RSVian("ProdDesc")
				else
					ReadProdDesc=""
				End if
			RSVian.Close
		Set RSVian=Nothing
	End Function
	

	sub GetModuleCode()
		'当请求可位于两个以上模块的页面时，需传递模块参数 表Sys_Module中ModuleCode（如:ModuleCode=AB）
		if Request("ModuleCode") <> "" then
			Session("ModuleCode") = Request("ModuleCode")		'接收并保存模块代码
		end if
	
		'如果页面内的 ModuleCode 不确定(包含*号),根据 Session("ModuleCode") 内的值计算出当前所需的模块代码
		i = Instr(1, ModuleCode, "*", 1)
		do while i > 0
			ModuleCode = Replace(ModuleCode, "*", Mid(Session("ModuleCode"), i, 1), 1, 1, 1)
			i = Instr(1, ModuleCode, "*", 1)
		loop
	
		'最后更新 Session("ModuleCode") 为最新的当前页使用的模块代码,如果模块代码大于两位,不包括最后两位数字
		if ModuleCode <> "" then
			if Len(ModuleCode) > 2 then				'if ModuleCode = "AA01" then Session("ModuleCode") = "AA"
				Session("ModuleCode") = Left(ModuleCode, Len(ModuleCode) - 2)
			else									'if ModuleCode = "AA" then Session("ModuleCode") = "AA"
				Session("ModuleCode") = ModuleCode
			end if
		end if
	end sub


'================================================================================================================================
'	合并字符串
'================================================================================================================================
function JointStr(DestStr, SourceStr, Delimiter)
	if DestStr = "" or IsNull(DestStr) then
		JointStr = SourceStr
	else
		if SourceStr <> "" then
			JointStr = DestStr & Delimiter & SourceStr
		else
			JointStr = DestStr
		end if
	end if
end function


'================================================================================================================================
'	计算金额(包括多币别)的和(Sort = "Add")或差值(Sort = "Sub")	( "USD10 + RMB100" +|- "USD5 + EUR20" ), 
'================================================================================================================================
function ComputeAmt(Amt1, Amt2, Delimiter, Sort)
	dim arrAmt1,arrAmt2,Der, ZeroCurr, Sign
	dim A1,A2, subed, NegativeDelimiter, arrTemp, strTemp, IsAddDelimiter
	
	if Sort = "+" then Sort = "Add"
	if Sort = "-" then Sort = "Sub"
	
	'记录分隔符是否包含 + 
	if InStr(1, Delimiter, "+", 1) > 0 then
		IsAddDelimiter = true
	end if
	
	'如果用加号作分隔符,如(USD10.00 + JPY20.00 - RMB5.00)负数的符号可能替换分隔符中的加号，以负号为分隔符取得其中的负数
	if IsAddDelimiter then
		arrAmt1 = SplitAmt(Amt1, Delimiter)
		arrAmt2 = SplitAmt(Amt2, Delimiter)
	else
		arrAmt1 = Split(Amt1, Delimiter, -1, 1)
		arrAmt2 = Split(Amt2, Delimiter, -1, 1)
	end if

	Der = ""
	Amt1ExistNoCur = false		'判断金额Amt1中是否存在没有币别的金额
	for each A1 in arrAmt1
		if IsNumeric(A1) then
			CA1 = ""			'取得金额中的币别字符
			VA1 = CCur(A1)		'取得金额中的数字值
			Amt1ExistNoCur = true
		else
			CA1 = Left(A1, 3)	
			VA1 = CCur(Right(A1, len(A1) - 3))
		end if

		'查找金额Amt2中是否有与金额Amt1中相同币别的数量，如果有则计算
		for each A2 in arrAmt2
			if IsNumeric(A2) then
				CA2 = ""
				VA2 = CCur(A2)
			else
				CA2 = Left(A2, 3)
				VA2 = CCur(Right(A2, len(A2) - 3))
			end if

			if CA1 = CA2 then
				if Sort = "Add" then
					ComputeAmt = ComputeAmt & MakeAmt(IsAddDelimiter, Der, CA1, VA1 + VA2)
				else
					'做减法时,如果一个币别的结果为0,不输出此币别
					if VA1 - VA2 <> 0 then 
						ComputeAmt = ComputeAmt & MakeAmt(IsAddDelimiter, Der, CA2, VA1 - VA2)
					else
						ZeroCurr = ZeroCurr &"<"& CA1 &">"		'记录相减后金额为0的币别
					end if
				end if
				subed = 1		'已经输出了此币别的金额
				Exit for
			end if			
		next

		'如没有则输出金额Amt1中的数量
		if subed <> 1 then
			ComputeAmt = ComputeAmt & MakeAmt(IsAddDelimiter, Der, CA1, VA1)
		end if
		
		subed = 0
		if ComputeAmt <> "" then
			Der = Delimiter				'间隔符
		end if
	next
	
	'输出币别只在金额Amt2中的数量
	if Sort = "Sub" then
		Sign = -1
	else
		Sign = 1
	end if
	
	for each A2 in arrAmt2
		if IsNumeric(A2) then
			CA2 = ""
			VA2 = CCur(A2)
		else
			CA2 = Left(A2, 3)
			VA2 = CCur(Right(A2, len(A2) - 3))
		end if

		if (Instr(1, ComputeAmt, CA2, 1) = 0 or not Amt1ExistNoCur and CA2 = "") and InStr(1, ZeroCurr, "<"& CA2 &">", 1) = 0 then
			ComputeAmt = ComputeAmt & MakeAmt(IsAddDelimiter, Der, CA2, VA2 * Sign)
		end if

		if ComputeAmt <> "" then
			Der = Delimiter
		end if
	next
end function


'================================================================================================================================
'	筛选用户可以查看的数据	数据用户代码字段名(EmpCode), 数据客户ID字段名(CustCode)
'================================================================================================================================
function CanSee(EmpCode, CustCode)
	'从系统的数据权限处获得给当前用户设置的数据权限
	CanSee = CanSeeEx(EmpCode, CustCode, UserId, ModuleCode)
end function

'================================================================================================================================
'	筛选指定用户(UseEmpCode)在指定模块(UseModuleCode)可以查看的数据(返回SQL(用于WHERE后)串)
'	数据用户代码字段名(EmpCode), 数据客户ID字段名(CustCode)
'================================================================================================================================
function CanSeeEx(EmpCode, CustCode, UseEmpCode, UseModuleCode)
	'从系统的数据权限处获得给当前用户设置的数据权限
	CanSeeEx = SeeEmpEx(EmpCode, "Sel", UseEmpCode, UseModuleCode) &" AND "& SeeCustEx(CustCode, "Sel", UseEmpCode, UseModuleCode)
end function

'================================================================================================================================
'	筛选用户可以查看的职员	(Sort="Sel")时返回SQL(用于WHERE后)串,或(Sort="Edi|Del|Chk")时返回布尔值
'================================================================================================================================
function SeeEmp(EmpCode, Sort)
	SeeEmp = SeeEmpEx(EmpCode, Sort, UserId, ModuleCode)
end function

'================================================================================================================================
'	筛选用户可以查看的职员(返回SQL串)	数据用户代码字段名(EmpCode), 判断类型(Sort):"Sel" | "Edi" | "Del" | "Chk"
'	Sort = "Sel"时 EmpCode:职员代码字段名如:"A.EmpCode", Sort = "Edi"|"Del"|"Chk"时 EmpCode:当前记录的职员代码如:'04'
'	要检测的用户代码(UseEmpCode), 要检测的模块代码(UseModuleCode)
'================================================================================================================================
function SeeEmpEx(EmpCode, Sort, UseEmpCode, UseModuleCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	CurMCode = Left(UseModuleCode, 2)
	EmpY = Sort &"Y"
	EmpN = Sort &"N"

	'当前用户对当前模块的数据权限
	RS.Open "SELECT REPLACE( (CASE WHEN ISNULL("& EmpY &", '') = ',' THEN '' ELSE ISNULL("& EmpY &", '') END), '_Self', '"& UseEmpCode &"') "& EmpY &", "&_
			"		REPLACE( (CASE WHEN ISNULL("& EmpN &", '') = ',' THEN '' ELSE ISNULL("& EmpN &", '') END), '_Self', '"& UseEmpCode &"') "& EmpN &" "&_
			"FROM Sys_DataRight WHERE EmpCode = '"& UseEmpCode &"' AND "&_
			"	(SELECT ModuleRight FROM Employee WHERE EmpCode = '"& UseEmpCode &"') LIKE '%"& CurMCode &",%'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		EY = RS(EmpY)
		EN = RS(EmpN)
	end if
	RS.Close 

	'当前用户所在组对当前模块的所有数据权限
	RS.Open "SELECT REPLACE( (CASE WHEN ISNULL(R."& EmpY &", '') = ',' THEN '' ELSE ISNULL(R."& EmpY &", '') END), '_Self', '"& UseEmpCode &"') "& EmpY &", "&_
			"		REPLACE( (CASE WHEN ISNULL(R."& EmpN &", '') = ',' THEN '' ELSE ISNULL(R."& EmpN &", '') END), '_Self', '"& UseEmpCode &"') "& EmpN &" "&_
			"FROM EmployeeRole A LEFT JOIN Sys_DataRight R ON A.GroupCode = R.GroupCode "&_
			"WHERE A.EmpCode = '"& UseEmpCode &"' AND (SELECT ModuleRight FROM EmployeeGroup WHERE GroupCode = A.GroupCode) LIKE '%"& CurMCode &",%'", G_DBConn, 0, 1, 1
	do while not RS.EOF
		GY = Split(RS(EmpY), ",", -1, 1)
		GN = Split(RS(EmpN), ",", -1, 1)

		'取得有效的数据权限最大值
		for each i in GY
			if InStr(1, EY, i, 1) = 0 then
				EY = EY & i &","
			end if
		next
		'取得禁止的数据权限最大值
		for each i in GN
			if InStr(1, EN, i, 1) = 0 then
				EN = EN & i &","
			end if
		next
		RS.MoveNext
	loop
	RS.Close
	
	'取得最终效有效的数据权限
	if EN <> "" then
		EGN = Split(EN, ",", -1, 1)
		for each i in EGN
			if i <> "" then
				EY = Replace(EY, i &",", "", 1, -1, 1)
			end if
		next
	end if

	if EY <> "" then
		EY = Left(EY, Len(EY) -1)			'当EY 有值时,去处最右侧的","
	end if
	
	if Sort = "Sel" then		'判断能否查看时,返回用于SQL(WHERE后)的限制条件(True | False)
		if EY <> "" then
			'为职员代码加入单引号 "02,04,06" -> "'02','04','06'"
			aEY = Split(EY, ",", -1, 1)
			for i = 0 to UBound(aEY)
				aEY(i) = "'"& aEY(i) &"'"
			next
			sEY = Join(aEY, ",")

			SeeEmpEx = EmpCode &" IN ("& sEY &")"
		else
			SeeEmpEx = "1 <> 1"
		end if
	else						'判断能否修改、删除、审核时,返回输出给脚本用的布尔值, Line706
		if InStr(1, EY, EmpCode, 1) > 0 then
			SeeEmpEx = True
		else
			SeeEmpEx = False
		end if
	end if

	set RS = nothing
end function


'================================================================================================================================
'	筛选用户可以查看的客户	(Sort="Sel")时返回SQL(用于WHERE后)串,或(Sort="Edi|Del|Chk")时返回SQL(用于SELECT后最终用于脚本的布尔值)串
'================================================================================================================================
function SeeCust(CustCode, Sort)
	SeeCust = SeeCustEx(CustCode, Sort, UserId, ModuleCode)
end function


'================================================================================================================================
'	筛选用户可以查看的客户(返回SQL串)	数据客户ID字段名(CustCode),  判断类型(Sort):"Sel" | "Edi" 
'	CustCode: 客户ID字段名
'	要检测的用户代码(UseEmpCode), 要检测的模块代码(UseModuleCode)
'================================================================================================================================
function SeeCustEx(CustCode, Sort, UseEmpCode, UseModuleCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	CurMCode = Left(UseModuleCode, 2)
	CustY = "Cust"& Sort &"Y"

	'当前用户对当前模块的的客户权限,
	RS.Open "SELECT "& CustY &" FROM Sys_DataRight WHERE EmpCode = '"& UseEmpCode &"' AND "&_
			"(SELECT ModuleRight FROM Employee WHERE EmpCode = '"& UseEmpCode &"') LIKE '%"& CurMCode &",%'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		EY = RS(CustY)
	end if
	RS.Close

	'当前用户所在组对当前模块的的客户权限
	RS.Open "SELECT R."& CustY &" FROM EmployeeRole A LEFT JOIN Sys_DataRight R ON A.GroupCode = R.GroupCode "&_
			"WHERE A.EmpCode = '"& UseEmpCode &"' AND (SELECT ModuleRight FROM EmployeeGroup WHERE GroupCode = A.GroupCode) LIKE '%"& CurMCode &",%'", G_DBConn, 0, 1, 1
	do while not RS.EOF
		GY = RS(CustY)
		if EY = "All" or GY = "All" then
			EY = "All"
		else
			if GY <> "" then
				EY = GY
			end if
		end if
		RS.MoveNext
	loop
	RS.Close

	if Sort = "Sel" then		'判断能否查看时,返回用于SQL(WHERE后)的限制条件(True | False)
		if EY <> "" then
			if EY = "All" then
				SeeCustEx = "1 = 1"				'全部客户权利
			else												'本人客户权利
				SeeCustEx = CustCode &" IN (SELECT CustCode FROM Customer WHERE EmpScope LIKE '%"& UseEmpCode &"%' UNION SELECT 0)"
			end if
		else
			SeeCustEx = "1 <> 1"					'无客户权利
		end if
	else						'判断能否修改、删除时,返回(用在SELECT后的SQL)最终用于脚本的限制条件("True" | "False")
		if EY <> "" then
			if EY = "All" then
				SeeCustEx = "CAST(1 AS Bit)"
			else
				SeeCustEx = "CASE WHEN "& CustCode &" IN (SELECT CustCode FROM Customer WHERE EmpScope LIKE '%"& UseEmpCode &"%' UNION SELECT 0) "&_
						   "THEN CAST(1 AS Bit) ELSE CAST(0 AS Bit) END"
			end if
		else
			SeeCustEx = "CASE WHEN "& CustCode &" = '' THEN CAST(1 AS Bit) ELSE CAST(0 AS Bit) END"
		end if
	end if

	set RS = nothing
end function


'================================================================================================================================
'	直接向SQL语句中传值时,把单引号替换成两个单引号
'================================================================================================================================
function Valid(Content)
	if IsNull(Content) then
		Valid = ""
	else
		Valid = Replace(Content, "'", "''", 1, -1, 1)
	end if
end function


'================================================================================================================================
'	生成判断保存,删除操作能否执行的函数体	职员代码(EmpCode):数据的EmpCode，
'	IE中客户Id的选择框或客户名框名称(CustId):"<select name='CustId'...|<input type="text" name='CustId'"， 
'	操作类型(Sort):"Edi|Del|Chk"
'================================================================================================================================
function CanOper(EmpCode, CustId, CustState, Sort)
	set RS = Server.CreateObject("ADODB.Recordset")
	if EmpCode = "" and (CustId = "" or CustState = "") then
		stop		'参数无效 EmpCode, 或 CustId,CustState 必须一个有效	
	end if
	RS.Open "SELECT EmpNameChs FROM Employee WHERE EmpCode = '"& EmpCode &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		EmpName = RS("EmpNameChs")
	end if
	RS.Close
	
	select case Sort
		case "Edi"
			NoEmpMsg = "修改"
		case "Del"
			NoEmpMsg = "删除"
		case "Chk"
			NoEmpMsg = "审核"
	end select

	'在客户端脚本中得到操作名称的字符串
	OperName = "window.event.srcElement.value"
	
	'判断对用户是否有操作的权限,返回用于SCRIPT中的 VBS
	if EmpCode <> "" then
		if not SeeEmp(EmpCode, Sort) then
			CanOper = "MsgBox ""您没有"& NoEmpMsg &" """""& EmpName &""""" 单据的权限！""& vbCrLf & vbCrLf & "&_
					  " "& OperName &" &""操作被取消。"", vbOKOnly & vbCritical, ""提示"": "&_
					  "window.event.returnValue = false "& vbCrLf &" exit sub"& vbCrLf
			exit function
		else 
'			CanOper = "window.event.returnValue = true "	2007.11.23
		end if
	end if
	
	'判断对客户是否有操作的权限,返回用于SCRIPT中的 VBS, CustId为客户选择框或客户名文本框
	if CustId <> "" and CustState <> "" then
		CanOper = "if "& CustState &".value = ""False"" then "& vbCrLf &_
				  "	if "& CustId &".type = ""text"" then "& vbCrLf &_
				  "		CustName = "& CustId &".value "& vbCrLf &_
				  "	else "& vbCrLf &_
				  "		CustName = "& CustId &".options("& CustId &".selectedIndex).text "& vbCrLf &_
				  "	end if "& vbCrLf &_
				  "	MsgBox ""您没有使用客户 """"""& CustName &"""""" 的权限！""& vbCrLf & vbCrLf & "&_
				  " "& OperName &" &""操作被取消。"", vbOKOnly & vbCritical, ""提示"""& vbCrLf &_
				  "	window.event.returnValue = false "& vbCrLf &" exit sub"& vbCrLf &_				  
				  "end if"
'				  "Else "& vbCrLf &_			2007.11.23
'				  "window.event.returnValue = True"& vbCrLf &_

	end if
	
	set RS = nothing
end function


'================================================================================================================================
'	设置合同的是否占用风险标记表(Contract:IsCredit)		出口合同Id(Id)
'================================================================================================================================
function SetIsCredit(Id)
	set RS = Server.CreateObject("ADODB.Recordset")		'当出口合同审核通过时调用此函数,根据当前设置风险条件,设置其风险占用标记
	'判断合同是否符合需受信的付款条件
	'====客户批准授信额度不为空，0也算授信；并且合同付款条件需要授信
	RS.Open "SELECT A.CustCode, A.PayTerms, A.PayTermsDays ,B.MaxAmt,P.IsExam "&_
			"FROM Contract A "&_
			"LEFT JOIN Payment P ON A.PayTerms = P.PaymentCode "&_
			"Left join Customer B On A.CustCode=B.CustCode "&_
			"WHERE A.ContractId = '"& Id &"' AND  "&_
			" B.MaxAmt Is Not  Null And P.IsExam=1 ", G_DBConn, 0, 1, 1
	if not RS.EOF then
		IsCredit = 1
	else
		IsCredit = 0
	end if
	RS.Close
	
	G_DBConn.Execute "UPDATE Contract SET IsCredit = "& IsCredit &" WHERE ContractId = '"& Id &"'"
	set RS = nothing
end function


'================================================================================================================================
'	取得请求的页面名称
'================================================================================================================================
function GetReqPageName()
	reqScriptName = Request.ServerVariables("SCRIPT_NAME")
	i = 1
	j = 0
	do
		i = j + 1
		j = Instr(i, reqScriptName, "/", 1)
	loop while j > 0
	GetReqPageName = Mid(reqScriptName, i)
end function

'================================================================================================================================
'	取得在表中新添加记录的ID
'================================================================================================================================
function GetNewId()
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT @@IDENTITY NewBhdId", G_DBConn, 0, 1, 1
	if not RS.EOF then
		if IsNull(RS("NewBhdId")) then
			stop
		else
			GetNewId = RS("NewBhdId")
		end if
	else
		stop
	end if
	RS.Close 
	SET RS = NOTHING
end function

'================================================================================================================================
'	记录并取得页面当前选定的值,在跳转页面后仍然可以得到用户曾指定的值	表单元素名称:Name
'================================================================================================================================
function CurSelValue(Name)
	CurPageName = GetReqPageName()
	ReqName = Request(Name)
	
	if IsEmpty(ReqName) then
		CurSelValue = Request.Cookies(CurPageName)(Name)
	else
		Response.Cookies(CurPageName)(Name) = ReqName
		CurSelValue = ReqName
	end if
end function

'================================================================================================================================
'	判断是否为表中不重复的有效的值.	(Value:待检测值,Table:表名,Field:字段名,CurId:当前记录ID)
'================================================================================================================================
function ValidateHad(Value, CurId, Table, Field, CurIdField)
	set RS = Server.CreateObject("ADODB.Recordset")

	RS.Open "SELECT * FROM "& Table &" WHERE "& Field &" = '"& Valid(Value) &"' AND "& CurIdField &" <> '"& CurId &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		ValidateHad = ""										'Code值无效时返回空串
	else
		ValidateHad = Value
	end if
	RS.Close
	
	set RS = nothing
end function

'定义必填标记
	EmphasisTag = "<font color=""red"" size=""2"">*</font>"
%>