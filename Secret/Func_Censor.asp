<script language="VBS" runat="Server">
'================================================================================================================================
'	显示每个人员审核结果对应的文字说明
'================================================================================================================================
function ChkState(State)
	if IsNull(State) then
		ChkState = "未审"
	else
		select case State
			case False		ChkState = "否决"
			case True		ChkState = "通过"
		end select
	end if
end function

'================================================================================================================================
'	显示审核状态对应的文字说明
'================================================================================================================================
function ChkResult(Status)
	if IsNumeric(Status) then
		select case Status
			case 0		ChkResult = "尚未提交"					'Self	ProcTable	不需要使用审核规则
			case 1		ChkResult = "已经提交"					'Censor	RuleTable	需要使用审核规则
			case 2		ChkResult = "申请退回"					'Censor	RuleTable
			case 3		ChkResult = "审核未通过"				'Self	ProcTable
			case 4		ChkResult = "已经退回"					'Self	ProcTable
			case 5		ChkResult = "无需审核"					'Self	ProcTable
			case 6		ChkResult = "审核通过"					'Self	ProcTable	AllSee
		end select
	else
		ChkResult = Status
	end if
end function

'===============================================================================================================================
'	输出审核所需的操作按钮	审核对象对象的审核状态(Status),	审核规则Id(RuleId)
'===============================================================================================================================
function ChkButton(Sort,Status, IsCensor, RuleId, EmpCode)
	if Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O" then		'位于审核页面(N|O)
		if SeeEmp(EmpCode, "Chk") then						'用户可以审核当前的单据
			sClk = "Rule"& RuleId &".name=""RuleId"": Rule"& RuleId &".value="""& RuleId &""""
		else												'用户不可以审核当前的单据
			sClk = CanOper(EmpCode, "", "", "Chk")
		end if

		select case Status					'因为允许一个人审核多个规则,当审核(同意,不同意)时,同时提交当前审核的规则的Id (RuleId)
			case 1
				ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""同意"" language=""VBS"" onclick='"& sClk &"'> "&_
							"<input type=""submit"" name="""& Sort &"Submit"" value=""不同意"" language=""VBS"" onclick='"& sClk &"'>"
			case 2
				ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""退回"" language=""VBS"" onclick='"& sClk &"'>"
		end select
		ChkButton = ChkButton &"<input type=""hidden"" id=""Rule"& RuleId &""" name="""" value="""">"
	else									'用户业务状态
		'用户单击提交或退回按钮时，判断用户是否对当前单据由修改权限，依赖每个页面的 EdiOper()函数 2007.11.23
		select case Status
			case 0, 3, 4
				if DisplayChkButton(Sort, Id, 1) then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""提交"" onclick=""EdiOper()"">"
				end if
			case 5							'单据无需审核时,当此类审核对象当前设为需审时,显示提交按钮,提供转换到审核循环的功能
				if IsCensor then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""提交"" onclick=""EdiOper()"">"
				end if
			case 1, 6
				if Status = 6 then
					sClk = "Rt = MsgBox(""该单据已审核通过，确实要退回吗？"", vbOKCancel + vbQuestion + vbDefaultButton2, ""确认""): if Rt = vbCancel then window.event.returnValue = false end if: "
					if Sort = "Invoice" or Sort = "NMActualBudget" then 		'凯路特殊要求：财务经理审核通过后便没有申请退回功能，需要总经理撤销成提交状态。
						exit function
					end if
				end if
				if DisplayChkButton(Sort, Id, 2) then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""申请退回"" language=""VBS"" onclick='"& sClk &" EdiOper'>"
				end if
		end select
	end if
end function

'================================================================================================================================
'	审核操作处理	审核对象的类别(Sort): 表CensorObject中的ObjectCode, 对象的ID(Id)
'================================================================================================================================
function CensorOper(Sort, Id)
	Submit = Request(Sort &"Submit")
	RuleId = Request("RuleId")
	InureMsg=Request("InureMsg")
	if Submit = "提交" then	
		set RS = Server.CreateObject("ADODB.Recordset")	
		if Sort="ExpContract" then 			'外销合同提交时判断是否符合提交条件！
			If AllowSubmit(Id) =0 then 
				'提交时需要指定表(CensorRules)中的审核规则
				if Request(Sort &"RuleId").Count <> 0 then
					G_DBConn.Execute "DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
					for each r in Request(Sort &"RuleId")
						G_DBConn.Execute "INSERT INTO CensorProcess(ObjectCode, ObjectId, RuleId, SubmitDate, ChkResult) "&_
									   "VALUES('"& Sort &"', '"& Id &"', '"& r &"', '"& Date &"', 1)"
					next
				else
					'当未指定规则时,判断审核对象是否仍需要审核,如不需审核,直接置为无需审核
					G_DBConn.Execute "IF (SELECT COUNT(*) FROM CensorRules WHERE ObjectCode = '"& Sort &"') = 0 "&_
								   "BEGIN DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'; "&_
								   "INSERT INTO CensorProcess(ObjectCode, ObjectId, ChkResult) VALUES('"& Sort &"', '"& Id &"', 5) END"
				end if
			else
				response.end
			end if 
		else
			'提交时需要指定表(CensorRules)中的审核规则
			'===========合同有效性时，保存生效说明
			if Request(Sort &"RuleId").Count <> 0 then
				G_DBConn.Execute "DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
				for each r in Request(Sort &"RuleId")
					G_DBConn.Execute "INSERT INTO CensorProcess(ObjectCode, ObjectId, RuleId, SubmitDate, ChkResult,InureMsg) "&_
								   "VALUES('"& Sort &"', '"& Id &"', '"& r &"', '"& Date &"', 1,'"&InureMsg&"')"
				next
			else
				'当未指定规则时,判断审核对象是否仍需要审核,如不需审核,直接置为无需审核
				G_DBConn.Execute "IF (SELECT COUNT(*) FROM CensorRules WHERE ObjectCode = '"& Sort &"') = 0 "&_
							   "BEGIN DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'; "&_
							   "INSERT INTO CensorProcess(ObjectCode, ObjectId, ChkResult) VALUES('"& Sort &"', '"& Id &"', 5) END"
			end if
		end if
		
		if Sort ="PayApplication" then 
			bPrePay = true
			RS.open "SELECT ISNULL(MIN(BhdId), 0) BhdId FROM AccountFeeItem A  WHERE AccId='"& Id &"'", G_DBConn, 0, 1, 1
			if RS("BhdId") <> 0 then
				bPrePay = false
			end if
			RS.close

			if bPrePay then 	'在提交审核过程中，判断是否录入发票号。如果没有发票号，属于预付类型财务需要冲账；如果有发票号，财务不需要冲账(3)
				Response.Write "<body onclick=""location.replace('PaymentReportEdit.asp?AccId="& Id &"')""><center><font color=red>缺少发票号，当前为付款，支付后需要确录入正确的发票号！</font></center>"
				Response.End
				'ErrMsg("缺少发票号，当前为付款，支付后需要确录入正确的发票号！")
			else
				G_DBConn.Execute "update AccountFee set IsPrePay=3 where AccId='"& Id &"' "
			end if
		end if 
		
		set RS = Nothing
	end if
	
	if Submit = "申请退回" then
		'首先删除表(CensorProcess)中已经在表(CensorRules)中不存在的审核规则,
		'然后根据审核对象当前是否已经被审核过,如果没有直接置为尚未提交,否则置为申请退回,
		'最后判断审核对象在审核表(CensorProcess)是否还有审核记录,如没有根据对象在表(CensorObject)中是否需要审核的状态,
		'写入尚未提交(有规则),无需审核(无规则)
		G_DBConn.Execute "DELETE FROM CensorProcess FROM CensorProcess A WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' "&_
					   "AND NOT EXISTS(SELECT * FROM CensorRules WHERE ObjectCode = '"& Sort &"' AND RuleId = A.RuleId); "&_
					   "IF NOT EXISTS(SELECT * FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND ChkState IS NOT NULL) "&_
					   " begin "&_
					   "	UPDATE CensorProcess SET ChkResult = 0 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' "&_
					   " 	Update AccountFee set IsPrePay=1 where AccId='"& Id &"' and '"& Sort &"'='PayApplication' " &_
					   " end "&_
					   "ELSE "&_
					   "	UPDATE CensorProcess SET ChkResult = 2 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'; "&_
					   "IF (SELECT COUNT(*) FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"') = 0 "&_
					   "	INSERT INTO CensorProcess(ObjectCode, ObjectId, ChkResult) SELECT '"& Sort &"', '"& Id &"', "&_
					   "	CASE WHEN (SELECT IsCensor FROM CensorObject WHERE ObjectCode = '"& Sort &"') = 0 THEN 5 ELSE 0 END"
	end if
	
	if Submit = "同意" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkState = 1, ChkEmpCode = '"& UserId &"', ChkDate = '"& Date &"', "&_
					   "ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = "& RuleId
	end if
	
	if Submit = "不同意" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkState = 0, ChkEmpCode = '"& UserId &"', ChkDate = '"& Date &"', "&_
					   "ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = "& RuleId
		if Sort ="PayApplication" then		''否决和退回时，连同审核状态一同退回。
			G_DBConn.Execute "update AccountFee set IsPrePay=1 where AccId='"& Id &"' "
		end if
	end if
	
	if Submit = "退回" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = '"& RuleId &"'; "&_
					   "UPDATE CensorProcess SET ChkResult = 4 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
					   
		if Sort ="PayApplication" then	''否决和退回时，连同审核状态一同退回。
			G_DBConn.Execute "update AccountFee set IsPrePay=1 where AccId='"& Id &"' "
		end if
	end if
end function

'===============================================================================================================================
'	对于没有审核记录的审核对象,根据当前审核设置向表(CensorProcess)写入审核记录
'===============================================================================================================================
function SetCensorProcess(ObjectCode, Id, IsCensor)
	set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT * FROM CensorProcess WHERE ObjectCode = '"& ObjectCode &"' AND ObjectId = '"& Id &"'", G_DBConn, 1, 3, 1
	if RS.EOF then
		RS.AddNew
		RS("ObjectCode") = ObjectCode
		RS("ObjectId") = Id
		if IsCensor then
			RS("ChkResult") = 0
		else
			RS("ChkResult") = 5
		end if
		RS.Update
	end if
	RS.Close
	
	set RS = nothing
end function

'===============================================================================================================================
'	生成提交给的审核人员名称	部门名(DeptName),	身份名(GroupName),	职员名(EmpName)
'===============================================================================================================================
function ChkName(DeptName, GroupName, EmpName)
	if EmpName <> "" then						'如果有审核人,显示此名称,否则显示部门名与身份名的组合
		ChkName = EmpName
	else
		'if DeptName <> "" and GroupName <> "" then
		'	sSpace = " "
		'end if
		ChkName = DeptName & sSpace & GroupName
	end if
end function

'===============================================================================================================================
'	取得审核设置处(CensorRules)的审核对象的某一级审核设置	审核对象(ObjectCode),	审核级别(Level): 1 | 2 | 3
'===============================================================================================================================
function CensorLevel(ObjectCode, EmpCode, Level)
	set RS = Server.CreateObject("ADODB.Recordset")
	set CensorLevel = Server.CreateObject("Scripting.Dictionary")		'返回 Dictionary 对象

 
	RS.Open "SELECT A.RuleId, D.DeptName, G.GroupName, E.EmpNameChs, A.EmpCode, "&_
			"(SELECT CensorModuleCode FROM CensorObject WHERE ObjectCode = '"& ObjectCode &"') CensorModuleCode "&_
			"FROM CensorRules A LEFT JOIN Dept D ON A.DeptCode = D.DeptCode "&_
			"LEFT JOIN EmployeeGroup G ON A.GroupCode = G.GroupCode "&_
			"LEFT JOIN Employee E ON A.EmpCode = E.EmpCode "&_
			"WHERE A.ObjectCode = '"& ObjectCode &"' AND A.CensorLevel = '"& Level &"'", G_DBConn, 0, 1, 1
			
	do while not RS.EOF 
		ChkEmpCode = RS("EmpCode")
		CensorModuleCode = RS("CensorModuleCode")
		CanUsed=false	
		if EmpCode <> "" and ChkEmpCode <> "" then		
			'判断在对象的审核模块CensorModuleCode处，指定的审核人ChkEmpCode是否对EmpCode的单据有审核权限
			if SeeEmpEx(EmpCode, "Chk", ChkEmpCode, CensorModuleCode) then
				CanUsed = true
			end if
		else
			CanUsed = true
		end if
		
		if CanUsed then
			CensorLevel.Item(RS("RuleId").value) = ChkName(RS("DeptName"), RS("GroupName"), RS("EmpNameChs"))		'key = RuleId			
		end if
		RS.MoveNext
	Loop
	RS.Close 
	
	set RS = nothing
end function
'===============================================================================================================================
'	取得审核设置处(CensorRules)的审核对象的审核级数	审核对象(ObjectCode)
'===============================================================================================================================
function CensorLevelCount(ObjectCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT COUNT(*) FROM CensorRules WHERE ObjectCode = '"& ObjectCode &"' GROUP BY CensorLevel", G_DBConn, 1, 1, 1
	if not RS.EOF then
		CensorLevelCount = RS.RecordCount
	end if
	RS.Close
	set RS = nothing
end function

'===============================================================================================================================
'	判断当前用户是否是对象的审核规则中的可审人员
'===============================================================================================================================
function CanChk(RuleId)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT A.* FROM CensorRules A WHERE A.RuleId = '"& RuleId &"' AND "&_
			"(A.EmpCode = '' OR A.EmpCode <> '' AND A.EmpCode = '"& UserId &"') AND "&_
			"(A.DeptCode = '' OR A.DeptCode <> '' AND A.DeptCode = '"& UDept &"') AND "&_
			"(A.GroupCode = '' OR A.GroupCode <> '' AND EXISTS(SELECT * FROM EmployeeRole WHERE EmpCode = '"& UserId &"' AND GroupCode = A.GroupCode) )", G_DBConn, 0, 1, 1
	if not RS.EOF then
		CanChk = True
	else
		CanChk = False
	end if
	RS.Close
	
	set RS = nothing
end function

'===============================================================================================================================
'	设置单据的最终审核结果,取得当前(CensorProcess)的审核级别	审核对象的类别(Sort), 对象的ID(Id)
'===============================================================================================================================
function CurCensorLevel(Sort, Id)
	set RS = Server.CreateObject("ADODB.Recordset")
	set LevelDenyState = Server.CreateObject("Scripting.Dictionary")
	set LevelPassState = Server.CreateObject("Scripting.Dictionary")
	Submit = Request(Sort &"Submit")
	
	'执行用户审核操作
	CensorOper Sort, Id

	'从表 CensorProcess 中取得单据提交的审核设置
	RS.Open "SELECT A.ChkState, ISNULL(R.CensorLevel, '') CensorLevel, ISNULL(R.InnerLevelCode, '') InnerLevelCode, "&_
			"ISNULL(R.DenyTerm, '') DenyTerm, ISNULL(R.PassTerm, '') PassTerm, "&_
			"(SELECT NeedAllCensor FROM CensorObject WHERE ObjectCode = '"& Sort &"') NeedAllCensor "&_
			"FROM CensorProcess A LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
			"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' ORDER BY R.CensorLevel ASC", G_DBConn, 0, 1, 1	
	do while not RS.EOF
		NeedAllCensor = RS("NeedAllCensor")				'是否需要提交全部审核级别审核(true)

		if CurLevel <> RS("CensorLevel") then
			DenyTerm = RS("DenyTerm")
			DenyTerm = Replace(DenyTerm, "AND", "and", 1, -1, 1)		'替换条件中的操作符为小写，以使替换为审核结果时不会替换(AND)的A
			DenyTerm = Replace(DenyTerm, "OR", "or", 1, -1, 1)
			
			PassTerm = RS("PassTerm")
			PassTerm = Replace(PassTerm, "AND", "and", 1, -1, 1)
			PassTerm = Replace(PassTerm, "OR", "or", 1, -1, 1)

			'替换审核通过或否决条件中的参数，如合同金额
			PassTerm = ReplaceParam(Sort, Id, PassTerm)
			DenyTerm = ReplaceParam(Sort, Id, DenyTerm)
		end if
		CurLevel = RS("CensorLevel")									'< 最终等于提交到的最高的审核级别 >
		
		if IsNull(RS("ChkState")) then
			State = "null"
		else
			State = LCase(CStr(RS("ChkState")))
		end if


		DenyTerm = Replace(DenyTerm, RS("InnerLevelCode"), State &" = false", 1, -1, 0)
		LevelDenyState.Item(CurLevel) = DenyTerm									'key = CensorLevel

		PassTerm = Replace(PassTerm, RS("InnerLevelCode"), State &" = true", 1, -1, 0)
		LevelPassState.Item(CurLevel) = PassTerm									'key = CensorLevel
		RS.MoveNext
	loop
	RS.Close


	'替换审核条件中当前未提交到的审核人的代码(A|B|...)为空(null = 0|1)
	for i = 1 to Len(DenyTerm)
		Ch = Mid(DenyTerm, i, 1)
		if Ch > "A" and Ch < "Z" then
			DenyTerm = Replace(DenyTerm, Ch, "null = false", 1, 1, 0)
		end if
	next
	for i = 1 to Len(PassTerm)
		Ch = Mid(PassTerm, i, 1)
		if Ch > "A" and Ch < "Z" then
			PassTerm = Replace(PassTerm, Ch, "null = true", 1, 1, 0)
		end if
	next
	
	'得到当前的审核级别,对于在审核人处的单据(ChkResult = 1|2)返回的审核结果正确,
	'其它状态返回结果可能不正确(当审核规则已不在表(CensorRules)处时！！！
	for each level in LevelDenyState
		CurCensorLevel = level
		if LevelDenyState.Item(Level) <> "" then			'如果表(CensorProcess)当前的审核规则仍在表(CensorRules)中
			DenyState = Eval(LevelDenyState.Item(level))					'某一级的否决条件结果(True|False)
		end if
		if LevelPassState.Item(Level) <> "" then
			PassState = Eval(LevelPassState.Item(level))					'某一级的通过条件结果(True|False)
		end if

		'if DenyState or IsNull(DenyState) and not PassState or IsNull(PassState) then	'如果当前的审核级别未通过或未审，审核停留在此级
		if not PassState or IsNull(PassState) then
			exit for
		end if
	next

	' ==================================================================================<<< 设置审核结果 >>>======================
	'在执行完审核操作,取得当前审核级别后,判断单据的最终审核结果
	if Submit = "同意" or Submit = "不同意" or Submit = "设置" then		'当同意,不同意,在审核设置处更改审核条件时(CensorRulesEdit.asp Line167),需要重新计算单据最终审核结果
		TotalCenLevel = CensorLevelCount(Sort)			'当前审核对象设置的总的审核级数
		SubmitCenLevel = LevelDenyState.Count 
		if DenyState then												'任何一级未通过时,单据已遭最终否决
			G_DBConn.Execute "UPDATE CensorProcess SET ChkResult = 3 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
		else
			'NeedAllCensor=true 时，当审核对象的所有审核级别都通过后,单据最终通过
			'NeedAllCensor=false 时，当已提交的最后一级通过时,单据最终通过
			if NeedAllCensor and PassState and CurCensorLevel = TotalCenLevel and SubmitCenLevel = TotalCenLevel _
				or not NeedAllCensor and PassState and CurCensorLevel = CurLevel then
				G_DBConn.Execute "UPDATE CensorProcess SET ChkResult = 6 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"

				'如果出口合同审核通过后,根据当前设置风险条件,设置其风险占用标记！！！！！！！！！！
				if Sort = "ExpContract" then
					SetIsCredit Id
				end if
				
			end if
		end if
		
		'判断是否审核结束
		if DenyState or NeedAllCensor and PassState and CurCensorLevel = TotalCenLevel and SubmitCenLevel = TotalCenLevel _
			or not NeedAllCensor and PassState and CurCensorLevel = CurLevel then
			NeedSetChkName = true
		end if
	end if

	if Submit = "退回" then
		'退回时根据对象是否需要全部审核人执行退回操作，判断当前是否可以设为已退回状态
		RS.Open "SELECT CASE WHEN (SELECT NeedAllWithdrawal FROM CensorObject WHERE ObjectCode = '"& Sort &"') = 1 "&_
				"				  AND NOT EXISTS(SELECT A.* FROM CensorProcess A LEFT JOIN CensorRules B ON A.RuleId = B.RuleId "&_
				"				  WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' AND B.CensorLevel <= '"& CurCensorLevel &"' "&_
				"				  AND A.Withdrawal = 0) "&_
				"				  OR "&_
				"				  (SELECT NeedAllWithdrawal FROM CensorObject WHERE ObjectCode = '"& Sort &"') = 0 "&_
				"THEN 1 ELSE 0 END NeedSetChkName ", G_DBConn, 0, 1, 1
		if not RS.EOF then
			if RS("NeedSetChkName") = "1" then
				G_DBConn.Execute "UPDATE CensorProcess SET ChkResult = 4 FROM CensorProcess A WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"'"
				NeedSetChkName = true
			end if
		end if
		RS.Close
	end if

	'当审核结束时(最终审核通过，被否决，),把提交的审核人及审核人的名称写入(CensorProcess)的 ChkName中,
	'以后显示的审核信息全部来自于(CensorProcess),在(CensorRules)修改审核规则后仍能正确显示当时的审核情况
	if NeedSetChkName then
		RS.Open "SELECT A.ProcessId, ChkName, D.DeptName, G.GroupName, E.EmpNameChs "&_
				"FROM CensorProcess A LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
				"LEFT JOIN Dept D ON R.DeptCode = D.DeptCode "&_
				"LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
				"LEFT JOIN Employee E ON R.EmpCode = E.EmpCode "&_
				"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"'", G_DBConn, 0, 1, 1
		do while not RS.EOF
			Name = ChkName(RS("DeptName"), RS("GroupName"), RS("EmpNameChs"))
			ChkEmpName = RS("EmpNameChs")

			if ChkEmpName <> Name and ChkEmpName <> "" then
				Name = Name &"("& ChkEmpName &")"
			end if
			
			'ODBC
			G_DBConn.Execute "UPDATE CensorProcess SET ChkName = '"& Name &"' WHERE ProcessId = '"& RS("ProcessId") &"'"
			'RS("ChkName") = Name	
			'RS.Update
			RS.MoveNext
		loop
		RS.Close
	end if

	set RS = nothing
	set LevelDenyState = nothing
	set LevelPassState = nothing
end function

'===============================================================================================================================
'	显示在业务处的提交所需的侯选审核人和提交按钮	审核对象(Sort),	对象ID(Id),	 对象的职员代码(EmpCode), 
'	对象的审核结果(Result):0 | 1,	当前的审核级别(CurLevel)
'===============================================================================================================================
function CensorState(Sort, Id, EmpCode, IsCensor, Result, CurLevel)
	set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT A.InureMsg FROM CensorProcess A "&_
			" LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
			" WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' AND R.CensorLevel = '"& CurLevel &"'", G_DBConn, 0, 1, 1
	If Not RS.Eof then
		InureMsg=RS("InureMsg")
	End If
	RS.Close
	if Left(ModuleCode, 1) <> "N" and Left(ModuleCode, 1) <> "O" then
		select case Result
			case 0, 3, 4, 5									'未提交,未通过,退回时,从表(CensorRules)显示所有级别的审核候选人
				LevelCount = CensorLevelCount(Sort)
				Response.Write "<tr><td width=""85%"" align=""left"" colspan=""3""><span style=""width:85%"">"
				if not(Result = 5 and not IsCensor) then		'单据无需审核时,当此类审核对象当前设为需审时,显示审核候选人,提供转换到审核循环的功能
					for i = 1 to LevelCount
						'取得某一级的审核设置
						set Content = CensorLevel(Sort, EmpCode, i)
						if i <> 1 then
							Response.Write "<br>"
						end if

						Response.Write "<b>"& i &".</b>"
						for each RuleId in Content
							'根据CensorRules:Fixed，判断当前审核规则是否可由用户修改
							sDisabled = ""
							RS.Open "SELECT Fixed FROM CensorRules WHERE RuleId = '"& RuleId &"'", G_DBConn, 0, 1, 1
							if not RS.EOF then
								if RS("Fixed") then
									sDisabled = "checked onclick=""window.event.returnValue=false"""
								end if
							end if
							RS.Close 

							Response.Write "<input type=""checkbox"" id=""Rule"& RuleId &""" name="""& Sort &"RuleId"" value="""& RuleId &""" "& sDisabled &">"&_
										   "<label for=""Rule"& RuleId &""" style=""width:100px; margin-bottom:-3px"">"& Content.item(RuleId) &"</label>"
						next
					next
				end if
				'===============增加生效说明
				Response.Write "</span><Br><span style=""width:85%;"">"
				If Sort="ValidateExpCon" Then '有效性审核，添加生效说明		
				Response.Write "生效说明:<input type=""text"" name=""InureMsg"" value="""" class=""input"" maxlength=""400"" style=""width:80%"">"				
				End If
				Response.Write "&nbsp;</span><span style=""width:15%; word-wrap:normal""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></span></td>"&_
							   "<td width=""15%"" valign=""bottom"">"& ChkButton(Sort,Result, IsCensor, 0, "") &"</td></tr>"
			case 1										'已提交时,从表(CensorProcess)显示等待审核的某一级选定的审核人
				set RS = Server.CreateObject("ADODB.Recordset")
			
				Response.Write "<tr><td align=""left"">"
				RS.Open "SELECT A.RuleId,A.InureMsg, D.DeptName, G.GroupName, E.EmpNameChs "&_
						"FROM CensorProcess A LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
						"LEFT JOIN Dept D ON R.DeptCode = D.DeptCode "&_
						"LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
						"LEFT JOIN Employee E ON R.EmpCode = E.EmpCode "&_
						"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' AND R.CensorLevel = '"& CurLevel &"'", G_DBConn, 0, 1, 1
				If Sort="ValidateExpCon" Then '有效性审核，添加生效说明				
				Response.Write "生效说明：<font color=red>"&RS("InureMsg")&"</font>"				
				End If
				Response.Write("&nbsp;</td><td align=""right"" colspan=""2"">")
				do while not RS.EOF
					Name = ChkName(RS("DeptName"), RS("GroupName"), RS("EmpNameChs"))
					Response.Write "<input type=""checkbox"" name="""" value="""" checked disabled>"&_
								   "<label style=""width:100px; text-align:left; margin-bottom:-3px"">"& Name &"</label>"
					RS.MoveNext
				loop
				RS.Close
				Response.Write "</td><td><b id="""& Sort &"ChkResult"" style=""display:none;"">"& ChkResult(Result) &"</b>"&_
								ChkButton(Sort,Result, IsCensor, 0, "") &"</td></tr>"
				
				set Level = nothing
				set RS = nothing
			case 2, 6
				Response.Write "<tr><td align=""left"">"
				If Sort="ValidateExpCon" Then '有效性审核，添加生效说明		
				Response.Write "生效说明：<font color=red>"&InureMsg&"</font>"				
				End If
				Response.Write "&nbsp;</td><td align=""right"" valign=""bottom"" colspan=""2"" style=""padding-right:30px""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></td>"&_
							   "<td width=""15%"">"& ChkButton(Sort,Result, IsCensor, 0, "") &"</td></tr>"
		end select
	else
		Response.Write "<tr><td align=""left"">"
		If Sort="ValidateExpCon" Then '有效性审核，添加生效说明
		Response.Write "生效说明：<font color=red>"&InureMsg&"</font>"	
		End If		
		Response.Write "&nbsp;</td><td align=""right"" valign=""bottom"" colspan=""2"" style=""padding-right:30px""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></td>"&_
					   "<td width=""15%""></td></tr>"
	end if

	set RS = nothing
end function

'================================================================================================================================
'	显示审核信息	审核对象的类别(Sort): 表CensorObject中的ObjectCode, 对象的ID(Id), 单据的用户代码(EmpCode)
'================================================================================================================================
function CensorInfo(Sort, Id, EmpCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	set RSTP = Server.CreateObject("ADODB.Recordset")

	'判断当前的对象是否需要审核
	RS.Open "SELECT * FROM CensorObject WHERE ObjectCode = '"& Sort &"'", G_DBConn, 0, 1, 1
	if RS.EOF then
		stop							'对不需要审核的对象显示审核信息
		exit function
	else
		IsCensor = RS("IsCensor")		'对象是否需要审核
	end if	
	RS.Close 	
	Result = 0								'单据的审核结果
	CurLevel = CurCensorLevel(Sort, Id)		'当前审核停留在此级
	'可审核对象在表(CensorProcess)中必须至少有一条审核记录,如果没有通过以下函数加入
	SetCensorProcess Sort, Id, IsCensor
	Response.Write "<table class=""pagetable"">"

	'读取对象的全部审核信息,供显示用
	RSTP.Open "SELECT A.RuleId, A.SubmitDate, A.ChkName, A.ChkState, A.ChkResult, A.ChkDate, A.ChkMessage, E.EmpNameChs, "&_
			"A.RuleId,A.InureMsg, D.DeptName, G.GroupName, F.EmpNameChs EmpName "&_
			"FROM CensorProcess A LEFT JOIN Employee E ON A.ChkEmpCode = E.EmpCode "&_
			"LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
			"LEFT JOIN Dept D ON R.DeptCode = D.DeptCode "&_
			"LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
			"LEFT JOIN Employee F ON R.EmpCode = F.EmpCode "&_
			"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' AND "&_
			"(A.ChkResult IN (1, 2) AND R.CensorLevel <= '"& CurLevel &"' OR A.ChkResult IN(0, 3, 4, 5, 6)) "&_
			"ORDER BY R.CensorLevel ASC", G_DBConn, 0, 1, 1
	Result = RSTP("ChkResult")			'单据的审核结果
	if Result <> 0 and Result <> 5 then					'当单据不为无须审核时,显示详细审核情况
		do while not RSTP.EOF
			RuleId = RSTP("RuleId")				'
			ChkEmpName = RSTP("EmpNameChs")		'级内审核人员名称
			State = RSTP("ChkState")				'级内人员的审核结果
			
			'当提交,申请退回时审核人从(CensorRules)获取,其它情况从(CensorProcess)中ChkName
			select case Result
				case 1, 2
					Name = ChkName(RSTP("DeptName"), RSTP("GroupName"), RSTP("EmpName"))
				case else
					Name = RSTP("ChkName")
			end select

			'当提交的审核人为部门获身份时,加入实际审核的人员名
			ChkEmpName = RSTP("EmpNameChs")
			if ChkEmpName <> Name and ChkEmpName <> "" then
				Name = Name &"("& ChkEmpName &")"
			end if
			'显示详细审核信息
			Response.Write "<tr><td width=""40%"" nowrap>审核人: "& Name &"</td><td width=""25%"" nowrap>审核时间: "& RSTP("ChkDate") &"</td>"&_
						   "<td width=""20%"" nowrap>审核状态: "& ChkState(State) &"</td><td width=""15%"" nowrap>提交时间: "& RSTP("SubmitDate") &"</td></tr>"

			if (Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O") and (Result = "1" and IsNull(State) or Result = "2") and CanChk(RuleId) then
				'如果用户可以审核，显示供审核的录入框和审核操作按钮
				Response.Write "<tr><td colspan=""3"">意见: <input type=""text"" name=""ChkMessage"& RuleId &""" value="""& RSTP("ChkMessage") &""" class=""input"" maxlength=""400"" style=""width:80%""></br>"&_
			" <td align=""right"">"& ChkButton(Sort,Result, IsCensor, RSTP("RuleId"), EmpCode) &"</td></tr>"
			else																		'仅显示审核信息
				Response.Write "<tr><td colspan=""3"">意见: <font color=blue>"& RSTP("ChkMessage") &"</font></td><td></td></tr>"				
			end if

			RSTP.MoveNext
		loop
	end if
	RSTP.Close
	Stopsubmit Sort, Id
	'在业务处显示等待审核人及操作按钮,	审核处显示当前的审核结果
	CensorState Sort, Id, EmpCode,IsCensor, Result, CurLevel
	Response.Write "</table><hr>"	
	set RS = nothing
end function

'================================================================================================================================
'	显示审核状态查询选项
'================================================================================================================================
function ChkQuery()
	ChkQ = CurSelValue("ChkQuery")
	dim opt(6)
	if ChkQ <> "" then
		opt(ChkQ) = "selected"
	end if
	
	Response.Write "<select name=""ChkQuery""><option value="""">审核状态</option>"
	select case Left(ModuleCode, 1)
		case "N", "O"
			Response.Write "<option value=""1"" "& opt(1) &">"& ChkResult(1) &"</option>"&_
						   "<option value=""2"" "& opt(2) &">"& ChkResult(2) &"</option>"
			Response.Write "<option value=""6"" "& opt(6) &">"& ChkResult(6) &"</option>"
		case else
			Response.Write "<option value=""0"" "& opt(0) &">"& ChkResult(0) &"</option>"&_
						   "<option value=""1"" "& opt(1) &">"& ChkResult(1) &"</option>"&_
						   "<option value=""2"" "& opt(2) &">"& ChkResult(2) &"</option>"&_
						   "<option value=""3"" "& opt(3) &">"& ChkResult(3) &"</option>"&_
						   "<option value=""4"" "& opt(4) &">"& ChkResult(4) &"</option>"&_
						   "<option value=""5"" "& opt(5) &">"& ChkResult(5) &"</option>"&_
						   "<option value=""6"" "& opt(6) &">"& ChkResult(6) &"</option>"
	end select
	Response.Write "</select>"
end function

'================================================================================================================================
'	返回页面可以查看的数据的审核状态值(0|1|...)
'================================================================================================================================
function ChkValue()
	'首先使用审核状态查询指定的审核结果(当位于列表页时),编辑页没有审核状态查询选项.2006.1.5
	'ChkValue = CurSelValue("ChkQuery")
	
	if ChkValue = "" then
		if Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O" then			'位于审核模块
			ChkValue = "1, 2, 6"
		else
			ChkValue = "0, 1, 2, 3, 4, 5, 6"
		end if
	end if
end function

'================================================================================================================================
'	返回数据审核状态限制串	审核对象(ObjectCode), 对象数据表的ID列名(Id),	查询的审核结果(Result)	如用于WHERE后( WHERE "& ChkSql("'ExpContract'", "A.ContractId", 6) &"..." )
'================================================================================================================================
function ChkSql(ObjectCode, Id, Result)
	'当查询审核状态时
	if Request("ChkQuery") <> "" then
		Result = Request("ChkQuery")
	end if

	'加入审核状态的条件限制, 传入的ObjectCode必须包括单引号如:'ImpContract' ！！
	if Result <> "" then
		ChkSql = "ISNULL((SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectId = "& Id &" AND ObjectCode = "& ObjectCode &"), 0) "&_
				 "IN ("& Result &")"
	end if
end function

'================================================================================================================================
'	返回审核结果(0|1|...)	审核对象(ObjectCode): 表CensorObject中的ObjectCode, 对象的ID(Id)
'================================================================================================================================
function CensorResult(ObjectCode, Id)

	set RS = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectCode = '"& ObjectCode &"' AND ObjectId = '"& Id &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		CensorResult = RS("ChkResult")
	else
		stop
		CensorResult = 0
	end if
	RS.Close

	set RS = nothing
end function

'================================================================================================================================
'	返回审核结果或当前查看人员的审核结果(0|1|通过|否决|...)	审核对象(ObjectCode): 表CensorObject中的ObjectCode, 对象的ID(Id)
'================================================================================================================================
function CensorResult2_Old(ObjectCode, Id)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectCode = '"& ObjectCode &"' AND ObjectId = '"& Id &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		CensorResult2 = RS("ChkResult")
	else
		stop
		CensorResult2 = 0
	end if
	RS.Close
	
	'如果状态为已经提交,而且显示审核页面时,判断当前用户的审核状态
	if CensorResult2 = 1 and (Left(ModuleCode ,1) = "N" or Left(ModuleCode, 1) = "O") then
		'取得对象当前的审核级别
		ObjCurCensorLevel = CurCensorLevel(ObjectCode, Id)
		
		for l = 1 to ObjCurCensorLevel
			Needed = 0
			Passed = 0
			Denyed = 0
			NoChked = 0
			'在当前审核级别内查找当前用户需要审核的规则
			RS.Open "SELECT A.ChkState, A.RuleId FROM CensorProcess A LEFT JOIN CensorRules B ON A.RuleId = B.RuleId "&_
					"WHERE A.ObjectCode = '"& ObjectCode &"' AND A.ObjectId = '"& Id &"' "&_
					"AND B.CensorLevel = '"& l &"'", G_DBConn, 0, 1, 1
			do while not RS.EOF
				if CanChk(RS("RuleId")) then		'对于需要审核的规则,计算用户所有的审核结果
					Needed = Needed + 1
					if RS("ChkState") then
						Passed = Passed + 1			'通过的个数
					end if
					if not RS("ChkState") then
						Denyed = Denyed + 1			'否定的个数
					end if
					if IsNull(RS("ChkState")) then
						NoChked = NoChked + 1
					end if
				end if
				RS.MoveNext
			loop
			RS.Close
		
			'在当前审核级别内,当前用户须审的规则全部通过时,输出'通过',有一个不通过输出'否决',有一个未审时输出'未审'
			if Needed = Passed and Needed <> 0 then
				CensorResult2 = ChkState(True)
			end if
			if Denyed <> 0 then
				CensorResult2 = ChkState(False)
			end if
			if NoChked > 0 then
				CensorResult2 = ChkState(Null)
			end if
		next
	end if

	set RS = nothing
end function

'================================================================================================================================
'	判断操作按钮是否显示的函数	审核对象(ObjectCode),	审核对象Id(ID),	要显示的按钮类型(OptBtn): (保存"Save"|打印"Print")
'================================================================================================================================
function Visible(ObjectCode, Id, OptBtn)
	if ObjectCode <> "" and Id <> "" then
		'取得审核对象的审核状态
		Result = CensorResult(ObjectCode, Id)
	else
		Result = 0								'如果指定的审核对象或对象ID无效,认为尚未提交
	end if
	
	select case OptBtn
		case "Save"					'要显示保存按钮
			select case Result
				case 0, 3, 4, 5		Visible = true
				case else			Visible = false
			end select
		case "Print"				'要显示打印按钮
			select case Result
				case 5, 6			Visible = true
				case else			Visible = false
			end select
	end select
end function


'================================================================================================================================
'	取得当前用户是否属于某一身份(True | False)	身份代码(GroupCode)
'================================================================================================================================
function CurUserInGroup(GroupCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT * FROM EmployeeRole WHERE EmpCode = '"& UserId &"' AND GroupCode = '"& GroupCode &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		CurUserInGroup = True
	else
		CurUserInGroup = False
	end if
	RS.Close
	
	set RS = nothing
end function

'================================================================================================================================
'	取得审核对象某一级的审核人的姓名	审核对象(ObjectCode),	审核对象Id(ID),	审核级别Level
'================================================================================================================================
function GetCensorName(ObjectCode, Id, Level)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT CASE WHEN A.ChkName IS NOT NULL THEN A.ChkName ELSE "&_
			"	(SELECT EmpNameChs FROM Employee WHERE EmpCode = A.ChkEmpCode) END CensorName "&_
			"FROM CensorProcess A LEFT JOIN CensorRules B ON A.RuleId = B.RuleId "&_
			"WHERE A.ObjectCode = '"& ObjectCode &"' AND A.ObjectId = '"& Id &"' AND "&_
			"B.CensorLevel = '"& Level &"' AND ChkState = 1", G_DBConn, 0, 1, 1
	Comma = ""
	do while not RS.EOF
		if GetCensorName <> "" then
			Comma = ", "
		end if
		GetCensorName = GetCensorName & Comma & RS("CensorName") 
		RS.MoveNext
	loop
	RS.Close 
	
	set RS = nothing
end function

'================================================================================================================================
'	返回是否应显示审核操作按钮(是True 否False)	审核对象(Sort), 对象Id(Id), 按钮类型(BtnType):提交1 申请退回2
'================================================================================================================================
function DisplayChkButton(Sort, Id, BtnType)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	if BtnType = 1 then						'要显示提交按钮
		select case Sort
			case "Contract"								'合同评审表
				DisplayChkButton = true
			case "SaleContract"							'销售合同
				RS.Open "SELECT ContractId FROM Contract A LEFT JOIN CensorProcess CP ON A.ContractId = CP.ObjectId AND CP.ObjectCode = 'Contract' "&_
						"WHERE A.ContractId = '"& Id &"' AND CP.ChkResult = 6", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true				'合同对应的评审表审核通过，显示提交按钮
				else
					DisplayChkButton = false
				end if
				RS.Close
			case "PlanProduct"							'排产单
				RS.Open "SELECT C.ContractId FROM PlanProduct A LEFT JOIN Contract C ON A.ConId = C.ContractId "&_
						"LEFT JOIN CensorProcess CP ON C.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.PlanId = '"& Id &"' AND CP.ChkResult = 6", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true
				else
					DisplayChkButton = false
				end if
			case "Bhd"									'发货单
				RS.Open "SELECT C.ContractId FROM Bhd A LEFT JOIN Contract C ON A.ConId = C.ContractId "&_
						"LEFT JOIN CensorProcess CP ON C.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.BhdId = '"& Id &"' AND (CP.ChkResult = 6 OR A.ConId = 0)", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true
				else
					DisplayChkButton = false
				end if
			case else
				if Id <> "-2" then						'其它情况对已经存在记录的对象显示提交按钮
					DisplayChkButton = true
				end if
		end select
	else									'要显示申请退回按钮
		select case Sort
			case "Contract"								'合同评审表
				RS.Open "SELECT ContractId FROM Contract A LEFT JOIN CensorProcess CP ON A.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.ContractId = '"& Id &"' AND CP.ChkResult IN(1, 2, 6)", G_DBConn, 0, 1, 1 
				if not RS.EOF  then
					DisplayChkButton = false
				else
					DisplayChkButton = true
				end if
				RS.Close
			case "SaleContract"							'销售合同
				'判断此合同是否有正在审核的排产单 和正在审核的发货单,如有则不允许退回合同
				RS.Open "SELECT A.PlanId FROM PlanProduct A LEFT JOIN CensorProcess CP ON A.PlanId = CP.ObjectId AND CP.ObjectCode = 'PlanProduct' "&_
						"WHERE A.ConId = '"& Id &"' AND CP.ChkResult IN (1, 2, 6) "&_
						"UNION  "&_
						"SELECT A.BhdId FROM Bhd A LEFT JOIN CensorProcess CP ON A.BhdId = CP.ObjectId AND CP.ObjectCode = 'Bhd' "&_
						"WHERE A.ConId = '"& Id &"' AND CP.ChkResult IN(1, 2, 6)", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = false
				else
					DisplayChkButton = true
				end if
				RS.Close
			case "PlanProduct"							'排产单
				DisplayChkButton = true
			case "Bhd"									'发货单
				DisplayChkButton = true
			case else
				DisplayChkButton = true
		end select
	end if

	set RS = nothing
end function

'===========================================================================================
'	判断是否符合提交条件
'		if 付款条件需要加减授信 then 
'			if 客户为新客户 then 
'				如果累计合同金额 〉USD5000 ，不允许继续提交。提示给新客户分配授信额度，不能提交。
'			else 
'				如果剩余额度不足，或者授信有效期已过，不能提交。
'			end if 
'		else 
'			
'				
'===========================================================================================
Function AllowSubmit(ContractId)
		'判断合同是否允许提交
		AllowSubmit=0
		Set RSA = Server.CreateObject("ADODB.Recordset")
		RSA.open "select C.Checkresult,C.FinalCFRDate,C.MaxAmt,A.CustCode, B.ConAmt "&_
				 "FROM  Customer C Join Contract A on A.CustCode=C.Custcode "&_
				 "LEFT JOIN VContract B ON A.ContractId = B.ContractId "&_
				 "where A.ContractId = '"& ContractId  &"'", G_DBConn, 0, 1, 1
			if RSA.eof = false then 
				CheckResult=RSA("Checkresult")
				if CheckResult<>6 then 				'没有授信许可
					if CustUsedAmt(RSA("CustCode")) > 5000 then 
						response.Redirect("ErrInfo.asp?Info=1")
						AllowSubmit=1
					end if 
				else 
					if RSA("FinalCFRDate") < date() then 
						response.Redirect("ErrInfo.asp?Info=2")
						AllowSubmit=1
					end if 

					if RSA("MaxAmt") < CDBL(CustUsedCredit(RSA("CustCode")))+CDBL(RSA("ConAmt")) then 
						response.Redirect("ErrInfo.asp?Info=3")
						AllowSubmit=1
					end if 
				end if
			end if 
		RSA.close
		Set RSA = Nothing
end Function

'==================================================================================
'	计算新客户累计未收汇合同金额。
'==================================================================================
	Function CustUsedAmt (CustId)
		'取得美元汇率
		Set RSB = Server.CreateObject("ADODB.Recordset")
		RSB.Open "SELECT ExRate FROM ExRate WHERE Currency = 'USD'", G_DBConn, 0, 1, 1
		if not RSB.EOF then
			UsdExRate =CDBL(RSB("ExRate"))
		end if
		RSB.Close 
		
		CustUsedAmt=0
		RSB.Open "SELECT ISNULL((SELECT SUM(ExpPrice * Qty) FROM ContractItem WHERE ContractId = A.ContractId), 0) ProdAmt, "&_
				"ISNULL((SELECT SUM(AddInSign * AddInValue) FROM ContractAddIn WHERE ContractId = A.ContractId), 0) AddInAmt, "&_
				"ISNULL((SELECT SUM(RecAmt) FROM AccountRecAmt WHERE ContractNo = A.ContractNo), 0) RecAmt, "&_
				"A.ConCurr,A.ContractNo,F.ExRate "&_
				"FROM Contract A INNER JOIN Customer C ON A.CustCode = C.CustCode "&_
				"Join Exrate F on A.ConCurr=F.Currency "&_
				"WHERE C.CustCode = '"& CustId &"' And " & ChkSql("'ExpContract'", "A.ContractId", "2,5,6"), G_DBConn, 2, 3, 1

		do while not RSB.EOF
			ProdAmt = CDBL(RSB("ProdAmt"))
			RecConAmt = CDBL(RSB("RecAmt"))
			if RSB("ConCurr") <> "USD" then
				ProdAmt = ProdAmt * CDBL(RSB("ExRate")) / UsdExRate
				AddInAmt = AddInAmt * CDBL(RSB("ExRate")) / UsdExRate
				RecConAmt = RecConAmt * CDBL(RSB("ExRate")) / UsdExRate
			end if
	
			CustUsedAmt = CustUsedAmt + CDbl(ProdAmt) + CDbl(RSB("AddInAmt")) - CDbl(RecConAmt)
			RSB.MoveNext
		loop
		RSB.Close 
		RSB.open "select ISNULL(SUM(ExpPrice * Qty),0) as CurrAmt FROM ContractItem WHERE ContractId = '"&ContractId&"'",G_DBConn,2,3,1
			if RSB.eof = false then 
				CustUsedAmt = CustUsedAmt + CDBL(RSB("CurrAmt"))
			End if 
		RSB.close
		Set RSB = Nothing 
	end function

'========================================================================================
' 计算老客户已占用的授信额度
'========================================================================================
'function CustUsedCredit(CustId)
'	'取得美元汇率
'	Set RSC = Server.CreateObject("ADODB.Recordset")
'	RSC.Open "SELECT ExRate FROM ExRate WHERE Currency = 'USD'", G_DBConn, 0, 1, 1
'	if not RSC.EOF then
'		UsdExRate = CDBL(RSC("ExRate"))
'	end if
'	RSC.Close 
'	'计算以美圆计的总的已用的授信额度	认为只有已审核通过的合同才占用授信额度(GMChkResult=6, 2)
'	CustUsedCredit=0
'	RSC.Open "SELECT ISNULL((SELECT SUM(ExpPrice * Qty) FROM ContractItem WHERE ContractId = A.ContractId), 0) ProdAmt, "&_
'			"ISNULL((SELECT SUM(AddInSign * AddInValue) FROM ContractAddIn WHERE ContractId = A.ContractId), 0) AddInAmt, "&_
'			"ISNULL((SELECT SUM(RecAmt) FROM AccountRecAmt WHERE ContractNo = A.ContractNo), 0) RecAmt, "&_
'			"A.ConCurr, A.ExRate, A.ContractNo "&_
'			"FROM Contract A INNER JOIN Customer C ON A.CustCode = C.CustCode "&_
'			"WHERE IsCredit=1 and  C.CustCode = '"& Custcode &"' and " & ChkSql("'ExpContract'", "A.ContractId", "2,5,6"), G_DBConn, 1, 1, 1
'	do while not RSC.EOF
'		ProdAmt = CDBL(RSC("ProdAmt"))
'		RecConAmt = CDBL(RSC("RecAmt"))
'		if RSC("ConCurr") <> "USD" then
'			ProdAmt = ProdAmt * CDBL(RSC("ExRate")) /  UsdExRate 
'			RecConAmt = RecConAmt * CDBL(RSC("ExRate")) / UsdExRate
'		end if
'
'		CustUsedCredit = CustUsedCredit + CDbl(ProdAmt) + CDbl(RSC("AddInAmt")) - CDbl(RecConAmt)
'		RSC.MoveNext
'	loop
'	RSC.Close 
'	set RSC = Nothing
'end function


'===============================================================
'	防止刷新时重复Submit
'===============================================================
	function Stopsubmit(Sort, Id)
		Submit = Request(Sort &"Submit")
		select case sort
			case "ExpContract"
				ObjectId="?ContractId="
			case "ValidateExpCon"		'出口合同生效审核
				ObjectId="?Errorx=" & Errorx &"&ContractId="				
			case "DomContract"
				ObjectId="?Errorx=" & Errorx &"&DomId="
			case "AgentContract"
				ObjectId="?DomId="
			case "PayApplication"
				ObjectId="?AccId="
			case "Bhd"
				ObjectId="?BhdId=?"
			case "JobReport"
				ObjectId="?ReportId="	
			case "Invoice"
				ObjectId="?InvId="	
			case "Evection"
				ObjectId="?EvectionId="	
			case "EvectionReport"
				ObjectId="?ReportId="	
			case "LeaveApply"
				ObjectId="?LeaveId="
			case "UsecarApply"
				ObjectId="?UsecarID="
			case "EvectionReport"
				ObjectId="?ReportID="	
			case "NMSalCon"
				ObjectId = "?ConId="
			case "NMBuyCon"
				ObjectId = "?ConId="
			case "NMBuyOrder"
				ObjectId = "?BuyOrderId="
			case "NMActualBudget"
				ObjectId = "?ActualBudgetId="	
			case "NMPayApply"
				ObjectId = "?PreSubmit="& Submit & "&Errorx=" & Errorx & "&PayId="
			case "RouteConfirm"
				ObjectId = "?RouteConfirmId="	
			Case "ExpInv"
				ObjectId = "?Errorx=" & Errorx &"&InvId="
			Case "DomInvHK"
				ObjectId = "?Errorx=" & Errorx &"&InvId="
			Case "DomInvPK"
				ObjectId = "?Errorx=" & Errorx &"&InvId="	
			Case "TrsInv"
				ObjectId = "?Errorx=" & Errorx &"&InvId="
			Case "ConsignmentContract"
				ObjectId = "?Errorx=" & Errorx &"&ContractId="
			Case "ConsignDomCon"
				ObjectId = "?DomId="
			Case "OtherIncome"
				ObjectId = "?IncomeId="
			Case "CurrExchange"
				ObjectId = "?ExId="
			Case "CurrTransfer"
				ObjectId = "?TrId="	
			Case "Offer"
				ObjectId = "?OfferId="
			Case "Supply"
				ObjectId = "?OfferId="				
										
		end select
		if Submit<>"" or Submits<>"" then 
			response.redirect request.ServerVariables("SCRIPT_NAME")&ObjectId&Id
			response.end
		end if
	end function
	
'======================================================================
'用于按审核状态排序的函数。用法：作为一个临时字段放在主语句中，Order by 后面直接引用临时字段。
'OrderBy(Parameter1,Parameter2),Parameter1审核对象，Parameter2对象ID
'======================================================================
	function OrderBy(ObjectCode, Id)

		OrderBy = "ISNULL((SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectCode = '"& ObjectCode &"' AND ObjectId = "& Id &"),0)"
	
	end function
	
'======================================================================
'用于仅显示审核信息，不带有操作按钮
	function ChkInfo(sort,Id)
		Set RSS = Server.CreateObject("ADODB.Recordset")
		
	 	response.Write("<table width=""80%"" align=""center"" style=""font-size:14px"">")

		RSS.open " SELECT ChkName,ChkDate FROM CensorProcess WHERE (ObjectCode = '"& Sort &"') AND (ObjectId = '"& Id &"')",G_DBConn,3,1,1
			if RSS.eof = true then 
				response.Write("<tr>")
				response.Write("<td width=""33%"">审核人：</td>")
				response.Write("<td width=""33%"">审核日期：</td>")
				response.Write("<td width=""33%"" rowspan='" & RowCount &"'>审核状态：尚未提交</td>")
				response.Write("</tr>")
			end if
			
			RowCount = RSS.recordcount
			Recount=1
			do while RSS.eof = false 
				if Recount = 1 then 
					response.Write("<tr>")
					response.Write("<td width=""33%"">审核人：" & RSS("ChkName") & "</td>")
					response.Write("<td width=""33%"">审核日期：" & RSS("ChkDate") & "</td>")
					response.Write("<td width=""33%"" rowspan='" & RowCount &"'>审核状态：" & ChkResult(CensorResult(Sort, ID)) &"</td>")
					response.Write("</tr>")
				else 
					response.Write("<tr>")
					response.Write("<td>审核人：" & RSS("ChkName") & "</td>")
					response.Write("<td>审核日期：" & RSS("ChkDate") & "</td>")
					response.Write("</tr>")
				end if
				Recount = Recount + 1
			RSS.movenext
			Loop
		RSS.close
    	response.Write("</table>")
		response.Write("<hr>")		
   		
		set RSS = nothing
	end function

	
'替换审核条件中的判断条件代码为判断条件的值
'一个判断条件代码不能是另一个条件代码中的一部分，不能是 PrePayAmt 和 PayAmt
function ReplaceParam(Sort, Id, sTerm)
	Set RS = Server.CreateObject("ADODB.Recordset")
	Set RS_1 = Server.CreateObject("ADODB.Recordset")
	
	RS.Open "SELECT ParamCode, ParamValue FROM CensorObject A INNER JOIN CensorObjectParam B ON A.ObjectId = B.ObjectId "&_
			"WHERE A.ObjectCode = '"& Sort &"'", G_DBConn, 0, 1, 1
	do while not RS.EOF
		ParamCode = RS("ParamCode")
		ParamValue = Replace(RS("ParamValue"), "@Id", Id, 1, -1, 1)

		RS_1.Open ParamValue, G_DBConn, 0, 1, 1
		if not RS_1.EOF then
			sTerm = Replace(sTerm, ParamCode, RS_1(0).Value, 1, -1, 1)
		else
			'如果没有找到判断参数的值，替换为0值
			sTerm = Replace(sTerm, ParamCode, "null", 1, -1, 1)
		end if
		RS_1.Close
		RS.MoveNext
	loop
	RS.Close

	ReplaceParam = sTerm
		
	set RS = nothing
	set RS_1 = nothing
end function
'判断用户是否为总经理
function IsManage()
	set RS = Server.CreateObject("ADODB.Recordset")
	strResult="False"
	RS.open "select * from EmployeeRole where EmpCode='"&userId&"'",G_DBConn,1,1,1
	do while not RS.eof
		GroupCode=RS("GroupCode")
		if GroupCode="GSLD" then
			strResult="True"
		end if
		RS.movenext
	loop
	RS.close
	if strResult="True" then
		strSql=" 1=1 "
	else
		strSql=" 1<>1 "
	end if
	IsManage=strSql
end function
'判断用户是否为抄送人员
function IsDetail(EmpCode,ExamineId)
	set rsDetail = Server.CreateObject("ADODB.Recordset")
	rsDetail.open "select * from Examine where ExamineId="&ExamineId&"",G_DBConn,1,1,1
	if not rsDetail.eof then
		DetailRight=rsDetail("DetailRight")
	end if
	rsDetail.close
	if instr(DetailRight,EmpCode)>0 then
		IsDetail="True"
	else
		IsDetail="False"
	end if
end Function
</script>