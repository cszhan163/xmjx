<script language="VBS" runat="Server">
'================================================================================================================================
'	��ʾÿ����Ա��˽����Ӧ������˵��
'================================================================================================================================
function ChkState(State)
	if IsNull(State) then
		ChkState = "δ��"
	else
		select case State
			case False		ChkState = "���"
			case True		ChkState = "ͨ��"
		end select
	end if
end function

'================================================================================================================================
'	��ʾ���״̬��Ӧ������˵��
'================================================================================================================================
function ChkResult(Status)
	if IsNumeric(Status) then
		select case Status
			case 0		ChkResult = "��δ�ύ"					'Self	ProcTable	����Ҫʹ����˹���
			case 1		ChkResult = "�Ѿ��ύ"					'Censor	RuleTable	��Ҫʹ����˹���
			case 2		ChkResult = "�����˻�"					'Censor	RuleTable
			case 3		ChkResult = "���δͨ��"				'Self	ProcTable
			case 4		ChkResult = "�Ѿ��˻�"					'Self	ProcTable
			case 5		ChkResult = "�������"					'Self	ProcTable
			case 6		ChkResult = "���ͨ��"					'Self	ProcTable	AllSee
		end select
	else
		ChkResult = Status
	end if
end function

'===============================================================================================================================
'	����������Ĳ�����ť	��˶����������״̬(Status),	��˹���Id(RuleId)
'===============================================================================================================================
function ChkButton(Sort,Status, IsCensor, RuleId, EmpCode)
	if Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O" then		'λ�����ҳ��(N|O)
		if SeeEmp(EmpCode, "Chk") then						'�û�������˵�ǰ�ĵ���
			sClk = "Rule"& RuleId &".name=""RuleId"": Rule"& RuleId &".value="""& RuleId &""""
		else												'�û���������˵�ǰ�ĵ���
			sClk = CanOper(EmpCode, "", "", "Chk")
		end if

		select case Status					'��Ϊ����һ������˶������,�����(ͬ��,��ͬ��)ʱ,ͬʱ�ύ��ǰ��˵Ĺ����Id (RuleId)
			case 1
				ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""ͬ��"" language=""VBS"" onclick='"& sClk &"'> "&_
							"<input type=""submit"" name="""& Sort &"Submit"" value=""��ͬ��"" language=""VBS"" onclick='"& sClk &"'>"
			case 2
				ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""�˻�"" language=""VBS"" onclick='"& sClk &"'>"
		end select
		ChkButton = ChkButton &"<input type=""hidden"" id=""Rule"& RuleId &""" name="""" value="""">"
	else									'�û�ҵ��״̬
		'�û������ύ���˻ذ�ťʱ���ж��û��Ƿ�Ե�ǰ�������޸�Ȩ�ޣ�����ÿ��ҳ��� EdiOper()���� 2007.11.23
		select case Status
			case 0, 3, 4
				if DisplayChkButton(Sort, Id, 1) then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""�ύ"" onclick=""EdiOper()"">"
				end if
			case 5							'�����������ʱ,��������˶���ǰ��Ϊ����ʱ,��ʾ�ύ��ť,�ṩת�������ѭ���Ĺ���
				if IsCensor then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""�ύ"" onclick=""EdiOper()"">"
				end if
			case 1, 6
				if Status = 6 then
					sClk = "Rt = MsgBox(""�õ��������ͨ����ȷʵҪ�˻���"", vbOKCancel + vbQuestion + vbDefaultButton2, ""ȷ��""): if Rt = vbCancel then window.event.returnValue = false end if: "
					if Sort = "Invoice" or Sort = "NMActualBudget" then 		'��·����Ҫ�󣺲��������ͨ�����û�������˻ع��ܣ���Ҫ�ܾ��������ύ״̬��
						exit function
					end if
				end if
				if DisplayChkButton(Sort, Id, 2) then
					ChkButton = "<input type=""submit"" name="""& Sort &"Submit"" value=""�����˻�"" language=""VBS"" onclick='"& sClk &" EdiOper'>"
				end if
		end select
	end if
end function

'================================================================================================================================
'	��˲�������	��˶�������(Sort): ��CensorObject�е�ObjectCode, �����ID(Id)
'================================================================================================================================
function CensorOper(Sort, Id)
	Submit = Request(Sort &"Submit")
	RuleId = Request("RuleId")
	InureMsg=Request("InureMsg")
	if Submit = "�ύ" then	
		set RS = Server.CreateObject("ADODB.Recordset")	
		if Sort="ExpContract" then 			'������ͬ�ύʱ�ж��Ƿ�����ύ������
			If AllowSubmit(Id) =0 then 
				'�ύʱ��Ҫָ����(CensorRules)�е���˹���
				if Request(Sort &"RuleId").Count <> 0 then
					G_DBConn.Execute "DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
					for each r in Request(Sort &"RuleId")
						G_DBConn.Execute "INSERT INTO CensorProcess(ObjectCode, ObjectId, RuleId, SubmitDate, ChkResult) "&_
									   "VALUES('"& Sort &"', '"& Id &"', '"& r &"', '"& Date &"', 1)"
					next
				else
					'��δָ������ʱ,�ж���˶����Ƿ�����Ҫ���,�粻�����,ֱ����Ϊ�������
					G_DBConn.Execute "IF (SELECT COUNT(*) FROM CensorRules WHERE ObjectCode = '"& Sort &"') = 0 "&_
								   "BEGIN DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'; "&_
								   "INSERT INTO CensorProcess(ObjectCode, ObjectId, ChkResult) VALUES('"& Sort &"', '"& Id &"', 5) END"
				end if
			else
				response.end
			end if 
		else
			'�ύʱ��Ҫָ����(CensorRules)�е���˹���
			'===========��ͬ��Ч��ʱ��������Ч˵��
			if Request(Sort &"RuleId").Count <> 0 then
				G_DBConn.Execute "DELETE FROM CensorProcess WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
				for each r in Request(Sort &"RuleId")
					G_DBConn.Execute "INSERT INTO CensorProcess(ObjectCode, ObjectId, RuleId, SubmitDate, ChkResult,InureMsg) "&_
								   "VALUES('"& Sort &"', '"& Id &"', '"& r &"', '"& Date &"', 1,'"&InureMsg&"')"
				next
			else
				'��δָ������ʱ,�ж���˶����Ƿ�����Ҫ���,�粻�����,ֱ����Ϊ�������
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

			if bPrePay then 	'���ύ��˹����У��ж��Ƿ�¼�뷢Ʊ�š����û�з�Ʊ�ţ�����Ԥ�����Ͳ�����Ҫ���ˣ�����з�Ʊ�ţ�������Ҫ����(3)
				Response.Write "<body onclick=""location.replace('PaymentReportEdit.asp?AccId="& Id &"')""><center><font color=red>ȱ�ٷ�Ʊ�ţ���ǰΪ���֧������Ҫȷ¼����ȷ�ķ�Ʊ�ţ�</font></center>"
				Response.End
				'ErrMsg("ȱ�ٷ�Ʊ�ţ���ǰΪ���֧������Ҫȷ¼����ȷ�ķ�Ʊ�ţ�")
			else
				G_DBConn.Execute "update AccountFee set IsPrePay=3 where AccId='"& Id &"' "
			end if
		end if 
		
		set RS = Nothing
	end if
	
	if Submit = "�����˻�" then
		'����ɾ����(CensorProcess)���Ѿ��ڱ�(CensorRules)�в����ڵ���˹���,
		'Ȼ�������˶���ǰ�Ƿ��Ѿ�����˹�,���û��ֱ����Ϊ��δ�ύ,������Ϊ�����˻�,
		'����ж���˶�������˱�(CensorProcess)�Ƿ�����˼�¼,��û�и��ݶ����ڱ�(CensorObject)���Ƿ���Ҫ��˵�״̬,
		'д����δ�ύ(�й���),�������(�޹���)
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
	
	if Submit = "ͬ��" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkState = 1, ChkEmpCode = '"& UserId &"', ChkDate = '"& Date &"', "&_
					   "ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = "& RuleId
	end if
	
	if Submit = "��ͬ��" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkState = 0, ChkEmpCode = '"& UserId &"', ChkDate = '"& Date &"', "&_
					   "ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = "& RuleId
		if Sort ="PayApplication" then		''������˻�ʱ����ͬ���״̬һͬ�˻ء�
			G_DBConn.Execute "update AccountFee set IsPrePay=1 where AccId='"& Id &"' "
		end if
	end if
	
	if Submit = "�˻�" then
		G_DBConn.Execute "UPDATE CensorProcess SET ChkMessage = '"& Valid(Request("ChkMessage"& RuleId)) &"' "&_
					   "WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"' AND RuleId = '"& RuleId &"'; "&_
					   "UPDATE CensorProcess SET ChkResult = 4 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
					   
		if Sort ="PayApplication" then	''������˻�ʱ����ͬ���״̬һͬ�˻ء�
			G_DBConn.Execute "update AccountFee set IsPrePay=1 where AccId='"& Id &"' "
		end if
	end if
end function

'===============================================================================================================================
'	����û����˼�¼����˶���,���ݵ�ǰ����������(CensorProcess)д����˼�¼
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
'	�����ύ���������Ա����	������(DeptName),	�����(GroupName),	ְԱ��(EmpName)
'===============================================================================================================================
function ChkName(DeptName, GroupName, EmpName)
	if EmpName <> "" then						'����������,��ʾ������,������ʾ������������������
		ChkName = EmpName
	else
		'if DeptName <> "" and GroupName <> "" then
		'	sSpace = " "
		'end if
		ChkName = DeptName & sSpace & GroupName
	end if
end function

'===============================================================================================================================
'	ȡ��������ô�(CensorRules)����˶����ĳһ���������	��˶���(ObjectCode),	��˼���(Level): 1 | 2 | 3
'===============================================================================================================================
function CensorLevel(ObjectCode, EmpCode, Level)
	set RS = Server.CreateObject("ADODB.Recordset")
	set CensorLevel = Server.CreateObject("Scripting.Dictionary")		'���� Dictionary ����

 
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
			'�ж��ڶ�������ģ��CensorModuleCode����ָ���������ChkEmpCode�Ƿ��EmpCode�ĵ��������Ȩ��
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
'	ȡ��������ô�(CensorRules)����˶������˼���	��˶���(ObjectCode)
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
'	�жϵ�ǰ�û��Ƿ��Ƕ������˹����еĿ�����Ա
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
'	���õ��ݵ�������˽��,ȡ�õ�ǰ(CensorProcess)����˼���	��˶�������(Sort), �����ID(Id)
'===============================================================================================================================
function CurCensorLevel(Sort, Id)
	set RS = Server.CreateObject("ADODB.Recordset")
	set LevelDenyState = Server.CreateObject("Scripting.Dictionary")
	set LevelPassState = Server.CreateObject("Scripting.Dictionary")
	Submit = Request(Sort &"Submit")
	
	'ִ���û���˲���
	CensorOper Sort, Id

	'�ӱ� CensorProcess ��ȡ�õ����ύ���������
	RS.Open "SELECT A.ChkState, ISNULL(R.CensorLevel, '') CensorLevel, ISNULL(R.InnerLevelCode, '') InnerLevelCode, "&_
			"ISNULL(R.DenyTerm, '') DenyTerm, ISNULL(R.PassTerm, '') PassTerm, "&_
			"(SELECT NeedAllCensor FROM CensorObject WHERE ObjectCode = '"& Sort &"') NeedAllCensor "&_
			"FROM CensorProcess A LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
			"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' ORDER BY R.CensorLevel ASC", G_DBConn, 0, 1, 1	
	do while not RS.EOF
		NeedAllCensor = RS("NeedAllCensor")				'�Ƿ���Ҫ�ύȫ����˼������(true)

		if CurLevel <> RS("CensorLevel") then
			DenyTerm = RS("DenyTerm")
			DenyTerm = Replace(DenyTerm, "AND", "and", 1, -1, 1)		'�滻�����еĲ�����ΪСд����ʹ�滻Ϊ��˽��ʱ�����滻(AND)��A
			DenyTerm = Replace(DenyTerm, "OR", "or", 1, -1, 1)
			
			PassTerm = RS("PassTerm")
			PassTerm = Replace(PassTerm, "AND", "and", 1, -1, 1)
			PassTerm = Replace(PassTerm, "OR", "or", 1, -1, 1)

			'�滻���ͨ�����������еĲ��������ͬ���
			PassTerm = ReplaceParam(Sort, Id, PassTerm)
			DenyTerm = ReplaceParam(Sort, Id, DenyTerm)
		end if
		CurLevel = RS("CensorLevel")									'< ���յ����ύ������ߵ���˼��� >
		
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


	'�滻��������е�ǰδ�ύ��������˵Ĵ���(A|B|...)Ϊ��(null = 0|1)
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
	
	'�õ���ǰ����˼���,����������˴��ĵ���(ChkResult = 1|2)���ص���˽����ȷ,
	'����״̬���ؽ�����ܲ���ȷ(����˹����Ѳ��ڱ�(CensorRules)��ʱ������
	for each level in LevelDenyState
		CurCensorLevel = level
		if LevelDenyState.Item(Level) <> "" then			'�����(CensorProcess)��ǰ����˹������ڱ�(CensorRules)��
			DenyState = Eval(LevelDenyState.Item(level))					'ĳһ���ķ���������(True|False)
		end if
		if LevelPassState.Item(Level) <> "" then
			PassState = Eval(LevelPassState.Item(level))					'ĳһ����ͨ���������(True|False)
		end if

		'if DenyState or IsNull(DenyState) and not PassState or IsNull(PassState) then	'�����ǰ����˼���δͨ����δ�����ͣ���ڴ˼�
		if not PassState or IsNull(PassState) then
			exit for
		end if
	next

	' ==================================================================================<<< ������˽�� >>>======================
	'��ִ������˲���,ȡ�õ�ǰ��˼����,�жϵ��ݵ�������˽��
	if Submit = "ͬ��" or Submit = "��ͬ��" or Submit = "����" then		'��ͬ��,��ͬ��,��������ô������������ʱ(CensorRulesEdit.asp Line167),��Ҫ���¼��㵥��������˽��
		TotalCenLevel = CensorLevelCount(Sort)			'��ǰ��˶������õ��ܵ���˼���
		SubmitCenLevel = LevelDenyState.Count 
		if DenyState then												'�κ�һ��δͨ��ʱ,�����������շ��
			G_DBConn.Execute "UPDATE CensorProcess SET ChkResult = 3 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"
		else
			'NeedAllCensor=true ʱ������˶����������˼���ͨ����,��������ͨ��
			'NeedAllCensor=false ʱ�������ύ�����һ��ͨ��ʱ,��������ͨ��
			if NeedAllCensor and PassState and CurCensorLevel = TotalCenLevel and SubmitCenLevel = TotalCenLevel _
				or not NeedAllCensor and PassState and CurCensorLevel = CurLevel then
				G_DBConn.Execute "UPDATE CensorProcess SET ChkResult = 6 WHERE ObjectCode = '"& Sort &"' AND ObjectId = '"& Id &"'"

				'������ں�ͬ���ͨ����,���ݵ�ǰ���÷�������,���������ռ�ñ�ǣ�������������������
				if Sort = "ExpContract" then
					SetIsCredit Id
				end if
				
			end if
		end if
		
		'�ж��Ƿ���˽���
		if DenyState or NeedAllCensor and PassState and CurCensorLevel = TotalCenLevel and SubmitCenLevel = TotalCenLevel _
			or not NeedAllCensor and PassState and CurCensorLevel = CurLevel then
			NeedSetChkName = true
		end if
	end if

	if Submit = "�˻�" then
		'�˻�ʱ���ݶ����Ƿ���Ҫȫ�������ִ���˻ز������жϵ�ǰ�Ƿ������Ϊ���˻�״̬
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

	'����˽���ʱ(�������ͨ�����������),���ύ������˼�����˵�����д��(CensorProcess)�� ChkName��,
	'�Ժ���ʾ�������Ϣȫ��������(CensorProcess),��(CensorRules)�޸���˹����������ȷ��ʾ��ʱ��������
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
'	��ʾ��ҵ�񴦵��ύ����ĺ�ѡ����˺��ύ��ť	��˶���(Sort),	����ID(Id),	 �����ְԱ����(EmpCode), 
'	�������˽��(Result):0 | 1,	��ǰ����˼���(CurLevel)
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
			case 0, 3, 4, 5									'δ�ύ,δͨ��,�˻�ʱ,�ӱ�(CensorRules)��ʾ���м������˺�ѡ��
				LevelCount = CensorLevelCount(Sort)
				Response.Write "<tr><td width=""85%"" align=""left"" colspan=""3""><span style=""width:85%"">"
				if not(Result = 5 and not IsCensor) then		'�����������ʱ,��������˶���ǰ��Ϊ����ʱ,��ʾ��˺�ѡ��,�ṩת�������ѭ���Ĺ���
					for i = 1 to LevelCount
						'ȡ��ĳһ�����������
						set Content = CensorLevel(Sort, EmpCode, i)
						if i <> 1 then
							Response.Write "<br>"
						end if

						Response.Write "<b>"& i &".</b>"
						for each RuleId in Content
							'����CensorRules:Fixed���жϵ�ǰ��˹����Ƿ�����û��޸�
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
				'===============������Ч˵��
				Response.Write "</span><Br><span style=""width:85%;"">"
				If Sort="ValidateExpCon" Then '��Ч����ˣ������Ч˵��		
				Response.Write "��Ч˵��:<input type=""text"" name=""InureMsg"" value="""" class=""input"" maxlength=""400"" style=""width:80%"">"				
				End If
				Response.Write "&nbsp;</span><span style=""width:15%; word-wrap:normal""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></span></td>"&_
							   "<td width=""15%"" valign=""bottom"">"& ChkButton(Sort,Result, IsCensor, 0, "") &"</td></tr>"
			case 1										'���ύʱ,�ӱ�(CensorProcess)��ʾ�ȴ���˵�ĳһ��ѡ���������
				set RS = Server.CreateObject("ADODB.Recordset")
			
				Response.Write "<tr><td align=""left"">"
				RS.Open "SELECT A.RuleId,A.InureMsg, D.DeptName, G.GroupName, E.EmpNameChs "&_
						"FROM CensorProcess A LEFT JOIN CensorRules R ON A.RuleId = R.RuleId "&_
						"LEFT JOIN Dept D ON R.DeptCode = D.DeptCode "&_
						"LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
						"LEFT JOIN Employee E ON R.EmpCode = E.EmpCode "&_
						"WHERE A.ObjectCode = '"& Sort &"' AND A.ObjectId = '"& Id &"' AND R.CensorLevel = '"& CurLevel &"'", G_DBConn, 0, 1, 1
				If Sort="ValidateExpCon" Then '��Ч����ˣ������Ч˵��				
				Response.Write "��Ч˵����<font color=red>"&RS("InureMsg")&"</font>"				
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
				If Sort="ValidateExpCon" Then '��Ч����ˣ������Ч˵��		
				Response.Write "��Ч˵����<font color=red>"&InureMsg&"</font>"				
				End If
				Response.Write "&nbsp;</td><td align=""right"" valign=""bottom"" colspan=""2"" style=""padding-right:30px""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></td>"&_
							   "<td width=""15%"">"& ChkButton(Sort,Result, IsCensor, 0, "") &"</td></tr>"
		end select
	else
		Response.Write "<tr><td align=""left"">"
		If Sort="ValidateExpCon" Then '��Ч����ˣ������Ч˵��
		Response.Write "��Ч˵����<font color=red>"&InureMsg&"</font>"	
		End If		
		Response.Write "&nbsp;</td><td align=""right"" valign=""bottom"" colspan=""2"" style=""padding-right:30px""><b id="""& Sort &"ChkResult"">"& ChkResult(Result) &"</b></td>"&_
					   "<td width=""15%""></td></tr>"
	end if

	set RS = nothing
end function

'================================================================================================================================
'	��ʾ�����Ϣ	��˶�������(Sort): ��CensorObject�е�ObjectCode, �����ID(Id), ���ݵ��û�����(EmpCode)
'================================================================================================================================
function CensorInfo(Sort, Id, EmpCode)
	set RS = Server.CreateObject("ADODB.Recordset")
	set RSTP = Server.CreateObject("ADODB.Recordset")

	'�жϵ�ǰ�Ķ����Ƿ���Ҫ���
	RS.Open "SELECT * FROM CensorObject WHERE ObjectCode = '"& Sort &"'", G_DBConn, 0, 1, 1
	if RS.EOF then
		stop							'�Բ���Ҫ��˵Ķ�����ʾ�����Ϣ
		exit function
	else
		IsCensor = RS("IsCensor")		'�����Ƿ���Ҫ���
	end if	
	RS.Close 	
	Result = 0								'���ݵ���˽��
	CurLevel = CurCensorLevel(Sort, Id)		'��ǰ���ͣ���ڴ˼�
	'����˶����ڱ�(CensorProcess)�б���������һ����˼�¼,���û��ͨ�����º�������
	SetCensorProcess Sort, Id, IsCensor
	Response.Write "<table class=""pagetable"">"

	'��ȡ�����ȫ�������Ϣ,����ʾ��
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
	Result = RSTP("ChkResult")			'���ݵ���˽��
	if Result <> 0 and Result <> 5 then					'�����ݲ�Ϊ�������ʱ,��ʾ��ϸ������
		do while not RSTP.EOF
			RuleId = RSTP("RuleId")				'
			ChkEmpName = RSTP("EmpNameChs")		'���������Ա����
			State = RSTP("ChkState")				'������Ա����˽��
			
			'���ύ,�����˻�ʱ����˴�(CensorRules)��ȡ,���������(CensorProcess)��ChkName
			select case Result
				case 1, 2
					Name = ChkName(RSTP("DeptName"), RSTP("GroupName"), RSTP("EmpName"))
				case else
					Name = RSTP("ChkName")
			end select

			'���ύ�������Ϊ���Ż����ʱ,����ʵ����˵���Ա��
			ChkEmpName = RSTP("EmpNameChs")
			if ChkEmpName <> Name and ChkEmpName <> "" then
				Name = Name &"("& ChkEmpName &")"
			end if
			'��ʾ��ϸ�����Ϣ
			Response.Write "<tr><td width=""40%"" nowrap>�����: "& Name &"</td><td width=""25%"" nowrap>���ʱ��: "& RSTP("ChkDate") &"</td>"&_
						   "<td width=""20%"" nowrap>���״̬: "& ChkState(State) &"</td><td width=""15%"" nowrap>�ύʱ��: "& RSTP("SubmitDate") &"</td></tr>"

			if (Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O") and (Result = "1" and IsNull(State) or Result = "2") and CanChk(RuleId) then
				'����û�������ˣ���ʾ����˵�¼������˲�����ť
				Response.Write "<tr><td colspan=""3"">���: <input type=""text"" name=""ChkMessage"& RuleId &""" value="""& RSTP("ChkMessage") &""" class=""input"" maxlength=""400"" style=""width:80%""></br>"&_
			" <td align=""right"">"& ChkButton(Sort,Result, IsCensor, RSTP("RuleId"), EmpCode) &"</td></tr>"
			else																		'����ʾ�����Ϣ
				Response.Write "<tr><td colspan=""3"">���: <font color=blue>"& RSTP("ChkMessage") &"</font></td><td></td></tr>"				
			end if

			RSTP.MoveNext
		loop
	end if
	RSTP.Close
	Stopsubmit Sort, Id
	'��ҵ����ʾ�ȴ�����˼�������ť,	��˴���ʾ��ǰ����˽��
	CensorState Sort, Id, EmpCode,IsCensor, Result, CurLevel
	Response.Write "</table><hr>"	
	set RS = nothing
end function

'================================================================================================================================
'	��ʾ���״̬��ѯѡ��
'================================================================================================================================
function ChkQuery()
	ChkQ = CurSelValue("ChkQuery")
	dim opt(6)
	if ChkQ <> "" then
		opt(ChkQ) = "selected"
	end if
	
	Response.Write "<select name=""ChkQuery""><option value="""">���״̬</option>"
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
'	����ҳ����Բ鿴�����ݵ����״ֵ̬(0|1|...)
'================================================================================================================================
function ChkValue()
	'����ʹ�����״̬��ѯָ������˽��(��λ���б�ҳʱ),�༭ҳû�����״̬��ѯѡ��.2006.1.5
	'ChkValue = CurSelValue("ChkQuery")
	
	if ChkValue = "" then
		if Left(ModuleCode, 1) = "N" or Left(ModuleCode, 1) = "O" then			'λ�����ģ��
			ChkValue = "1, 2, 6"
		else
			ChkValue = "0, 1, 2, 3, 4, 5, 6"
		end if
	end if
end function

'================================================================================================================================
'	�����������״̬���ƴ�	��˶���(ObjectCode), �������ݱ��ID����(Id),	��ѯ����˽��(Result)	������WHERE��( WHERE "& ChkSql("'ExpContract'", "A.ContractId", 6) &"..." )
'================================================================================================================================
function ChkSql(ObjectCode, Id, Result)
	'����ѯ���״̬ʱ
	if Request("ChkQuery") <> "" then
		Result = Request("ChkQuery")
	end if

	'�������״̬����������, �����ObjectCode���������������:'ImpContract' ����
	if Result <> "" then
		ChkSql = "ISNULL((SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectId = "& Id &" AND ObjectCode = "& ObjectCode &"), 0) "&_
				 "IN ("& Result &")"
	end if
end function

'================================================================================================================================
'	������˽��(0|1|...)	��˶���(ObjectCode): ��CensorObject�е�ObjectCode, �����ID(Id)
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
'	������˽����ǰ�鿴��Ա����˽��(0|1|ͨ��|���|...)	��˶���(ObjectCode): ��CensorObject�е�ObjectCode, �����ID(Id)
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
	
	'���״̬Ϊ�Ѿ��ύ,������ʾ���ҳ��ʱ,�жϵ�ǰ�û������״̬
	if CensorResult2 = 1 and (Left(ModuleCode ,1) = "N" or Left(ModuleCode, 1) = "O") then
		'ȡ�ö���ǰ����˼���
		ObjCurCensorLevel = CurCensorLevel(ObjectCode, Id)
		
		for l = 1 to ObjCurCensorLevel
			Needed = 0
			Passed = 0
			Denyed = 0
			NoChked = 0
			'�ڵ�ǰ��˼����ڲ��ҵ�ǰ�û���Ҫ��˵Ĺ���
			RS.Open "SELECT A.ChkState, A.RuleId FROM CensorProcess A LEFT JOIN CensorRules B ON A.RuleId = B.RuleId "&_
					"WHERE A.ObjectCode = '"& ObjectCode &"' AND A.ObjectId = '"& Id &"' "&_
					"AND B.CensorLevel = '"& l &"'", G_DBConn, 0, 1, 1
			do while not RS.EOF
				if CanChk(RS("RuleId")) then		'������Ҫ��˵Ĺ���,�����û����е���˽��
					Needed = Needed + 1
					if RS("ChkState") then
						Passed = Passed + 1			'ͨ���ĸ���
					end if
					if not RS("ChkState") then
						Denyed = Denyed + 1			'�񶨵ĸ���
					end if
					if IsNull(RS("ChkState")) then
						NoChked = NoChked + 1
					end if
				end if
				RS.MoveNext
			loop
			RS.Close
		
			'�ڵ�ǰ��˼�����,��ǰ�û�����Ĺ���ȫ��ͨ��ʱ,���'ͨ��',��һ����ͨ�����'���',��һ��δ��ʱ���'δ��'
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
'	�жϲ�����ť�Ƿ���ʾ�ĺ���	��˶���(ObjectCode),	��˶���Id(ID),	Ҫ��ʾ�İ�ť����(OptBtn): (����"Save"|��ӡ"Print")
'================================================================================================================================
function Visible(ObjectCode, Id, OptBtn)
	if ObjectCode <> "" and Id <> "" then
		'ȡ����˶�������״̬
		Result = CensorResult(ObjectCode, Id)
	else
		Result = 0								'���ָ������˶�������ID��Ч,��Ϊ��δ�ύ
	end if
	
	select case OptBtn
		case "Save"					'Ҫ��ʾ���水ť
			select case Result
				case 0, 3, 4, 5		Visible = true
				case else			Visible = false
			end select
		case "Print"				'Ҫ��ʾ��ӡ��ť
			select case Result
				case 5, 6			Visible = true
				case else			Visible = false
			end select
	end select
end function


'================================================================================================================================
'	ȡ�õ�ǰ�û��Ƿ�����ĳһ���(True | False)	��ݴ���(GroupCode)
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
'	ȡ����˶���ĳһ��������˵�����	��˶���(ObjectCode),	��˶���Id(ID),	��˼���Level
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
'	�����Ƿ�Ӧ��ʾ��˲�����ť(��True ��False)	��˶���(Sort), ����Id(Id), ��ť����(BtnType):�ύ1 �����˻�2
'================================================================================================================================
function DisplayChkButton(Sort, Id, BtnType)
	set RS = Server.CreateObject("ADODB.Recordset")
	
	if BtnType = 1 then						'Ҫ��ʾ�ύ��ť
		select case Sort
			case "Contract"								'��ͬ�����
				DisplayChkButton = true
			case "SaleContract"							'���ۺ�ͬ
				RS.Open "SELECT ContractId FROM Contract A LEFT JOIN CensorProcess CP ON A.ContractId = CP.ObjectId AND CP.ObjectCode = 'Contract' "&_
						"WHERE A.ContractId = '"& Id &"' AND CP.ChkResult = 6", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true				'��ͬ��Ӧ����������ͨ������ʾ�ύ��ť
				else
					DisplayChkButton = false
				end if
				RS.Close
			case "PlanProduct"							'�Ų���
				RS.Open "SELECT C.ContractId FROM PlanProduct A LEFT JOIN Contract C ON A.ConId = C.ContractId "&_
						"LEFT JOIN CensorProcess CP ON C.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.PlanId = '"& Id &"' AND CP.ChkResult = 6", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true
				else
					DisplayChkButton = false
				end if
			case "Bhd"									'������
				RS.Open "SELECT C.ContractId FROM Bhd A LEFT JOIN Contract C ON A.ConId = C.ContractId "&_
						"LEFT JOIN CensorProcess CP ON C.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.BhdId = '"& Id &"' AND (CP.ChkResult = 6 OR A.ConId = 0)", G_DBConn, 0, 1, 1
				if not RS.EOF then
					DisplayChkButton = true
				else
					DisplayChkButton = false
				end if
			case else
				if Id <> "-2" then						'����������Ѿ����ڼ�¼�Ķ�����ʾ�ύ��ť
					DisplayChkButton = true
				end if
		end select
	else									'Ҫ��ʾ�����˻ذ�ť
		select case Sort
			case "Contract"								'��ͬ�����
				RS.Open "SELECT ContractId FROM Contract A LEFT JOIN CensorProcess CP ON A.ContractId = CP.ObjectId AND CP.ObjectCode = 'SaleContract' "&_
						"WHERE A.ContractId = '"& Id &"' AND CP.ChkResult IN(1, 2, 6)", G_DBConn, 0, 1, 1 
				if not RS.EOF  then
					DisplayChkButton = false
				else
					DisplayChkButton = true
				end if
				RS.Close
			case "SaleContract"							'���ۺ�ͬ
				'�жϴ˺�ͬ�Ƿ���������˵��Ų��� ��������˵ķ�����,�����������˻غ�ͬ
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
			case "PlanProduct"							'�Ų���
				DisplayChkButton = true
			case "Bhd"									'������
				DisplayChkButton = true
			case else
				DisplayChkButton = true
		end select
	end if

	set RS = nothing
end function

'===========================================================================================
'	�ж��Ƿ�����ύ����
'		if ����������Ҫ�Ӽ����� then 
'			if �ͻ�Ϊ�¿ͻ� then 
'				����ۼƺ�ͬ��� ��USD5000 ������������ύ����ʾ���¿ͻ��������Ŷ�ȣ������ύ��
'			else 
'				���ʣ���Ȳ��㣬����������Ч���ѹ��������ύ��
'			end if 
'		else 
'			
'				
'===========================================================================================
Function AllowSubmit(ContractId)
		'�жϺ�ͬ�Ƿ������ύ
		AllowSubmit=0
		Set RSA = Server.CreateObject("ADODB.Recordset")
		RSA.open "select C.Checkresult,C.FinalCFRDate,C.MaxAmt,A.CustCode, B.ConAmt "&_
				 "FROM  Customer C Join Contract A on A.CustCode=C.Custcode "&_
				 "LEFT JOIN VContract B ON A.ContractId = B.ContractId "&_
				 "where A.ContractId = '"& ContractId  &"'", G_DBConn, 0, 1, 1
			if RSA.eof = false then 
				CheckResult=RSA("Checkresult")
				if CheckResult<>6 then 				'û���������
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
'	�����¿ͻ��ۼ�δ�ջ��ͬ��
'==================================================================================
	Function CustUsedAmt (CustId)
		'ȡ����Ԫ����
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
' �����Ͽͻ���ռ�õ����Ŷ��
'========================================================================================
'function CustUsedCredit(CustId)
'	'ȡ����Ԫ����
'	Set RSC = Server.CreateObject("ADODB.Recordset")
'	RSC.Open "SELECT ExRate FROM ExRate WHERE Currency = 'USD'", G_DBConn, 0, 1, 1
'	if not RSC.EOF then
'		UsdExRate = CDBL(RSC("ExRate"))
'	end if
'	RSC.Close 
'	'��������Բ�Ƶ��ܵ����õ����Ŷ��	��Ϊֻ�������ͨ���ĺ�ͬ��ռ�����Ŷ��(GMChkResult=6, 2)
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
'	��ֹˢ��ʱ�ظ�Submit
'===============================================================
	function Stopsubmit(Sort, Id)
		Submit = Request(Sort &"Submit")
		select case sort
			case "ExpContract"
				ObjectId="?ContractId="
			case "ValidateExpCon"		'���ں�ͬ��Ч���
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
'���ڰ����״̬����ĺ������÷�����Ϊһ����ʱ�ֶη���������У�Order by ����ֱ��������ʱ�ֶΡ�
'OrderBy(Parameter1,Parameter2),Parameter1��˶���Parameter2����ID
'======================================================================
	function OrderBy(ObjectCode, Id)

		OrderBy = "ISNULL((SELECT TOP 1 ChkResult FROM CensorProcess WHERE ObjectCode = '"& ObjectCode &"' AND ObjectId = "& Id &"),0)"
	
	end function
	
'======================================================================
'���ڽ���ʾ�����Ϣ�������в�����ť
	function ChkInfo(sort,Id)
		Set RSS = Server.CreateObject("ADODB.Recordset")
		
	 	response.Write("<table width=""80%"" align=""center"" style=""font-size:14px"">")

		RSS.open " SELECT ChkName,ChkDate FROM CensorProcess WHERE (ObjectCode = '"& Sort &"') AND (ObjectId = '"& Id &"')",G_DBConn,3,1,1
			if RSS.eof = true then 
				response.Write("<tr>")
				response.Write("<td width=""33%"">����ˣ�</td>")
				response.Write("<td width=""33%"">������ڣ�</td>")
				response.Write("<td width=""33%"" rowspan='" & RowCount &"'>���״̬����δ�ύ</td>")
				response.Write("</tr>")
			end if
			
			RowCount = RSS.recordcount
			Recount=1
			do while RSS.eof = false 
				if Recount = 1 then 
					response.Write("<tr>")
					response.Write("<td width=""33%"">����ˣ�" & RSS("ChkName") & "</td>")
					response.Write("<td width=""33%"">������ڣ�" & RSS("ChkDate") & "</td>")
					response.Write("<td width=""33%"" rowspan='" & RowCount &"'>���״̬��" & ChkResult(CensorResult(Sort, ID)) &"</td>")
					response.Write("</tr>")
				else 
					response.Write("<tr>")
					response.Write("<td>����ˣ�" & RSS("ChkName") & "</td>")
					response.Write("<td>������ڣ�" & RSS("ChkDate") & "</td>")
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

	
'�滻��������е��ж���������Ϊ�ж�������ֵ
'һ���ж��������벻������һ�����������е�һ���֣������� PrePayAmt �� PayAmt
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
			'���û���ҵ��жϲ�����ֵ���滻Ϊ0ֵ
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
'�ж��û��Ƿ�Ϊ�ܾ���
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
'�ж��û��Ƿ�Ϊ������Ա
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