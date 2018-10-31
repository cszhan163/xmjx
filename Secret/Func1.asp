<script language="VBS" runat="server">
'	�жϺ�ͬ�ĸ�������		����ֵ: ����֤����(True) ����(False)	----------------------------------------------------------------------------------------
function ConPayType(ContractId)
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open "SELECT PaymentEng FROM Contract A LEFT JOIN Payment P ON A.PayTerms = P.PaymentId "&_
			"WHERE A.ContractId = '"& ContractId &"'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		Payments = RS("PaymentEng")
		if Instr(1, Payments, "L/C", 1) > 0 or Instr(1, Payments, "LC", 1) > 0 then		'�ı��Ƚ�
			ConPayType = True
		else
			ConPayType = False
		end if
	else
		stop
	end if
	RS.Close 
	set RS = Nothing
end function


'����ͻ����������Ŷ��(USD)	-------------------------------------------------------------------------------------------------------------
function CustUsedCredit(CustId)
	Set RS = Server.CreateObject("ADODB.Recordset")
	'ȡ����Ԫ����
	RS.Open "SELECT ExRate FROM ExRate WHERE Currency = 'USD'", G_DBConn, 0, 1, 1
	if not RS.EOF then
		UsdExRate = CDBL(RS("ExRate"))
	end if
	RS.Close 
	'��������Բ�Ƶ��ܵ����õ����Ŷ��	��Ϊֻ�������ͨ���ĺ�ͬ��ռ�����Ŷ��(GMChkResult=6, 2)
	CustUsedCredit=0
	RS.Open "SELECT ISNULL((SELECT SUM(ExpAmt) FROM ContractItem WHERE ContractId = A.ContractId), 0) ProdAmt, "&_
			"ISNULL((SELECT SUM(AddInSign * AddInValue) FROM ContractAddIn WHERE ContractId = A.ContractId), 0) AddInAmt, "&_
			"ISNULL((SELECT SUM(RecAmt) FROM AccountRecAmt WHERE ConId = A.ContractId), 0) RecAmt, "&_
			"A.ConCurr, A.ExRate, A.ContractNo "&_
			"FROM Contract A INNER JOIN Customer C ON A.CustCode = C.CustCode "&_
			"WHERE "& ChkSql("'ExpContract'", "A.ContractId", "5, 6") &" AND IsCredit=1 and  C.CustId = '"& CustId &"'", G_DBConn, 1, 1, 1
	do while not RS.EOF
		ProdAmt = CDBL(RS("ProdAmt"))
		RecConAmt = CDBL(RS("RecAmt"))
		if RS("ConCurr") <> "USD" then
			ProdAmt = ProdAmt * CDBL(RS("ExRate")) /  UsdExRate 
			RecConAmt = RecConAmt * CDBL(RS("ExRate")) / UsdExRate
		end if

		CustUsedCredit = CustUsedCredit + CDbl(ProdAmt) + CDbl(RS("AddInAmt")) - CDbl(RecConAmt)
		RS.MoveNext
	loop
	RS.Close 
	SET RS = NOTHING
end function


	Function CustUsedAmt (CustId)
		Set RS = Server.CreateObject("ADODB.Recordset")
		'ȡ����Ԫ����
		RS.Open "SELECT ExRate FROM ExRate WHERE Currency = 'USD'", G_DBConn, 0, 1, 1
		if not RS.EOF then
			UsdExRate =CDBL(RS("ExRate"))
		end if
		RS.Close 
		
		CustUsedAmt=0
		RS.Open "SELECT ISNULL((SELECT SUM(ExpPrice * Qty) FROM ContractItem WHERE ContractId = A.ContractId), 0) ProdAmt, "&_
				"ISNULL((SELECT SUM(AddInSign * AddInValue) FROM ContractAddIn WHERE ContractId = A.ContractId), 0) AddInAmt, "&_
				"ISNULL((SELECT SUM(RecAmt) FROM AccountRecAmt WHERE ContractNo = A.ContractNo), 0) RecAmt, "&_
				"A.ConCurr,A.ContractNo,F.ExRate "&_
				"FROM Contract A INNER JOIN Customer C ON A.CustCode = C.CustCode "&_
				"Join Exrate F on A.ConCurr=F.Currency "&_
				"WHERE "& ChkSql("'ExpContract'", "A.ContractId", "5, 6") &" AND C.CustCode = '"& CustId &"'", G_DBConn, 2, 3, 1

		do while not RS.EOF
			ProdAmt = CDBL(RS("ProdAmt"))
			RecConAmt = CDBL(RS("RecAmt"))
			if RS("ConCurr") <> "USD" then
				ProdAmt = ProdAmt * CDBL(RS("ExRate")) / UsdExRate
				AddInAmt = AddInAmt * CDBL(RS("ExRate")) / UsdExRate
				RecConAmt = RecConAmt * CDBL(RS("ExRate")) / UsdExRate
			end if
	
			CustUsedAmt = CustUsedAmt + CDbl(ProdAmt) + CDbl(RS("AddInAmt")) - CDbl(RecConAmt)
			RS.MoveNext
		loop
		RS.Close 
		RS.open "select ISNULL(SUM(ExpPrice * Qty),0) as CurrAmt FROM ContractItem WHERE ContractId = '"&ContractId&"'",G_DBConn,2,3,1
			if RS.eof = false then 
				CustUsedAmt = CustUsedAmt + CDBL(rs("CurrAmt"))
			End if 
		RS.close

		Set RS = NOTHING
	end function


	Function AllowSubmit(ContractId)
		Set RS = Server.CreateObject("ADODB.Recordset")
		'�жϺ�ͬ�Ƿ������ύ
		AllowSubmit=1
		RS.open "select C.Checkresult,C.FinalCFRDate,C.MaxAmt FROM  Customer C Join Contract A on A.CustCode=C.Custcode left Join Payment B On A.PayTerms=B.PaymentCode where A.ContractId = '"& ContractId  &"' and B.IsExam=1", G_DBConn, 0, 1, 1
			if RS.eof = false then 
				CheckResult=RS("Checkresult")
				if CheckResult<>6 then 				'û���������
					if CustUsedAmt(CustCode) > 5000 then 
						ErrMsg("�ͻ�û�б���׼���ţ��ۼƺ�ͬ�����5000����������չ���Ա��ϵ��")
					end if 
				else 
					if RS("FinalCFRDate") < date() then 
						ErrMsg("�ͻ����������ѵ���������չ���Ա��ϵ��")
					end if 
					if RS("MaxAmt") < CustUsedCredit(CustId)  then 
						ErrMsg("�ͻ����ö�Ȳ��㣬������չ���Ա��ϵ��")
					end if 
				end if
			end if 
		RS.close
		SET RS = NOTHING
	end function
'**********************************************************

	'�ж����ڸ�ʽ�Ƿ�Ϸ���
	function DateCheck(DateValue,MSG)
		if DateValue<>"" then 
			if IsDate(DateValue) then 
				DateCheck = Cdate(DateValue)
			else 
				ErrMsg("�������"&MSG&"���ڸ�ʽ�������顣")
			end if 
		else 
			DateCheck = Null
		end if 
	end function
'**********************************************************

	'�ж����������Ƿ�Ϸ�
	Function NumericCheck(NumValue,MSG)
		if NumValue<> "" then 
			if IsNumeric(NumValue) then 
				NumericCheck = NumValue
			else 
				ErrMsg("�������"&MSG&"���ָ�ʽ�������顣")
			end if 
		else 
			NumericCheck = 0
		end if 
	end Function

'================================================================================================================================
'	ֱ����SQL����д�ֵʱ
'================================================================================================================================

function Valid(Content)
	if IsNull(Content) then
		Valid = ""
	else
		Valid = Replace(Content, "'", "''", 1, -1, 1)
	end if
end function


'================================================================================================================================
'	�ж��Ƿ�Ϊ���в����ڵ���Ч��ֵ,�������¼�¼�Ĵ����ֶ�.	(Value:�����ֵ,Table:����,Field:�ֶ���)
'================================================================================================================================
function ValidateCode(Value, Table, Field)
	set RS = Server.CreateObject("ADODB.Recordset")

	RS.Open "SELECT * FROM "& Table &" WHERE "& Field &" = '"& Valid(Value) &"'", G_DBConn, 0, 1, 1	
	if not RS.EOF then
		ValidateCode = ""										'Codeֵ��Чʱ���ؿմ�
	else
		ValidateCode = Value
	end if
	RS.Close
	
	set RS = nothing
end function


'=================================================================================================================================
'	����Ԥ�㵥��š�
'=================================================================================================================================
	Function NewBudgetCode(CorpId)
		set RS = server.CreateObject("ADODB.RecordSet")
		dim CorpCode
		
		RS.Open "SELECT BillPrefix FROM CorpInfo WHERE CorpId = "& CorpId, g_DBConn, 0, 1, 1
		if not RS.EOF then
			CorpCode = RS("BillPrefix")
		end if
		RS.Close
		
		YY = right(Year(Date()),2)
		RS.open "select Max(Right(BudgetCode, 3)) as BudgetCode from Contract where BudgetCode like '"& YY & CorpCode &"%'",G_DBConn, 0, 1, 1
		if IsNull(RS("BudgetCode")) then
			AA="001" 		'AA ��˾��������ˮ��
		else 
			AA=int(RS("BudgetCode"))+1
			AA=Left("000",3-Len(AA))&AA
		end if 
		RS.close
		
		NewBudgetCode =   YY & CorpCode  & UserID & "YS" & AA
		
		set RS = Nothing
	End Function

'==================================================================================================================================
'	����ѯ����š�
'==================================================================================================================================
	Function NewOfferCode()
		set RS = server.CreateObject("ADODB.RecordSet")		
		YY = right(Year(Date()),2)
		RS.open "select Max(Right(BudgetCode, 3)) as BudgetCode from offer where BudgetCode like '"&YY&"%'",G_DBConn, 0, 1, 1
		if IsNull(RS("BudgetCode")) then
			AA="001" 		'AA ��˾��������ˮ��
		else 
			AA=int(RS("BudgetCode"))+1
			AA=Left("000",3-Len(AA))&AA
		end if 
		RS.close		
		NewOfferCode = YY & "YS" & UserID & AA		
		set RS = Nothing
	End Function
	'===========================================================
	' �����µı��ݵ�ѯ��(���۵�)���
	'               ����Ҫ���Ƶı��
	'===========================================================
Function NewOriginCode(OriginCode)
    if OriginCode<>"" then
		set RS = server.CreateObject("ADODB.RecordSet")		
		RS.open "select Max(Right(BudgetCode, 3)) as OfferCode from Offer where OriginConId is not null and ConCopyId is not null and BudgetCode like '"&OriginCode&"%'",G_DBConn,2,2,1
			if IsNull(RS("OfferCode")) then
				AA="001" 		'AA δ�б��ݹ��ı��
			else 
				AA=int(RS("OfferCode"))+1
				AA=Left("000",3-Len(AA))&AA
			end if 
		RS.close
		NewOriginCode = OriginCode & AA		
		set RS = Nothing
     end if
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------
'	���ɱ��������
'--------------------------------------------------------------------------------------------------------------------------------------------------
function GetBhdNo(BhdId)
	set RS = server.CreateObject("ADODB.RecordSet")
	dim CorpCode
	
	RS.Open "SELECT C.BillPrefix FROM ConShip A INNER JOIN Contract B ON A.ContractId = B.ContractId "&_
			"INNER JOIN CorpInfo C ON B.CorpId = C.CorpId "&_
			"WHERE A.BhdId = "& BhdId, g_DBConn, 0, 1, 1		
	if not RS.EOF then
		CorpCode = RS("BillPrefix")
	end if
	RS.Close
	RS.Open "SELECT RIGHT(A.ConShipNo, 2) Num, B.ContractNo, B.MultiShip FROM ConShip A INNER JOIN Contract B ON A.ContractId = B.ContractId "&_
			"WHERE A.BhdId = '"& BhdId &"'", g_DBConn, 0, 1, 1
	do while not RS.EOF
		if RS("MultiShip") = "1" then
			Str = JointStr(Str, Right(RS("ContractNo"), 3), "/")
		else
			Str = JointStr(Str, Right(RS("ContractNo"), 3) & Chr(64 + RS("Num")), "/")
		end if
		RS.MoveNext
	loop
	RS.Close
	
	GetBhdNo = Right(Year(Date), 2) & CorpCode & UserId & Str
	
	set RS = nothing	
end function


'--------------------------------------------------------------------------------------------------------------------------------------------------
'	Ԥ��ȷ�ϵ�״̬��
'--------------------------------------------------------------------------------------------------------------------------------------------------
	function InvSheetState(Values)
		select case Values
			case "0"  State="δ���"
			case "1"  State="�����"
			case "2"  State="�����˻�"
		end select
		InvSheetState = State
	end function


	function InvSheetDomState(DomIsOver, BuyAmtIsOver)
		if BuyAmtIsOver = 2 then
			InvSheetDomState = "�˻ػ���"
		else
			InvSheetDomState = InvSheetState(DomIsOver)
		end if
	end function

'==================================================================================================================================
'	�����ʽ�֧����š�
'==================================================================================================================================
	Function NewAccId(CorpId)
		set RS = server.CreateObject("ADODB.RecordSet")
		dim CorpCode
		
		RS.Open "SELECT BillPrefix FROM CorpInfo WHERE CorpId = "& CorpId, g_DBConn, 0, 1, 1
		if not RS.EOF then
			CorpCode = RS("BillPrefix")
		end if
		RS.Close

		YY = CorpCode & Right(Year(Date()),2)
		MM = Left("00",2-Len(Month(Date()))) & Month(Date())
		
		RS.open "select Max(Right(AccNo,4)) as CC from AccountFee where AccNo like '"& YY & MM &"%'",G_DBConn, 0, 1, 1
		if IsNull(RS("CC")) then 	
			CC = "0001"					'BB ְԱ��������ˮ��
		else 
			CC = Int(RS("CC"))+1
			CC = Left("0000",4-Len(CC))&CC
		end if 
		RS.close
		
		NewAccId = YY & MM & CC	'�������������ɣ�AA + BB + CC�� AA(YY)����λ��ţ�BB(MM)����λ�·ݣ�CC�ǵ�ǰ����ȫ��˾����ˮ����

		set RS = Nothing
	End Function


'--------------------------------------------------------------------------------------------------------------------------------------------
'	�ʽ�֧���е���ʾ
'--------------------------------------------------------------------------------------------------------------------------------------------
	function OvesrTop(code, BhdId, ConId, SupplierChs)
		Set RS=Server.CreateObject("ADODB.RecordSet")

		RMBDueAmt = 0
		RMBPayedAmt = 0
		
		'������Ʊ֧��ʱ
		if BhdId <> 0 then
			if Code = "115" then
				'֧������ʱ����Ԥ��ȷ���е�ǰ�������Ӧ��������Ƚ�
				RS.Open "SELECT BuyAmtIsOver, DomIsOver FROM InvoiceSheet WHERE BhdId = "& BhdId, g_DBConn, 0, 1, 1
				if not RS.EOF then
					BuyAmtIsOver = RS("BuyAmtIsOver")		'�Ƿ��ύ����
					DomIsOver = RS("DomIsOver")				'�Ƿ��ύ���
				end if
				RS.Close
				
				if DomIsOver = 1 then		'�ύ���ʱ�빤��Ԥ��ȷ��Ӧ����Ƚ�
					SqlStr = "SELECT ISNULL(SuppFactPayAmt, 0) DueAmt FROM GInvoiceI_SuppId A LEFT OUTER JOIN Supplier B ON A.SuppId = B.Id "&_
							 "INNER JOIN InvoiceSheet C ON A.InvId = C.InvId WHERE (C.BhdId = '"& BhdId &"') AND  "&_
							 "B.SupplierChs ='"& SupplierChs &"' "
					RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
					if not RSTemp.EOF then
						DueAmt = "RMB"& RSTemp("DueAmt")
						RMBDueAmt = CDbl(RSTemp("DueAmt"))
					end if
					RSTemp.Close
				else						'�ύ����ʱ�빤������Ƚ�
					SqlStr = "SELECT ISNULL(SuppBuyAmt, 0) DueAmt FROM GInvoiceI_SuppId A LEFT OUTER JOIN Supplier B ON A.SuppId = B.Id "&_
							 "INNER JOIN InvoiceSheet C ON A.InvId = C.InvId WHERE (C.BhdId = '"& BhdId &"') AND  "&_
							 "B.SupplierChs ='"& SupplierChs &"' "
					RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
					if not RSTemp.EOF then
						DueAmt = "RMB"& RSTemp("DueAmt")
						RMBDueAmt = CDbl(RSTemp("DueAmt"))
					end if
					RSTemp.Close
				end if
				
				SqlStr = "SELECT B.Currency, ISNULL(SUM(PayAmt), 0) AS TAmt, MIN(B.ExRate) ExRate "&_
						 "FROM AccountFeeItem A join AccountFee B on A.AccId = B.AccId "&_
						 "WHERE A.BhdId = '"& BhdId &"' AND B.ChargeCode='"& Code &"' And B.Summary='"& SupplierChs &"' "&_
						 "GROUP BY B.Currency"
				RSTemp.Open Sqlstr, g_DBConn, 0, 1, 1
				do while not RSTemp.EOF
					PayedAmt = ComputeAmt(PayedAmt, RSTemp("Currency") & RSTemp("TAmt"), "+", "Add")
					RMBPayedAmt = RMBPayedAmt + CDbl(RSTemp("TAmt")) * CDbl(RSTemp("ExRate"))
					RSTemp.moveNext
				loop
				RSTemp.Close
			else
				if Code = "123" or Code = "124" then
					'֧�����ڹ�˰�������ֵ˰ʱ����Ԥ��ȷ���н����бȽ�
					SqlStr = "SELECT 'RMB' Currenc, ISNULL(CASE WHEN "& Code &" = '123' THEN A.BhdDeclareExpCustom ELSE A.BhdDeclareExpTax END, 0) UnitFee, "&_
							 "1 ExRate "&_
							 "FROM VBhd A WHERE BhdId = "& BhdId
				else
					'֧���ǻ���ʱ����Ԥ��ȷ���и÷��ý����бȽ�
					SqlStr = "SELECT A.Currenc, ISNULL(Sum(unitfee), 0) as UnitFee, MIN(A.ExRate) ExRate "&_
							"FROM InvoiceFee A INNER JOIN InvoiceSheet B ON A.InvId = B.InvId "&_
							"WHERE B.BhdId = '"& BhdId &"' AND A.ChargeId = '"& Code &"' "&_
							"GROUP BY A.Currenc"
				end if
				RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
				do while not RSTemp.EOF
					DueAmt = ComputeAmt(DueAmt, RSTemp("Currenc") & RSTemp("UnitFee"), "+", "Add")		'Ӧ�����
					RMBDueAmt = RMBDueAmt + CDbl(RSTemp("UnitFee")) * CDbl(RSTemp("ExRate"))			'ת��ΪRMB�Ľ�����֧����Ӧ��������
					RSTemp.MoveNext
				loop
				RSTemp.Close

				SqlStr = "SELECT B.Currency, ISNULL(SUM(A.PayAmt), 0) AS TAmt, MIN(B.ExRate) ExRate "&_
						 "FROM AccountFeeItem A INNER join AccountFee B on A.AccId = B.AccId "&_
						 "WHERE A.BhdId = '"& BhdId &"' AND B.ChargeCode = '"& ChargeCode &"' "&_
						 "GROUP BY B.Currency"
				RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
				do while not RSTemp.EOF
					PayedAmt = ComputeAmt(PayedAmt, RSTemp("Currency") & RSTemp("TAmt"), "+", "Add")	'�Ѹ����
					RMBPayedAmt = RMBPayedAmt + CDbl(RSTemp("TAmt")) * CDbl(RSTemp("ExRate"))
					RSTemp.MoveNext
				loop
				RSTemp.Close
			end if 
			Expr3 = RMBPayedAmt - RMBDueAmt				'RMB���

			if Expr3 > 0 then
				ExtAmt = ComputeAmt(PayedAmt, DueAmt, "+", "Sub")
				
				response.write "<table align=center>"
				response.write "<tr align=center><td style=""color:red"" height=25>Ԥ��ȷ�Ͻ�"& DueAmt &"��ʵ�ʽ�"& PayedAmt &"��������"& ExtAmt &"</td></tr>"
				response.write "</table>"
			end if

		elseif ConId <> 0 then			'����֧ͬ��Ԥ����ʱ
			if Code <> "115" then
				'֧���ǻ���ʱ�����ͬԤ���еĽ��Ƚ�
				SqlStr = "select max(A.Currenc) as Currenc, IsNull(Sum(UnitFee),0) as UnitFee, MIN(A.ExRate) ExRate "&_
						 "from BudgetFees A join Contract B on A.ContractId = B.ContractId "&_
						 "where B.ContractId = '"& ConId &"' and ChargeId = '"& ChargeCode &"' "&_
						 "GROUP BY A.Currenc"
				RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
				do while not RSTemp.eof
					DueAmt = ComputeAmt(DueAmt, RSTemp("Currenc") & RSTemp("UnitFee"), "+", "Add")
					RMBDueAmt = RMBDueAmt + CDbl(RSTemp("UnitFee")) * CDbl(RSTemp("ExRate"))
					RSTemp.MoveNext
				loop
				RSTemp.Close
				
				RSTemp.open "SELECT IsNull(SUM(A.PayAmt), 0) AS PayAmt, B.Currency, MIN(B.ExRate) ExRate "&_
							"FROM AccountFeeItem A INNER Join AccountFee B On A.AccId = B.AccId "&_
							"where B.chargeCode = '"& ChargeCode &"' and A.ConId = '"& ConId &"' GROUP by Currency ",G_DBConn, 0, 1, 1
				do while not RSTemp.eof 
					PayedAmt = ComputeAmt(PayedAmt, RSTemp("Currency") & RSTemp("PayAmt"), "+", "Add")
					RMBPayedAmt = RMBPayedAmt + CDbl(RSTemp("PayAmt")) * CDbl(RSTemp("ExRate"))
					RSTemp.MoveNext
				loop
				RSTemp.close
			else							'֧������ʱ����������ͬ�µ����вɹ��ɹ����ϼƱȽ�
				Sqlstr = "SELECT SUM(B.RMBConAmt) RMBConAmt "&_
						 "FROM DomContract A INNER JOIN VDomContract B ON A.DomId = B.DomId "&_
						 "WHERE A.ContractId = '"& ConId &"' "
				RSTemp.Open SqlStr, g_DBConn, 0, 1, 1
				do while not RSTemp.eof
					DueAmt = "RMB" & RSTemp("RMBConAmt")
					RMBDueAmt = RMBDueAmt + CDbl(RSTemp("RMBConAmt"))
					RSTemp.movenext
				Loop
				RSTemp.close

				SqlStr = "SELECT B.Currency, ISNULL(SUM(PayAmt), 0) AS TAmt, MIN(B.ExRate) ExRate "&_
						 "FROM AccountFeeItem A join AccountFee B on A.AccId = B.AccId "&_
						 "WHERE A.ConId = '"& ConId &"' AND B.ChargeCode='"& Code &"' And B.Summary='"& SupplierChs &"' "&_
						 "GROUP BY B.Currency"
				RSTemp.Open Sqlstr, g_DBConn, 0, 1, 1
				do while not RSTemp.EOF
					PayedAmt = ComputeAmt(PayedAmt, RSTemp("Currency") & RSTemp("TAmt"), "+", "Add")
					RMBPayedAmt = RMBPayedAmt + CDbl(RSTemp("TAmt")) * CDbl(RSTemp("ExRate"))
					RSTemp.moveNext
				loop
				RSTemp.Close
			end if
				
			Expr3 = RMBPayedAmt - RMBDueAmt				'RMB���
			if Expr3 > 0 then
				ExtAmt = ComputeAmt(PayedAmt, DueAmt, "+", "Sub")
				
				response.write "<table align=center>"
				response.write "<tr align=center><td style=""color:red"" height=25>Ԥ���"& DueAmt &"��ʵ�ʽ�"& PayedAmt &"��������"& ExtAmt &"</td></tr>"
				response.write "</table>"
			end if
		end if
		
		
		Set RS=Nothing
	end function


'--------------------------------------------------------------------------------------------------------------------------------------------
'	��ʽ�����֣�����FormatNumber(XX,С��λ,С����ǰ�Ƿ���0,�����Ƿ�������,�Ƿ���ǧ��λ)
'--------------------------------------------------------------------------------------------------------------------------------------------
	function FMTNumber(NumValue)

		FMTNumber=formatNumber(Numvalue,2,-1,0,-1)
		
	end function
	

'===========================================
'��������ұ����ڵ�ǰҳ�ֱұ��ۼ�
'===========================================	
	function NewArray()
	
		Set RS=Server.CreateObject("ADODB.Recordset")
		
		RS.open "Select CURRENCY FROM EXRATE  ",g_dbconn, 1, 1, 1
			if RS.eof = false then 
				Redim AmtArray(RS.recordcount,2)
				stop
				Session("RSTempCount")=RS.recordcount
				RSTempCount=Session("RSTempCount")
				RS.movefirst
			end if 
			ii=1
			do while RS.eof = false 
				AmtArray(ii,1)=RS("CURRENCY")
				AmtArray(ii,2)=0
				ii=ii+1
			RS.movenext
			Loop
		RS.close
		
		NewArray=AmtArray
		Set RS=Nothing
		
	end function
	
'=====================================================
'	��ֵ������
'=====================================================	
	Function WriteArray(AmtArray,Curr,Amt)
		For ii=0 to RSTempCount		'�ۼƵ�ǰҳ��ĸ���
			if AmtArray(ii,1)=Curr then
				  AmtArray(ii,2) = AmtArray(ii,2) + CDBL(Amt)
			end if
		next 
		WriteArray=AmtArray
	end function
	
'=====================================================
'	��ȡ���飬��ϼ�
'=====================================================
	function ReadArray(AmtArray)
		JG=""
		CurTotalAmt=""
		For ii=0 to RSTempCount
			if AmtArray(ii,2)<>0 then
				if ii= 0 then 
					CurTotalAmt = AmtArray(ii,1)&FMTNumber(AmtArray(ii,2))
					JG=" + "
				else 
					CurTotalAmt = CurTotalAmt & JG & AmtArray(ii,1)&FMTNumber(AmtArray(ii,2))
					JG=" + "
				end if
			end if
		next
		
		if CurTotalAmt="" then 
			CurTotalAmt=0
		end if
		ReadArray=CurTotalAmt
	end function
	
'=====================================================
'	�������
'=====================================================
	Function CleanArray(AmtArray)
		For ii=0 to RSTempCount
			AmtArray(ii,2) = 0		'�������
		next 
		CleanArray=AmtArray
	end function
	
	RSTempCount=Session("RSTempCount")
	
	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'	���������������ñ�־ʱ�����ݸñ�־�ж��Ƿ���޸�һ����������
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
function CanEdit(EmpCode, IsNew)
	Set RS=Server.CreateObject("ADODB.Recordset")

	RS.Open "select * from EmployeeRole where EmpCode='"& UserId &"' and GroupCode='GENERAL MANAGER' ",G_DBConn, 0, 1, 1
	if RS.EOF = true and not IsNew then 
		CanEdit = "disabled"
	end if
	RS.Close	
	
	Set RS=Nothing
end function
'=============================����ʣ����,operdate���� ʱ�䣬bankname��������accountno�˻�
function BankSurplusAmt(bankname,accountno,operdate)
	OutAmt=0
	inAmt=0
	SurplusAmt=0
	Set RS=Server.CreateObject("ADODB.Recordset")
	Set RSOri=Server.CreateObject("ADODB.Recordset")
	RSOri.open "select A.OriAmt From AccountInfo A inner join DeclareBank B On A.BankId=B.DeclareId Where A.AccountNo='"&accountno&"' And B.DeclareBank='"&bankname&"' ",G_DBConn, 0, 1, 1	
		If Not RSOri.Eof Then
		OriAmt=CDBL(RSOri("OriAmt"))				'=======��������˺ŵĳ�ʼֵ
		Else
		OriAmt=0
		End If	
	RSOri.Close
	
	RS.Open "select * from VBankRecord Where bankname='"&bankname&"' And accountno='"&accountno&"' And operdate<='"&operdate&"' Order By OperDate",G_DBConn, 0, 1, 1
	'Response.Write(RS.source)
	TotalOutAmt=0
	TotalInAmt=0
	Do While Not RS.Eof
		OutAmt=CDBL(RS("outAmt"))
		inAmt=CDBL(RS("InAmt"))		 
		if OutAmt<>0 And not isnull(OutAmt) Then
			TotalOutAmt=TotalOutAmt+OutAmt		'=======֧������ܺ�
		ElseIf inAmt<>0 And not isnull(inAmt) Then
			TotalInAmt=TotalInAmt+inAmt			'=======�������ܺ�
		End If
	RS.MoveNext
	Loop
	RS.Close
	BankSurplusAmt=OriAmt-TotalOutAmt+TotalInAmt
	Set RS=Nothing
	Set RSOri=Nothing
end function
'=============================
'���к���
	Function coder(str)
		if not isnull(str) then   
			str=replace(str,"<","&lt;")   
			str=replace(str,">","&gt;")   
			str=replace(str,chr(34),"&quot;")   
			str=replace(str,"&","&amp;")   
			str=replace(str,chr(13),"<br>")   
			str=replace(str,chr(9),"&nbsp;   &nbsp;   ")   
			str=replace(str,chr(32),"&nbsp;")   
		end if
		coder=str   
	End Function
	Function Htmlcoder(str)
		if not isnull(str) then
			str=replace(str,"&lt;","<")   
			str=replace(str,"&gt;",">")   
			str=replace(str,"&quot;",chr(34))   
			str=replace(str,"&amp;","&")   
			str=replace(str,"<br>",chr(13))   
			str=replace(str,"&nbsp;   &nbsp;   ",chr(9))   
			str=replace(str,"&nbsp;",chr(32)) 
		end if  
    	Htmlcoder=str
	end function
'=============================
'��ѯ��Ա�Ƿ��ѱ�ѡ�й�
function IsSelect(strYear,FieldName,Code)
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	IsSelect="false"
	rsMain.open "select * from BegOfPer where Year(BOPYear)='"&strYear&"'",G_DBConn,1,1,1
	do while not rsMain.eof 
		BOPId=rsMain("BOPId")
		rsTemp.open "select "&FieldName&" as FN from BegOfPer where BOPId="&BOPId&"",G_DBConn,1,1,1
		if not rsTemp.eof then
			strCode=rsTemp("FN")
			if instr(strCode,Code)>0 then
				IsSelect="true"
			end if
		end if
		rsTemp.close
		rsMain.movenext
	loop
	rsMain.close
end function
'============================
'���ñ�������Ա��ɫ
Function GetColor(ExamineId)
	set rsColor=Server.CreateObject("ADODB.Recordset")
	rsColor.open "select GradeState from Examine where ExamineId="&ExamineId&"",G_DBConn,1,1,1
	if not rsColor.eof then
		GradeState=rsColor("GradeState")
		Select case GradeState 
			case "0" GetColor="#999999"					'δ����
			case "1" GetColor="#FF0000"					'������ δ���� 
			case "2" GetColor="#FFCC66"					'��ʼ����
			case "3" GetColor="#99FF99"					'�������
			case "4" GetColor="#000000"					'��ֹ����
		end Select
	end if
	rsColor.close
end Function
'============================
'���ҵ�һ������
Function SelFirst(ExaItemId,MarksType,Marks)
	set FirstCode=Server.CreateObject("ADODB.Recordset")
	FirstCode.open "select MinMarks,MaxMarks from ExaMark where ExaItemId="&ExaItemId&" and "&_
		"MaxMarks=(select max(MaxMarks) from ExaMark where ExaItemId="&ExaItemId&")",G_DBConn,1,1,1
	if not FirstCode.eof then
		MinMarks=FirstCode("MinMarks")
		MaxMarks=FirstCode("MaxMarks")
	end if
	FirstCode.close
	if MarksType="True" then
		strSql="select count(*) count from ExaEmpTab where ExaItemId="&ExaItemId&" and Marks="&MinMarks&""
	else
		strSql="select count(*) count from ExaEmpTab where ExaItemId="&ExaItemId&" "&_
			"and Marks between "&MinMarks&" and "&MaxMarks&""
	end if
	Num=0
	FirstCode.open strSql,G_DBConn,1,1,1
	if not FirstCode.eof then
		Num=FirstCode("count")
	end if
	if Num>0 then
		SelFirst="True"
	else
		SelFirst="False"
	end if	
end Function
'============================
'����������˵��ο����ܵ÷�
Function TotalScore(ExamineId)
	set TotSco=Server.CreateObject("ADODB.Recordset")
	TotSco.open "select EI.ExamineId,Sum(isnull(convert(decimal(10,2),round(Marks,2)),0)) Marks "&_
		"from ExamineItem EI "&_
		"left join (select ExaItemId,sum(isnull(convert(decimal(10,2),round(Weighing*Marks/100,2)),0)) as Marks "&_
		"from ExaEmpTab group by ExaItemId) EET on EET.ExaItemId=EI.ExaItemId where EI.ExamineId="&ExamineId&" "&_
		"group by EI.ExamineId ",G_DBConn,1,1,1
	if not TotSco.eof then
		TotalScore=TotSco("Marks")
	else
		TotalScore=0
	end if
	TotSco.close
end Function
'===========================
'����������˵���Ҫ�ص÷�
Function FactorScore(ExaItemId)
	set FactSco=Server.CreateObject("ADODB.Recordset")
	FactSco.open "select EI.ExamineId,isnull(convert(decimal(10,2),round(Marks,2)),0) Marks "&_
		"from ExamineItem EI "&_
		"left join (select ExaItemId,sum(isnull(convert(decimal(10,2),round(Weighing*Marks/100,2)),0)) as Marks "&_
		"from ExaEmpTab group by ExaItemId) EET on EET.ExaItemId=EI.ExaItemId where EI.ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
	if not FactSco.eof then
		FactorScore=FactSco("Marks")
	else
		FactorScore=0
	end if
	FactSco.close
end Function
'==========================
'��������˶Ա������������ܵ÷�
Function ExaTotalScore(ExamineId,EmpCode)
	set ExaTotSco=Server.CreateObject("ADODB.Recordset")
	ExaTotSco.open "select EI.ExamineId,EET.ExaEmpCode, "&_
		"sum(isnull(convert(decimal(10,2),round(Marks,2)),0)) Marks "&_
		"from ExamineItem EI "&_
		"left join (select ExaItemId,ExaEmpCode,isnull(convert(decimal(10,2),round(Weighing*Marks/100,2)),0) as Marks "&_
		"from ExaEmpTab) EET on EET.ExaItemId=EI.ExaItemId "&_
		"where EI.ExamineId="&ExamineId&" and EET.ExaEmpCode='"&EmpCode&"' group by EI.ExamineId,EET.ExaEmpCode ",G_DBConn,1,1,1
	if not ExaTotSco.eof then
		ExaTotalScore=ExaTotSco("Marks")
	else
		ExaTotalScore=0
	end if
	ExaTotSco.close
end Function
'==========================
'��������˶Ա�������ĳҪ�������÷�
Function ExaFactorScore(ExaItemId,EmpCode)
	set ExaFacSco=Server.CreateObject("ADODB.Recordset")
	ExaFacSco.open "select EI.ExamineId,EI.ExaItemId,EET.ExaEmpCode, "&_
		"isnull(convert(decimal(10,2),round(Marks,2)),0) Marks "&_
		"from ExamineItem EI "&_
		"left join (select ExaItemId,ExaEmpCode,isnull(convert(decimal(10,2),round(Weighing*Marks/100,2)),0) as Marks "&_
		"from ExaEmpTab) EET on EET.ExaItemId=EI.ExaItemId "&_
		"where EI.ExaItemId="&ExaItemId&" and EET.ExaEmpCode='"&EmpCode&"'",G_DBConn,1,1,1
	if not ExaFacSco.eof then
		ExaFactorScore=ExaFacSco("Marks")
	else
		ExaFactorScore=0
	end if
	ExaFacSco.close
end Function
'============================
'������������ڼ俼��ƽ����
Function AvgScore(ObjType,ObjCode,BOPId)
	set TotSco=Server.CreateObject("ADODB.Recordset")
	TotSco.open "select BOPId,avg(Marks) Marks from BegOfPerItem BI "&_
		"inner join (select EI.ExamineId,E.BOPItemId,E.GradeState, "&_
		"sum(isnull(convert(decimal(10,2),round(Marks,2)),0)) Marks "&_
		"from ExamineItem EI "&_
		"left join Examine E on E.ExamineId=EI.ExamineId "&_
		"left join (select ExaItemId,sum(isnull(convert(decimal(10,2),round(Weighing*Marks/100,2)),0)) as Marks "&_
		"from ExaEmpTab group by ExaItemId) EET on EET.ExaItemId=EI.ExaItemId  "&_
		"where E.ExaObjType="&ObjType&" and ExaObjCode='"&ObjCode&"' "&_
		"group by EI.ExamineId,E.BOPItemId,E.GradeState "&_
		") E on E.BOPItemId=BI.BOPItemId "&_
		"where BOPId="&BOPId&"  and E.GradeState=4 "&_
		"group by BOPId  ",G_DBConn,1,1,1
	if not TotSco.eof then
		AvgScore=TotSco("Marks")
	else
		AvgScore=0
	end if
	TotSco.close
end Function
'===========================
'���ҿ��Ը��Ƶ����ID
Function MaxId(ObjType,ObjCode,ExaPerId)
	set rsMaxId=Server.CreateObject("ADODB.Recordset")
	rsMaxId.open "select Max(ExamineId) ExamineId from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer BP on BP.BOPId=BI.BOPId "&_
		"where E.ExaObjType="&ObjType&" and ExaObjCode='"&ObjCode&"' and "&_
		"BP.ExaPerId="&ExaPerId&" and GradeState<>0 ",G_DBConn,1,1,1
	if not rsMaxId.eof then
		MaxId=rsMaxId("ExamineId")
	else
		MaxId=""
	end if
	rsMaxId.close	
end Function
'===========================
'����Ҫ��
Function CopyFactor(CurExamineId,ExamineId)
	set rsCory=Server.CreateObject("ADODB.Recordset")
	set rsItem=Server.CreateObject("ADODB.Recordset")
	set rsEmpTab=Server.CreateObject("ADODB.Recordset")
	set rsEmpItem=Server.CreateObject("ADODB.Recordset")
	rsItem.cursorlocation=3
	rsCory.open "select * from ExamineItem where ExamineId="&ExamineId&"",G_DBConn,1,1,1
	do while not rsCory.eof
		ExaItemId=rsCory("ExaItemId")
		rsItem.open "select * from ExamineItem ",G_DBConn,2,3,1
		'����Ҫ��
		rsItem.addnew
			rsItem("ExamineId")=CurExamineId
			rsItem("ExaProdId")=rsCory("ExaProdId")
			rsItem("ExaFactorId")=rsCory("ExaFactorId")
			rsItem("MarksType")=rsCory("MarksType")
			rsItem("IsRepeat")=rsCory("IsRepeat")
			rsItem("Weighing")=rsCory("Weighing")
			rsItem("SumEmpWeigh")=rsCory("SumEmpWeigh")
			rsItem("OrderNum")=rsCory("OrderNum")
		rsItem.update
		CurExaItemId=rsItem("ExaItemId")
		rsItem.close
		'���ƿ�����Ա
		rsEmpTab.open "select * from ExaEmpTab where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
		do while not rsEmpTab.eof 	
			rsEmpItem.open "select * from ExaEmpTab ",G_DBConn,2,3,1
			rsEmpItem.addnew
				rsEmpItem("ExaItemId")=CurExaItemId
				rsEmpItem("ExaEmpCode")=rsEmpTab("ExaEmpCode")
				rsEmpItem("Weighing")=rsEmpTab("Weighing")
			rsEmpItem.update
			rsEmpItem.close
			rsEmpTab.movenext
		loop
		rsEmpTab.close
		rsCory.movenext
	loop
	rsCory.close
end Function
'����״̬
'=============================
Function GetGradeState(GradeState)
	Select case GradeState 
		case "0" GetGradeState="δ����"					'δ����
		case "1" GetGradeState="������"					'������ δ���� 
		case "2" GetGradeState="��ʼ����"				'��ʼ����
		case "3" GetGradeState="�������"				'�������
		case "4" GetGradeState="��ֹ����"				'��ֹ����
	end Select
end Function
'����״̬
'=============================
Function GetState(EmpCode,ExamineId)
	set rsState=Server.CreateObject("ADODB.Recordset")
	rsState.open "select State from ExaEmpTab ET left join ExamineItem EI on (EI.ExaItemId=ET.ExaItemId) "&_
		"where ET.ExaEmpCode='"&EmpCode&"' and ExamineId="&ExamineId&"",G_DBConn,1,1,1
	if not rsState.eof then
		strState=rsState("State")
	end if
	rsState.close
	Select case strState 
		case "0" GetState="δ����"					'δ����
		case "4" GetState="���ύ"					'������ δ���� 
		case "2" GetState="�����˻�"				'��ʼ����
		case "3" GetState="���˻�"					'�������
		'case "4" GetState="��ȷ��"					'��ֹ����
	end Select
end Function
'=============================
'�ж��ַ���Ϊ��ʱ���ؿո�
'=============================
Function strCheck(strValue)
	if strValue="" or isnull(strValue) then
		strCheck="&nbsp;"
	else
		strCheck=strValue
	end if
end Function
'��������Ƿ�Ϸ�
'=============================
Function CheckGrade(ExaItemId,Marks)
	CheckGrade="False"
	set rsGrade=Server.CreateObject("ADODB.Recordset")
	rsGrade.open "select * from ExamineItem where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
	if not rsGrade.eof then
		ExamineId=rsGrade("ExamineId")
		MarksType=rsGrade("MarksType")
	end if
	rsGrade.close
	rsGrade.open "select * from ExaMark where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
	if MarksType="True" then
		do while not rsGrade.eof 
			MaxMarks=rsGrade("MaxMarks")
			if cdbl(Marks)=cdbl(MaxMarks) then
				CheckGrade="True"
			end if
			rsGrade.movenext
		loop
	end if
	if MarksType="False" then
		do while not rsGrade.eof 
			MaxMarks=rsGrade("MaxMarks")
			MinMarks=rsGrade("MinMarks")
			if (cdbl(Marks)>=cdbl(MaxMarks)) and (cdbl(Marks)<=cdbl(MinMarks)) then
				CheckGrade="True"
			end if
			rsGrade.movenext
		loop
	end if
	rsGrade.close
end Function
'��ѯ����������ֵ
'=============================
Function GetMaxOrder(ExamineId)
	maxNum=0
	set rsMaxOrder=Server.CreateObject("ADODB.Recordset")
	rsMaxOrder.open "select isnull(max(OrderNum),0) as maxNum from ExamineItem where ExamineId="&ExamineId&"",G_DBConn,1,1,1
	if not rsMaxOrder.eof then
		maxNum=rsMaxOrder("maxNum")
	end if
	rsMaxOrder.close
	GetMaxOrder=maxNum
end Function
</script>