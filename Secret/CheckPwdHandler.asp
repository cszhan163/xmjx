<script language="VBS" runat="server">
'---------------------------------------------------------------------------------------------------------------------------------
'	ComputeAmt ����
'---------------------------------------------------------------------------------------------------------------------------------
function SplitAmt(Str, byVal Delimiter)
	set Reg = new RegExp
	dim aRt(), i, Amt
	
	if Str = "" then
		SplitAmt = aRt
		exit function
	end if
	
	NegativeDelimiter = Replace(Delimiter, "+", "-", 1, -1, 1)
	Delimiter = Replace(Delimiter, "+", "\x2B", 1, -1, 1)
	NegativeDelimiter = Replace(NegativeDelimiter, "-", "\x2D", 1, -1, 1)
	
	Reg.Global = true
	Reg.IgnoreCase = true

	'����ڸ��ݼӺŷֵõĽ�����д��ڸ������(RMB20-USD10)���Ҳ�Ϊ(RMB20+USD-10)ʱ����ִ���
	OneCurrAmt = "([a-z]{3})?(-?[\d,.]+)"
	Delimiter = Delimiter &"|"& NegativeDelimiter
	
	Reg.Pattern = "(("& Delimiter &")?"& OneCurrAmt &")"
	set Matches = Reg.Execute(Str)
	
	Redim aRt(Matches.Count - 1)

	for i = 0 to Matches.Count - 1
		set Amt = Matches(i)
		aRt(i) = Amt.SubMatches(2) & CCur(Amt.SubMatches(1) &"1") * CCur(Amt.SubMatches(3))
	next

	SplitAmt = aRt

	set Reg = nothing
end function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	ComputeAmt ����		���ݽ��ָ�����һ�����ıұ����ֵ������ȷ�Ľ����ʾ��ʽ(RMB10-USD10)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function MakeAmt(IsAddDelimiter, byVal Delimiter, CA, VA)
	absVA = Abs(VA)
	if VA >= 0 then
		Sign = ""
	else
		Sign = "-"
	end if

	if IsAddDelimiter then
		if Sign <> "" then
			if Delimiter <> "" then
				Delimiter = Replace(Delimiter, "+", Sign, 1, -1, 1)
			else
				Delimiter = Sign
			end if
		end if
		MakeAmt = Delimiter &" "& CA & FormatNumber(absVA, 2, -1)
	else
		MakeAmt = Delimiter & Sign &" " & CA & FormatNumber(absVA, 2, -1)
	end if
end function
</script>