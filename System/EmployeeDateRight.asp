<%@ Language=VBScript %>
<%ModuleCode = "CB"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<HTML>
<HEAD>
<title>ְԱ����Ȩ�ޱ༭</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../secret/style.css" type=text/css rel=stylesheet> 
</HEAD>
<%
	Submit = Request("Submit")
	EmpID = Request("EmpID")
	GroupId = Request("GroupId")

	if Submit = "�� ��" then
		if GroupId <> "" then
			Response.Redirect "EmployeeGroupEdit.asp?GroupId="& GroupId &"&EmpId="& EmpId
		else
			Response.Redirect "EmployeeEdit.asp?GroupId="& GroupId &"&EmpId="& EmpId
		end if
		Response.End 
	end if

	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	Set RSItem = Server.CreateObject("ADODB.Recordset")
	
	'ȡ�õ�ǰ�û����������
	if GroupId <> "" then
		RSEmp.Open "SELECT GroupCode EGCode, GroupName UserName FROM EmployeeGroup WHERE GroupId = '"& GroupId &"'", G_DBConn, 0, 1, 1
	else
		RSEmp.Open "SELECT EmpCode EGCode, EmpNameChs UserName FROM Employee WHERE EmpId = '"& EmpId &"'", G_DBConn, 0, 1, 1
	end if
	if not RSEmp.EOF then
		EGCode = RSEmp("EGCode")
		UserName = RSEmp("UserName")
	end if
	RSEmp.Close 
	
	if Submit = "�� ��" then
		SelY = Request("SelY") &","
		SelY = Replace(SelY, " ", "")
		SelN = Request("SelN") &","
		SelN = Replace(SelN, " ", "")

		EdiY = Request("EdiY") &","
		EdiY = Replace(EdiY, " ", "")
		EdiN = Request("EdiN") &","
		EdiN = Replace(EdiN, " ", "")
		
		DelY = Request("DelY") &","
		DelY = Replace(DelY, " ", "")
		DelN = Request("DelN") &","
		DelN = Replace(DelN, " ", "")
		
		ChkY = Request("ChkY") &","
		ChkY = Replace(ChkY, " ", "")
		ChkN = Request("ChkN") &","
		ChkN = Replace(ChkN, " ", "")
		
		for each c in Request("CustSelY")
			if c = "All" then
				CustSelY = "All"
				exit for
			else
				CustSelY = c
			end if
		next
		for each c in Request("CustSelN")
			if c = "All" then
				CustSelN = "All"
				exit for
			else
				CustSelN = c
			end if
		next
		for each c in Request("CustEdiY")
			if c = "All" then
				CustEdiY = "All"
				exit for
			else
				CustEdiY = c
			end if
		next
		for each c in Request("CustEdiN")
			if c = "All" then
				CustEdiN = "All"
				exit for
			else
				CustEdiN = c
			end if
		next
		
		if GroupId <> "" then
			RSEmp.Open "SELECT * FROM Sys_DataRight WHERE GroupCode = '"& Valid(EGCode) &"'", G_DBConn, 1, 3, 1
			if RSEmp.EOF then
				RSEmp.AddNew
				RSEmp("GroupCode") = EGCode
			end if
		else
			RSEmp.Open "SELECT * FROM Sys_DataRight WHERE EmpCode = '"& Valid(EGCode) &"'", G_DBConn, 1, 3, 1
			if RSEmp.EOF then
				RSEmp.AddNew
				RSEmp("EmpCode") = EGCode
			end if
		end if
		RSEmp("SelY") = SelY
		RSEmp("SelN") = SelN
		RSEmp("EdiY") = EdiY
		RSEmp("EdiN") = EdiN
		RSEmp("DelY") = DelY
		RSEmp("DelN") = DelN
		RSEmp("ChkY") = ChkY
		RSEmp("ChkN") = ChkN
		RSEmp("CustSelY") = CustSelY
		RSEmp("CustSelN") = CustSelN
		RSEmp("CustEdiY") = CustEdiY
		RSEmp("CustEdiN") = CustEdiN
		RSEmp.Update 
		RSEmp.Close 
	end if
	
	if GroupId <> "" then
		RSEmp.Open "SELECT * FROM Sys_DataRight WHERE GroupCode = '"& Valid(EGCode) &"'", G_DBConn, 0, 1, 1
	else
		RSEmp.Open "SELECT * FROM Sys_DataRight WHERE EmpCode = '"& Valid(EGCode) &"'", G_DBConn, 0, 1, 1
	end if
	if not RSEmp.EOF then
		SelY = RSEmp("SelY")
		SelN = RSEmp("SelN")
		EdiY = RSEmp("EdiY")
		EdiN = RSEmp("EdiN")
		DelY = RSEmp("DelY")
		DelN = RSEmp("DelN")
		ChkY = RSEmp("ChkY")
		ChkN = RSEmp("ChkN")
		CustSelY = RSEmp("CustSelY")
		CustSelN = RSEmp("CustSelN")
		CustEdiY = RSEmp("CustEdiY")
		CustEdiN = RSEmp("CustEdiN")
	end if
	RSEmp.Close 
%>
<BODY class="pagebody">
<form name="qform" method="post" action="EmployeeDateRight.asp">
<table align="center" class="pagetable" style="width:700px">
	<tr>
		<td class="pagetitle">����Ȩ�� -- <%=UserName%></td>
	</tr>
	<tr>
		<td align="right">
			<input type="submit" name="Submit" value="�� ��">
			<input type="submit" name="Submit" value="�� ��">
			<input type="hidden" name="GroupId" value="<%=GroupId%>">
			<input type="hidden" name="EmpId" value="<%=EmpId%>">
		</td>
	</tr>
	<tr>
		<td>
			<table rules="all" bordercolor="gray" class="table" bgcolor="white" style="font:13px">
				<tr class="listheader">
					<td align="left" colspan="5"><b>&nbsp;�ͻ�</b></td>
				</tr>
				<tr class="header">
					<td></td>
					<td colspan="2">��ѯ</td>
					<td colspan="2">¼��</td>
				</tr>
				<tr align="center">
					<td class="header">����</td>
					<td colspan="2">
						<input type="checkbox" id="CustSelSelfY" name="CustSelY" value="Self" <%if CustSelY = "Self" or CustSelY = "All" then%>checked<%end if%> onClick="ClkCust(CustSelSelfY)"><label for="CustSelSelfY">����</label>
						<!--<input type="checkbox" id="CustSelSelfN" name="CustSelN" value="Self" <%if CustSelN = "Self" or CustSelN = "All" then%>checked<%end if%> onclick="ClkCust(CustSelSelfY)"><label for="CustSelSelfN">�ܾ�</label>-->
					</td>
					<td colspan="2">
						<input type="checkbox" id="CustEdiSelfY" name="CustEdiY" value="Self" <%if CustEdiY = "Self" or CustEdiY = "All" then%>checked<%end if%> onClick="ClkCust(CustEdiSelfY)"><label for="CustEdiSelfY">����</label>
						<!--<input type="checkbox" id="CustEdiSelfN" name="CustEdiN" value="Self" <%if CustEdiN = "Self" or CustEdiN = "All" then%>checked<%end if%> onclick="ClkCust(CustEdiSelfY)"><label for="CustEdiSelfN">�ܾ�</label>-->
					</td>
				</tr>
				<tr align="center">
					<td class="header">ȫ��</td>
					<td colspan="2">
						<input type="checkbox" id="CustSelAllY" name="CustSelY" value="All" <%if CustSelY = "All" then%>checked<%end if%> onClick="ClkCust(CustSelAllY)"><label for="CustSelAllY">����</label>
						<!--<input type="checkbox" id="CustSelAllN" name="CustSelN" value="All" <%if CustSelN = "All" then%>checked<%end if%> onclick="ClkCust(CustSelAllY)"><label for="CustSelAllN">�ܾ�</label>-->
					</td>
					<td colspan="2">
						<input type="checkbox" id="CustEdiAllY" name="CustEdiY" value="All" <%if CustEdiY = "All" then%>checked<%end if%> onClick="ClkCust(CustEdiAllY)"><label for="CustEdiAllY">����</label>
						<!--<input type="checkbox" id="CustEdiAllN" name="CustEdiN" value="All" <%if CustEdiN = "All" then%>checked<%end if%> onclick="ClkCust(CustEdiAllY)"><label for="CustEdiAllN">�ܾ�</label>-->
					</td>
				</tr>
				<tr class="listheader">
					<td colspan="5" align="left"><b>&nbsp;ҵ������</b></td>
				</tr>
				<tr class="header">
					<td>�û�����</td>
					<td>��ѯ</td>
					<td>�޸�</td>
					<td>ɾ��</td>
					<td>���</td>
				</tr>
<%	'��ʾ�������Ȩ��ʱ,�ṩ����ѡ��,ָ��ʹ�ô���ݵ��û�
	if GroupId <> "" then
		Sql = "UNION SELECT '_Self' EmpCode, '����' EmpNameChs ORDER BY EmpCode ASC "
	end if
	
	RSEmp.Open "SELECT EmpCode, EmpNameChs FROM Employee WHERE IsAdmin = 0 and IsDel=0 "& Sql, G_DBConn, 0, 1, 1
	do while not RSEmp.EOF
		if Instr(1, SelY, RSEmp("EmpCode"), 1) > 0 then
			chkSelY = " checked "
		else
			chkSelY = ""
		end if
		if Instr(1, SelN, RSEmp("EmpCode"), 1) > 0 then
			chkSelN = " checked "
		else
			chkSelN = ""
		end if
		if Instr(1, EdiY, RSEmp("EmpCode"), 1) > 0 then
			chkEdiY = " checked "
		else
			chkEdiY = ""
		end if
		if Instr(1, EdiN, RSEmp("EmpCode"), 1) > 0 then
			chkEdiN = " checked "
		else
			chkEdiN = ""
		end if
		if Instr(1, DelY, RSEmp("EmpCode"), 1) > 0 then
			chkDelY = " checked "
		else
			chkDelY = ""
		end if
		if Instr(1, DelN, RSEmp("EmpCode"), 1) > 0 then
			chkDelN = " checked "
		else
			chkDelN = ""
		end if
		if Instr(1, ChkY, RSEmp("EmpCode"), 1) > 0 then
			chkChkY = " checked "
		else
			chkChkY = ""
		end if
		if Instr(1, ChkN, RSEmp("EmpCode"), 1) > 0 then
			chkChkN = " checked "
		else
			chkChkN = ""
		end if
%>
				<tr align="center">
					<td class="header"><%=RSEmp("EmpNameChs")%></td>
					<td>
						<input type="checkbox" id="Sel<%=RSEmp("EmpCode")%>Y" name="SelY" value="<%=RSEmp("EmpCode")%>" <%=chkSelY%> onClick="ClkCheck(Sel<%=RSEmp("EmpCode")%>N)"><label for="Sel<%=RSEmp("EmpCode")%>Y">����</label>
						<input type="checkbox" id="Sel<%=RSEmp("EmpCode")%>N" name="SelN" value="<%=RSEmp("EmpCode")%>" <%=chkSelN%> onClick="ClkCheck(Sel<%=RSEmp("EmpCode")%>Y)"><label for="Sel<%=RSEmp("EmpCode")%>N">�ܾ�</label>
					</td>
					<td>
						<input type="checkbox" id="Edi<%=RSEmp("EmpCode")%>Y" name="EdiY" value="<%=RSEmp("EmpCode")%>" <%=chkEdiY%> onClick="ClkCheck(Edi<%=RSEmp("EmpCode")%>N)"><label for="Edi<%=RSEmp("EmpCode")%>Y">����</label>
						<input type="checkbox" id="Edi<%=RSEmp("EmpCode")%>N" name="EdiN" value="<%=RSEmp("EmpCode")%>" <%=chkEdiN%> onClick="ClkCheck(Edi<%=RSEmp("EmpCode")%>Y)"><label for="Edi<%=RSEmp("EmpCode")%>N">�ܾ�</label>
					</td>
					<td>
						<input type="checkbox" id="Del<%=RSEmp("EmpCode")%>Y" name="DelY" value="<%=RSEmp("EmpCode")%>" <%=chkDelY%> onClick="ClkCheck(Del<%=RSEmp("EmpCode")%>N)"><label for="Del<%=RSEmp("EmpCode")%>Y">����</label>
						<input type="checkbox" id="Del<%=RSEmp("EmpCode")%>N" name="DelN" value="<%=RSEmp("EmpCode")%>" <%=chkDelN%> onClick="ClkCheck(Del<%=RSEmp("EmpCode")%>Y)"><label for="Del<%=RSEmp("EmpCode")%>N">�ܾ�</label>
					</td>
					<td>
						<input type="checkbox" id="Chk<%=RSEmp("EmpCode")%>Y" name="ChkY" value="<%=RSEmp("EmpCode")%>" <%=chkChkY%> onClick="ClkCheck(Chk<%=RSEmp("EmpCode")%>N)"><label for="Chk<%=RSEmp("EmpCode")%>Y">����</label>
						<input type="checkbox" id="Chk<%=RSEmp("EmpCode")%>N" name="ChkN" value="<%=RSEmp("EmpCode")%>" <%=chkChkN%> onClick="ClkCheck(Chk<%=RSEmp("EmpCode")%>Y)"><label for="Chk<%=RSEmp("EmpCode")%>N">�ܾ�</label>
					</td>
				</tr>
<%
		RSEmp.MoveNext
	loop
	RSEmp.Close 
%>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
<script language="VBS">
sub ClkCheck(ele)
	if ele.disabled = False then
		ele.checked = false
	end if
end sub

sub ClkCust(ele)
	'ele.checked = false
	'alert(window.event.srcElement.id)

	'if Instr(1, ele.Id, "All", 1) > 0 then
	'	if Right(ele.Id, 1) = "Y" then
	'		name = Left(ele.Id, Len(ele.name) -1) &"N"
	'	else
	'		name = Left(ele.name, Len(ele.name) -1) &"Y"
	'	end if
	'	for each e in qform(name)
	'		e.checked = true
	'	next
	'end if
end sub
</script>
</HTML>
