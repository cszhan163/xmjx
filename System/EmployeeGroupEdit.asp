<%@ Language=VBScript %>
<%ModuleCode = "CB"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<HTML>
<HEAD>
<title>�û������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../secret/style.css" type=text/css rel=stylesheet>
</HEAD>
<%
	EmpId = Request("EmpId")
	GroupId = Request("GroupId")
	Submit = Request("Submit")
	
	if Submit = "�� ��" then
		Response.Redirect "EmployeeGroupList.asp?EmpId="& EmpId &"&GroupId="& GRoupId
		Response.End 
	end if
	
	if Submit = "��ϸȨ��" then
		Response.Redirect "EmployeeRightList.asp?GroupId="& GroupId &"&EmpId="& EmpId
		Response.End 
	end if

	if Submit = "����Ȩ��" then
		Response.Redirect "EmployeeDateRight.asp?GroupId="& GroupId &"&EmpId="& EmpId
		Response.End 
	end if
	
	if Submit = "ɾ ��" then
		G_DBConn.Execute "DELETE FROM EmployeeRole FROM EmployeeGroup A LEFT JOIN EmployeeRole R ON A.GroupCode = R.GroupCode "&_
					   "WHERE A.GroupId = '"& GroupId &"'; "&_
					   "DELETE FROM Sys_DataRight FROM Sys_DataRight A LEFT JOIN EmployeeGroup G ON A.GroupCode = G.GroupCode "&_
					   "WHERE G.GroupId = '"& GroupId &"'; "&_
					   "DELETE FROM EmployeeGroup WHERE GroupId = '"& GroupId &"'"
		Response.Redirect "EmployeeGroupList.asp?EmpId="& EmpId
	end if
	
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	
	if submit = "�� ��" then
		GroupName = Request("GroupName")
		Remark = Request("Remark")
		if GroupId = "-2" then
			GroupCode = Request("GroupCode")
			if GroupCode = "" then
				Message = "��¼����ݴ���!"
			else
				if ValidateCode(GroupCode, "EmployeeGroup", "GroupCode") = "" then
					Message = "����ݴ����Ѿ�����,������¼��!"
				else
					if GroupName = "" then
						Message = "��¼���������!"
					else
						if ValidateCode(GroupName, "EmployeeGroup", "GroupName") = "" then
							Message = "��� """& GroupName &""" �Ѿ�ʹ��,������¼��!"
						end if
					end if
				end if
			end if
			
			if Message = "" then
				G_DBConn.Execute "INSERT INTO EmployeeGroup(GroupCode, GroupName, Remark) "&_
							   "VALUES('"& Valid(GroupCode) &"', '"& Valid(GroupName) &"', '"& Valid(Remark) &"')"
				GroupId = GetNewId()
				
				EmpCodeNum = Request("EmpCode").Count
				for i = 1 to EmpCodeNum
					EmpCode = Valid(Request("EmpCode")(i))
					G_DBConn.Execute "INSERT INTO EmployeeRole(EmpCode, GroupCode) VALUES('"& EmpCode &"', '"& Valid(GroupCode) &"')"
				next
			end if
		else
			RSEmp.Open "SELECT * FROM EmployeeGroup WHERE GroupId = '"& GroupId &"'", G_DBConn, 1, 3, 1
			if not RSEmp.EOF then
				GroupCode = RSEmp("GroupCode")
				if GroupName = "" then
					Message = "���������Ϊ��!"
				else
					if ValidateHad(GroupName, GroupId, "EmployeeGroup", "GroupName", "GroupId") = "" then
						Message = "��� """& GroupName &""" �Ѿ�ʹ��,������¼��!"
					else
						RSEmp("GroupName") = GroupName
					end if
				end if
				RSEmp("Remark") = Remark
				RSEmp.Update 
			end if
			RSEmp.Close 
						
			RSEmp.Open "SELECT EmpCode FROM Employee WHERE IsAdmin = 0", G_DBConn, 0, 1, 1
			do while not RSEmp.EOF 
				for each ECode in Request("EmpCode")
					if ECode = RSEmp("EmpCode") then
						Finded = 1
						exit for
					end if
				next
				
				if Finded = 1 then
					Sql = "IF NOT EXISTS(SELECT * FROM EmployeeRole WHERE GroupCode = '"& Valid(GroupCode) &"' AND EmpCode = '"& Valid(RSEmp("EmpCode")) &"') "&_
						  "INSERT INTO EmployeeRole(EmpCode, GroupCode) VALUES('"& Valid(RSEmp("EmpCode")) &"', '"& Valid(GroupCode) &"')"
					G_DBConn.Execute Sql
				else
					G_DBConn.Execute "DELETE FROM EmployeeRole WHERE GroupCode = '"& Valid(GroupCode) &"' AND EmpCode = '"& Valid(RSEmp("EmpCode")) &"'"
				end if
				RSEmp.MoveNext
				Finded = 0
			loop
			RSEmp.Close 
		end if
	end if
	
	RSEmp.Open "SELECT * FROM EmployeeGroup WHERE GroupId = '"& GroupId &"'", G_DBConn, 0, 1, 1
	if not RSEmp.EOF then
		GroupCode = RSEmp("GroupCode")
		GroupName = RSEmp("GroupName")
		IsBuiltIn = RSEmp("IsBuiltIn")			'�Ƿ�Ϊϵͳ�������
		Remark = RSEmp("Remark")
	end if
	RSEmp.Close 
%>
<BODY class="pagebody">
<form name="qform" action="EmployeeGroupEdit.asp" method="post">
<table align="center" class="pagetable" style="width:700px">
	<tr>
		<td class="pagetitle">�û������Ϣ</td>
	</tr>
	<tr>
		<td align="right">
			<font color="red"><%=Message%></font>
			<%if GroupId <> "-2" then%>
				<input type="submit" name="Submit" value="��ϸȨ��">
				<input type="submit" name="Submit" value="����Ȩ��">
			<%end if%>
			<input type="submit" name="Submit" value="ɾ ��" onClick="Confirm()">
			<input type="submit" name="Submit" value="�� ��">
			<input type="submit" name="Submit" value="�� ��">
			<input type="hidden" name="EmpId" value="<%=EmpId%>">
			<input type="hidden" name="GroupId" value="<%=GroupId%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			<table rules="all" bordercolor="gray" class="table" bgcolor="white">
				<tr>
					<td class="header" nowrap>��ݴ���</td>
					<td nowrap>&nbsp;<%if GroupId = "-2" then%><input type="text" name="GroupCode" value="<%=GroupCode%>" class="input" maxlength="20"><%else%><%=GroupCode%><%end if%></td>
					<td class="header" nowrap>�������</td>
					<td nowrap>&nbsp;<input type="text" name="GroupName" value="<%=GroupName%>" class="input" maxlength="30"></td>
				</tr>
				<tr>
					<td nowrap class="header">��ע</td>
					<td colspan="3">&nbsp;<input type="text" name="Remark" value="<%=Remark%>" class="input" style="width:95%" maxlength="100"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
			<table rules="all" bordercolor="gray" class="listtable">
				<tr class="listheader">
					<td align="left" colspan="2"><b>&nbsp;�� Ա</b></td>
				</tr>
				<tr class="header">
					<td>�û�����</td>
					<td>�û�����</td>
				</tr>
<%
	if GroupId <> "-2" then
		RSEmp.Open "SELECT A.EmpId, A.EmpCode, A.EmpNameChs, R.GroupCode FROM Employee "&_
				   "A LEFT JOIN EmployeeRole R ON A.EmpCode = R.EmpCode AND R.GroupCode = '"& Valid(GroupCode) &"' "&_
				   "WHERE A.IsAdmin = 0 and IsDel=0 ORDER BY A.EmpCode ASC", G_DBConn, 0, 1, 1
		do while not RSEmp.EOF 
%>
				<tr>
					<td align="left">
						<input type="checkbox" Id="<%=RSEmp("EmpId")%>" name="EmpCode" value="<%=RSEmp("EmpCode")%>" <%if RSEmp("GroupCode") <> "" then%>checked<%end if%>>
						<label for="<%=RSEmp("EmpId")%>"><%=RSEmp("EmpCode")%></label>
					</td>
					<td><%=RSEmp("EmpNameChs")%></td>
				</tr>
<%
			RSEmp.MoveNext
		loop
		RSEmp.Close 
	end if
%>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
<script language="VBScript">
sub Confirm()
<%
	if IsBuiltIn = True then
%>
	Result = MsgBox("��� ""<%=GroupName%>"" Ϊϵͳ�������,���ܱ�ɾ��!", 16, "��ʾ")
	window.event.returnValue = false
<%
	else
%>
	Result = MsgBox("ɾ����� ""<%=GroupName%>"" ��?", 289, "ȷ��")
	if Result = 2 then
		window.event.returnValue = false
	end if
<%
	end if
%>
end sub
</script>
</HTML>
