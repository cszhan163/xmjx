<%@ Language=VBScript %>
<%ModuleCode = "CB"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<HTML>
<HEAD>
<title>�û�����б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../secret/style.css" type=text/css rel=stylesheet>
</HEAD>
<%
	EmpId = Request("EmpId")
	GroupID = Request("GroupID")
	Submit = Request("Submit")

	If Submit = "�� ��" Then
		Response.Redirect "EmployeeEdit.asp?EmpID="& EmpID
  		Response.End
	End If

	if Submit = "�� ��" then
		Response.Redirect "EmployeeGroupEdit.asp?EmpId="& EmpId &"&GroupId=-2"
		Response.End 
	end if
	
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	
%>
<BODY class="pagebody">
<form name="qform" action="EmployeeGroupList.asp" method="post">
<table align="center" class="pagetable" style="width:700">
	<tr>
		<td class="pagetitle">�û�����б�</td>
	</tr>
	<tr>
		<td align="right">
			<%=Message%>
			<input type="submit" name="Submit" value="�� ��">
			<input type="submit" name="Submit" value="�� ��">
			<input type="hidden" name="EmpId" value="<%=EmpId%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			<table rules="all" bordercolor="gray" class="listtable">
				<tr class="listheader">
					<td nowrap>��ݴ���</td>
					<td nowrap>�������</td>
					<td nowrap>��ע</td>
				</tr>
<%
	RSEmp.Open "SELECT * FROM EmployeeGroup ORDER BY GroupCode ASC", G_DBConn, 0, 1, 1
	do while not RSEmp.EOF 
%>
				<tr class="listitem">
					<td width="25%"><a href="EmployeeGroupEdit.asp?EmpId=<%=EmpId%>&GroupId=<%=RSEmp("GroupId")%>"><%=RSEmp("GroupCode")%></a></td>
					<td width="35%"><%=RSEmp("GroupName")%></td>
					<td align="left"><%=RSEmp("Remark")%></td>
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
</HTML>