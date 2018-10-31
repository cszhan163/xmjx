<%@ Language=VBScript %>
<%ModuleCode = "CB"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<HTML>
<HEAD>
<title>用户身份列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../secret/style.css" type=text/css rel=stylesheet>
</HEAD>
<%
	EmpId = Request("EmpId")
	GroupID = Request("GroupID")
	Submit = Request("Submit")

	If Submit = "返 回" Then
		Response.Redirect "EmployeeEdit.asp?EmpID="& EmpID
  		Response.End
	End If

	if Submit = "添 加" then
		Response.Redirect "EmployeeGroupEdit.asp?EmpId="& EmpId &"&GroupId=-2"
		Response.End 
	end if
	
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	
%>
<BODY class="pagebody">
<form name="qform" action="EmployeeGroupList.asp" method="post">
<table align="center" class="pagetable" style="width:700">
	<tr>
		<td class="pagetitle">用户身份列表</td>
	</tr>
	<tr>
		<td align="right">
			<%=Message%>
			<input type="submit" name="Submit" value="添 加">
			<input type="submit" name="Submit" value="返 回">
			<input type="hidden" name="EmpId" value="<%=EmpId%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			<table rules="all" bordercolor="gray" class="listtable">
				<tr class="listheader">
					<td nowrap>身份代码</td>
					<td nowrap>身份名称</td>
					<td nowrap>备注</td>
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