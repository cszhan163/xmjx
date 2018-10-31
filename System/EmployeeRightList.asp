<%@ LANGUAGE = VBScript %>
<%ModuleCode = "CB"%>
<!--#include file = "../secret/checkpwd.asp"-->
<html>
<head>
<title>职员模块权限</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../secret/style.css" type=text/css rel=stylesheet>
</head>
<%
	Submit = Request("Submit")
	EmpID = Request("EmpID")
	GroupId = Request("GroupId")
	IsView = Request("IsView")
	
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	Set RSItem = Server.CreateObject("ADODB.Recordset")
	Set RSTemp = Server.CreateObject("ADODB.Recordset")

	If Submit = "返 回" Then
		if GroupId <> "" then
			Response.Redirect "EmployeeGroupEdit.asp?GroupId="& GroupId &"&EmpId="& EmpId
		else
			Response.Redirect "employeeedit.asp?EmpID="& EmpID
		end if
		Response.End 
	End If

	'取得当前用户的管理员身份
	RSEmp.Open "SELECT IsAdmin FROM Employee WHERE EmpCode = '"& UserId &"'", G_DBConn, 0, 1, 1
	if not RSEmp.EOF then
		UserIsAdmin = RSEmp("IsAdmin")
	end if
	RSEmp.Close 

	'设置权限信息
	If Submit = "设 置" Then
		ModuleRight = Request("ModuleCode") &","
		ModuleRight = Replace(ModuleRight, " ", "")
		
		DenyModuleRight = Request("DenyModuleCode") &","
		DenyModuleRight = Replace(DenyModuleRight, " ", "")

		if GroupId <> "" then
			G_DBConn.Execute "UPDATE EmployeeGroup SET ModuleRight = '"& ModuleRight &"', DenyModuleRight = '"& DenyModuleRight &"' WHERE GroupID = "& GroupID
		else
			G_DBConn.Execute "UPDATE Employee SET ModuleRight = '"& ModuleRight &"', DenyModuleRight = '"& DenyModuleRight &"' WHERE EmpID = "& EmpID
		end if
	End If
	
	'读取用户权限信息
	if GroupId <> "" then
		RSEmp.Open "SELECT GroupName, ModuleRight, DenyModuleRight FROM EmployeeGroup WHERE GroupId = '"& GroupId &"'", G_DBConn, 0, 1, 1
		if not RSEmp.EOF then
			GroupName = RSEmp("GroupName")
			ModuleRight = RSEmp("ModuleRight")
			DenyModuleRight = RSEmp("DenyModuleRight")
		end if
		RSEmp.Close 
	else
		RSEmp.Open "SELECT EmpCode, EmpNameChs, IsAdmin, ModuleRight, DenyModuleRight FROM Employee WHERE EmpID = " & EmpID, G_DBConn, 0, 1, 1
		If Not RSEmp.EOF Then
			EmpCode = RSEmp("EmpCode")
			EmpNameChs = RSEmp("EmpNameChs")
			IsAdmin = RSEmp("IsAdmin")					'是否是系统管理员
			ModuleRight = RSEmp("ModuleRight")
			DenyModuleRight = RSEmp("DenyModuleRight")
		End If
		RSEmp.Close
		
		if IsAdmin = True then
			MainSql = " AND EXISTS(SELECT * FROM Sys_Module WHERE MenuPos = A.MenuPos AND IsAdmin = 1) "
			SubSql = " AND IsAdmin = 1"
		end if
	end if
		
	if GroupId <> "" then
		Title = "详细权限 -- "& GroupName
	else
		Title = "详细权限 -- "& EmpNameChs
	end if
%>
<body class="pagebody">
<form name="qform" method="post" action="EmployeeRightList.asp">
<table class="pagetable">
	<tr>
		<td class="pagetitle"><%=Title%></td>
	</tr>
	<tr align="right">
		<td>
			<%if UserIsAdmin = True and GroupId = "" then%>
			<input type="radio" id="View1" name="IsView" value="1" <%if IsView = "1" then%>checked<%end if%> onClick="qform.submit()"><label for="View1">有效设置</label>
			<input type="radio" id="View2" name="IsView" value="0" <%if IsView = "0" then%>checked<%end if%> onClick="qform.submit()"><label for="View2">用户设置</label>
			<%end if%>
			<%if UserIsAdmin = True then%><input type="submit" name="Submit" value="设 置" <%if IsView = "1" then%>Disabled<%end if%> style="margin-left:20px"><%end if%>
			<input type="submit" name="Submit" value="返 回">
			<input type="hidden" name="EmpID" value="<%=EmpID%>">
			<input type="hidden" name="GroupID" value="<%=GroupId%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			<table class="table" rules="rows" bordercolor="gray" bgcolor="white" style="font:13px">
<%	'读取主菜单
	RSEmp.Open "SELECT A.MenuPos, A.MenuName FROM Sys_Menu A WHERE A.MenuCode <> 'systemhelp' "& MainSql &_
			   "ORDER BY A.MenuPos ASC", G_DBConn, 0, 1, 1
	do while not RSEmp.EOF 
%>
				<tr class="listheader">
					<td align="left" colspan="6"><b>&nbsp;<%=RSEmp("MenuName")%></b></td>
				</tr>
<%		'读取主菜单下的所有模块,当前 ModuleName 为 M... 的是字菜单的父项,不对应模块,此处不显示
		'原外贸模块黑色，内贸模块蓝色
		i = 0
		RSItem.Open "SELECT A.ModuleCode, A.IsFixed, "&_
					"CASE WHEN EXISTS(SELECT * FROM SYS_Group_Module WHERE ModuleCode = A.ModuleCode AND GroupCode = 'Export') "&_
					"	  THEN 'Black' "&_
					"	  ELSE 'blue' "&_
					"END ItemColor, "&_
					"(CASE WHEN LEFT(ModuleName, 1) = 'S' THEN RIGHT(ModuleName, LEN(ModuleName) -1) ELSE ModuleName END) ModuleName "&_
					"FROM Sys_Module A "&_
					"WHERE A.MenuPos = '"& RSEmp("MenuPos") &"' AND LEFT(ModuleName, 1) <> 'M' "& SubSql &_
					"ORDER BY MenuItemPos ASC", G_DBConn, 0, 1, 1
		do while not RSItem.EOF
			'当前模块在用户的权限设置
			if Instr(1, ModuleRight, RSItem("ModuleCode") &",", 1) > 0 then
				Chked = "checked"
			else
				Chked = ""
			end if
			if Instr(1, DenyModuleRight, RSItem("ModuleCode") &",", 1) > 0 then
				DenyChked = "checked"
			else
				DenyChked = ""
			end if
			
			'当前模块在用户所在所有身份的权限设置
			RSTemp.Open "SELECT (CASE WHEN EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode AND R.EmpCode = '"& EmpCode &"' "&_
						"		WHERE G.DenyModuleRight LIKE '%"& RSItem("ModuleCode") &"%') THEN 'NO' "&_
						"		WHEN EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode AND R.EmpCode = '"& EmpCode &"' "&_
						"		WHERE G.ModuleRight LIKE '%"& RSItem("ModuleCode") &"%') THEN 'YES' END) GModuleRight", G_DBConn, 0, 1, 1
			if not RSTemp.EOF then
				GModuleRight = RSTemp("GModuleRight")
			end if
			RSTemp.Close 

			'根据两类权限设置当前显示方式
			if IsView = "1" then				'显示用户当前的有效权限
				if DenyChked <> "" or GModuleRight = "NO" then
					DenyChked = "checked"
					Chked = ""
				else
					if Chked <> "" or GModuleRight = "YES" then
						Chked = "checked"
						DenyChked = ""
					end if
				end if
			else								'显示用户的权限设置
				Hid = ""
				DisY = ""
				DisN = ""
				if IsAdmin = True and RSItem("IsFixed") = True then	'当显示管理员权限且模块使用权利不可更改时
					DisY = " Disabled "
					DisN = " Disabled "
					if Chked <> "" then				'使用<hidden>提交模块代码
						Hid = "<input type=""hidden"" name=""ModuleCode"" value="""& RSItem("ModuleCode") &""">"
					end if
					if DenyChked <> "" then
						Hid = "<input type=""hidden"" name=""DenyModuleCode"" value="""& RSItem("ModuleCode") &""">"
					end if
				else
					if GModuleRight = "YES" then
						DisY = " Disabled "
						Chked = "checked"
					end if
					if GModuleRight = "NO" then
						DisN = " Disabled "
						DenyChked = "checked"
					end if
				end if
			end if
			
			ModuleName = RSItem("ModuleName")

			if i mod 3 = 0 then
				Response.Write "<tr>"
			end if
			Response.Write "<td class=""header"" style=""border:none; text-align:center;color:"& RSItem("ItemColor") &""" nowrap>"& ModuleName &"</td>"&_
						   "<td style=""border:none"" nowrap>"&_
						   "<input type=""checkbox"" id="""& RSItem("ModuleCode") &""" name=""ModuleCode"" value="""& RSItem("ModuleCode") &""" "& Chked & DisY &" onclick=""ClkCheck(Deny"& RSItem("ModuleCode") &")"">"&_
						   "<label for="""& RSItem("ModuleCode") &""">允许</label> "&_
						   "<input type=""checkbox"" id=""Deny"& RSItem("ModuleCode") &""" name=""DenyModuleCode"" value="""& RSItem("ModuleCode") &""" "& DenyChked & DisN &" onclick=""ClkCheck("& RSItem("ModuleCode") &")"">"&_
						   "<label for=""Deny"& RSItem("ModuleCode") &""">拒绝</label>"& Hid &_
						   "</td>"
			if i mod 3 = 2 then
				Response.Write "</tr>"
			end if
			RSItem.MoveNext
			i = i + 1
		loop
		RSItem.Close 
		'当最后一行未满一行时,补满所需的表格.
		if i mod 3 <> 0 then
			for i = i mod 3 to 2
				Response.Write "<td class=""header"" style=""border:none""></td><td style=""border:none""></td>"
			next
			Response.Write "</tr>"
		end if
		RSEmp.MoveNext
	loop
	RSEmp.Close 
%>
			</table>
	  </td>
	</tr>
</table>
</form>
</body>
<script language="VBS">
sub ClkCheck(ele)
	if ele.disabled = False then
		ele.checked = false
	end if
end sub
</script>
</html>