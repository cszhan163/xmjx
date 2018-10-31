<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "**"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include virtual="/secret/Func_Censor.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>绩效考核列表</title>
</head>
<%
	SelYear=CurSelValue("SelYear")
	SelBOPId=CurSelValue("SelBOPId")
	SelBOPItemId=CurSelValue("SelBOPItemId")
	CorpCode=CurSelValue("CorpCode")
	DeptCode=CurSelValue("DeptCode")
	EmpCode=CurSelValue("EmpCode")
	GroupCode=CurSelValue("GroupCode")
	GState=CurSelValue("GState")
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
%>
<body>
<form method="post" action="ExamineList.asp" name="Form1">
  <Center>
    <h2><%if left(ModuleCode,1)="B" then%>人 员 绩 效 考 核 表<%Else%>本 人 考 核 结 果<%end if%></h2>
  </Center>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
<tr><td>期间：<select name="SelYear" onChange="Form1.submit()">
		<option value="">选择年份</option>
<%
	rsMain.open "select Year(BOPYear) as BOPYear from BegOfPer group by BOPYear",G_DBConn,1,1,1
	do while not rsMain.eof
		BOPYear=rsMain("BOPYear")
%>
		<option value="<%=BOPYear%>" <%if trim(SelYear)=trim(BOPYear) then response.Write("selected") end if%>><%=BOPYear%>年</option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	  </select> 
	  <select name="SelBOPId" onChange="Form1.submit()">
		<option value="">选择期间</option>
<%
	rsMain.open "select B.BOPId,E.ExaPerName from BegOfPer B left join ExaPeriod E on E.ExaPerId=B.ExaPerId "&_
		"where Year(BOPYear)='"&SelYear&"'",G_DBConn,1,1,1
	do while not rsMain.eof
		ExaPerName=rsMain("ExaPerName")
		BOPId=rsMain("BOPId")
%>
		<option value="<%=BOPId%>" <%if trim(SelBOPId)=trim(BOPId) then response.Write("selected") end if%>><%=ExaPerName%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	  </select>
	  <select name="SelBOPItemId" onChange="Form1.submit()">
	    <option value="">选择详细期间</option>
<%
	rsMain.open "select * from BegOfPerItem where BOPId='"&SelBOPId&"'",G_DBConn,1,1,1
	do while not rsMain.eof
		BOPItemId=rsMain("BOPItemId")
		BOPIName=rsMain("BOPIName")
%>
		<option value="<%=BOPItemId%>" <%if trim(SelBOPItemId)=trim(BOPItemId) then response.Write("selected") end if%> ><%=BOPIName%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	    </select>
		</td>
		</tr>
		<tr>
		<td>条件：<select name="CorpCode" onChange="Form1.submit()">
		  <option value="">选择公司</option>
<%
	rsMain.open "select * from CorpInfo ",G_DBConn,1,1,1
	do while not rsMain.eof
	CurCorpCode=rsMain("CorpCode")
%>
	<option value="<%=CurCorpCode%>" <%if trim(CurCorpCode)=trim(CorpCode) then response.Write("selected") end if %>><%=rsMain("CorpNameChs")%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
		</select>
		<select name="DeptCode" onChange="Form1.submit()">
		  <option value="">选择部门</option>
<%
	rsMain.open "select * from Dept where CorpId in (select CorpId from CorpInfo where CorpCode='"&CorpCode&"')",G_DBConn,1,1,1
	do while not rsMain.eof
	CurDeptCode=rsMain("DeptCode")
%>
	<option value="<%=CurDeptCode%>" <%if trim(CurDeptCode)=trim(DeptCode) then response.Write("selected") end if %>><%=rsMain("DeptName")%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
		</select>
		<select name="EmpCode" onChange="Form1.submit()">
		  <option value="">选择人员</option>
<%
	rsMain.open "select * from Employee where DeptCode='"&DeptCode&"' and IsDel=0 ",G_DBConn,1,1,1
	do while not rsMain.eof
	CurEmpCode=rsMain("EmpCode")
%>
	<option value="<%=CurEmpCode%>" <%if trim(CurEmpCode)=trim(EmpCode) then response.Write("selected") end if %>><%=rsMain("EmpNameChs")%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	<option value="-1" style="color:#0099FF">===已离职人员===</option>
<%
	rsMain.open "select * from Employee where DeptCode='"&DeptCode&"' and IsDel=1 ",G_DBConn,1,1,1
	do while not rsMain.eof
	CurEmpCode=rsMain("EmpCode")
%>
	<option value="<%=CurEmpCode%>" <%if trim(CurEmpCode)=trim(EmpCode) then response.Write("selected") end if %> style="color:#CCCCCC"><%=rsMain("EmpNameChs")%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
		</select>
		<select name="GroupCode" onChange="Form1.submit()">
		  <option value="">选择岗位</option>
<%
	rsMain.open "select * from EmployeeGroup ",G_DBConn,1,1,1
	do while not rsMain.eof
	CurGroupCode=rsMain("GroupCode")
%>
	<option value="<%=CurGroupCode%>" <%if trim(CurGroupCode)=trim(GroupCode) then response.Write("selected") end if %>><%=rsMain("GroupName")%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
		</select>
        <select name="GState" onChange="Form1.submit()">
		<option value="">选择评分状态</option>
		<option value="0" <%if GState="0" then response.Write("selected") end if%>>未设置</option>
		<option value="1" <%if GState="1" then response.Write("selected") end if%>>已设置</option>
		<option value="2" <%if GState="2" then response.Write("selected") end if%>>开始评分</option>
		<option value="3" <%if GState="3" then response.Write("selected") end if%>>评分完毕</option>
		<option value="4" <%if GState="4" then response.Write("selected") end if%>>终止评分</option>
        </select>
        <input type="submit" name="Submits" value="查询"></td></tr>
</table>
<table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
  <tr align="center" bgcolor="DDDDDD">
    <td scope="col">序号</td>
    <td scope="col">被考核期间</td>
    <td bgcolor="DDDDDD" scope="col">被考核公司</td>
    <td bgcolor="DDDDDD" scope="col">被考核部门</td>
    <td scope="col">被考核岗位</td>
    <td scope="col">被考核人员</td>
    <td scope="col">最终得分</td>
    <td scope="col">最终评分日期</td>
    <td scope="col">考核状态</td>
  </tr>
<%
	'查询条件
	Query=""
	if SelYear<>"" and not isnull(SelYear) then
		Query=Query & " and Year(B.BOPYear)='"&SelYear&"' "
	end if
	if SelBOPId<>"" and not isnull(SelBOPId) then
		Query=Query & " and B.BOPId='"&SelBOPId&"' "
	end if
	if SelBOPItemId<>"" and not isnull(SelBOPItemId) then
		Query=Query & " and BI.BOPItemId='"&SelBOPItemId&"' "
	end if
	if CorpCode<>"" and not isnull(CorpCode) then
		Query=Query & " and ((E.ExaObjType=1 and E.ExaObjCode='"&CorpCode&"') "&_
			"or (E.ExaObjType=2 and E.ExaObjCode in (select DeptCode from Dept "&_
			"where CorpId in (select CorpId from CorpInfo where CorpCode='"&CorpCode&"'))) "&_
			"or (E.ExaObjType=3 and E.ExaObjCode in (select EmpCode from Employee where DeptCode in "&_
			"(select DeptCode from Dept where CorpId in (select CorpId from CorpInfo where CorpCode='"&CorpCode&"')))))"
	end if
	if DeptCode<>"" and not isnull(DeptCode) then
		Query=Query & " and ((E.ExaObjType=2 and E.ExaObjCode='"&DeptCode&"') "&_
			"or (E.ExaObjType=3 and E.ExaObjCode in (select EmpCode from Employee where DeptCode='"&DeptCode&"'))) "
	end if
	if EmpCode<>"" and not isnull(EmpCode) then
		Query=Query & " and E.ExaObjType=3 and E.ExaObjCode='"&EmpCode&"'"
	end if
	if GroupCode<>"" and not isnull(GroupCode) then
		Query=Query & " and E.ExaObjType=3 and E.ExaObjCode in (select EmpCode from EmployeeRole "&_
			"where GroupCode='"&GroupCode&"')"
	end if
	if GState<>"" and not isnull(GState) then
		Query=Query & " and E.GradeState='"&GState&"'"
	end if
	if ModuleCode="AB" then
		Query=Query & " and E.GradeState=4 "
	end if
	RowIndex=0
	ExaObjType=""
	ExaObjCode=""
	rsTemp.open "select BI.BOPIName,Year(BOPYear) as BOPYear,E.ExaObjType, "&_
		"E.ExaObjCode,E.GradeState,P.ExaPerName,E.ExamineId,E.LastDate from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer B on B.BOPId=BI.BOPId "&_
		"left join ExaPeriod P on P.ExaPerId=B.ExaPerId "&_
		"where 1=1 "&Query&" and"&_
		"("&SeeEmp("E.ExaObjCode", "Sel")&" or "&IsManage()&" or E.DetailRight like '%"&userId&"%' or E.ExaObjCode=(select DeptCode from Employee where EmpCode='"&userId&"')) order by E.LastDate ",G_DBConn,1,1,1
	do while not rsTemp.eof
		RowIndex=RowIndex+1
		ExamineId=rsTemp("ExamineId")
		BOPIName=rsTemp("BOPIName")
		BOPYear=rsTemp("BOPYear")
		ExaPerName=rsTemp("ExaPerName")
		ExaObjType=rsTemp("ExaObjType")
		ExaObjCode=rsTemp("ExaObjCode")
		GradeState=rsTemp("GradeState")
		LastDate=rsTemp("LastDate")
		DeptName=""
		CorpNameChs=""
		EmpNameChs=""
		GroupName=""
		strGroupName=""
		select case ExaObjType
		case "1" 
			rsMain.open "select CorpNameChs from CorpInfo where CorpCode='"&ExaObjCode&"'",G_DBConn,1,1,1
			if not rsMain.eof then
				CorpNameChs=rsMain("CorpNameChs")
			end if
			rsMain.close
		case "2"
			rsMain.open "select D.DeptName,C.CorpNameChs from Dept D "&_
				"left join CorpInfo	C On (C.CorpId=D.CorpId) "&_
				"where DeptCode='"&ExaObjCode&"'",G_DBConn,1,1,1
			if not rsMain.eof then
				DeptName=rsMain("DeptName")
				CorpNameChs=rsMain("CorpNameChs")
			end if
			rsMain.close
		case "3"
			rsMain.open "select E.EmpCode,E.EmpNameChs,D.DeptName,C.CorpNameChs from Employee E "&_
				"left join Dept D on(D.DeptCode=E.DeptCode) "&_
				"left join CorpInfo C on(C.CorpId=D.CorpId) "&_
				"where EmpCode='"&ExaObjCode&"'",G_DBConn,1,1,1
			if not rsMain.eof then
				EmpCode=rsMain("EmpCode")
				EmpNameChs=rsMain("EmpNameChs")
				DeptName=rsMain("DeptName")
				CorpNameChs=rsMain("CorpNameChs")
			end if
			rsMain.close
			rsMain.open "select EG.GroupName from Employee E "&_
				"left join EmployeeRole ER on (ER.EmpCode=E.EmpCode) "&_
				"left join EmployeeGroup EG on (EG.GroupCode=ER.GroupCode) "&_
				"where E.EmpCode='"&EmpCode&"'",G_DBConn,1,1,1
			do while not rsMain.eof
				GroupName=rsMain("GroupName")
				if GroupName<>"" and not isnull(GroupName) then
					if strGroupName="" or isnull(strGroupName) then
						strGroupName=GroupName
					else
						strGroupName=strGroupName&"<br>"&GroupName
					end if
				end if
				rsMain.movenext	
			loop
			rsMain.close
	end select
%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="ExamineEdit.asp?ExamineId=<%=ExamineId%>"><%=RowIndex%></a>&nbsp;</td>
    <td><%=BOPYear&BOPIName%>&nbsp;</td>
    <td><%=CorpNameChs%>&nbsp;</td>
    <td><%=DeptName%>&nbsp;</td>
    <td nowrap><%=strCheck(strGroupName)%>&nbsp;</td>
    <td><%=EmpNameChs%>&nbsp;</td>
    <td><%=TotalScore(ExamineId)%></td>
    <td><%=Year(LastDate)&"-"&Month(LastDate)&"-"&Day(LastDate)%>&nbsp;</td>
    <td><%=GetGradeState(GradeState)%></td>
  </tr>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
</table>
</form>
</body>
</html>
