<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BD"%>
<html>
<head>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include virtual="/secret/Func_Censor.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>经理提示页</title>
</head>
<%
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
%>
<body>
<Center>
    <h2>经 理 提 示 页</h2>
</Center>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="800" >
<tr>
  <td><span class="STYLE1">评分表设置提示:</span></td>
</tr>
<tr><td>
<table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
  <tr align="center" bgcolor="DDDDDD">
    <td scope="col">序号</td>
    <td scope="col">被考核期间</td>
    <td bgcolor="DDDDDD" scope="col">被考核公司</td>
    <td bgcolor="DDDDDD" scope="col">被考核部门</td>
    <td scope="col">被考核岗位</td>
    <td scope="col">被考核人员</td>
    <td scope="col">考核状态</td>
  </tr>
<%
	RowIndex=0
	ExaObjType=""
	ExaObjCode=""
	rsTemp.open "select BI.BOPIName,Year(BOPYear) as BOPYear,E.ExaObjType, "&_
		"E.ExaObjCode,E.GradeState,P.ExaPerName,E.ExamineId,E.LastDate from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer B on B.BOPId=BI.BOPId "&_
		"left join ExaPeriod P on P.ExaPerId=B.ExaPerId "&_
		"where 1=1 "&Query&" and (E.GradeState=0) and "&_
		"("&SeeEmp("E.ExaObjCode", "Sel")&" or "&IsManage()&" or E.DetailRight like '%"&userId&"%') "&_
		"order by E.GradeState,E.LastDate",G_DBConn,1,1,1
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
			strGroupName=""
			do while not rsMain.eof
				if strGroupName="" then
					strGroupName=strGroupName&rsMain("GroupName")
				else
					strGroupName=strGroupName&"<br>"&rsMain("GroupName")
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
    <td><%=strCheck(strGroupName)%></td>
    <td><%=EmpNameChs%>&nbsp;</td>
    <td><%=GetGradeState(GradeState)%></td>
  </tr>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
</table>
</td></tr>
<tr>
  <td><span class="STYLE1">评分表完毕提示:</span></td>
</tr>
<tr><td>
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
	RowIndex=0
	ExaObjType=""
	ExaObjCode=""
	rsTemp.open "select BI.BOPIName,Year(BOPYear) as BOPYear,E.ExaObjType, "&_
		"E.ExaObjCode,E.GradeState,P.ExaPerName,E.ExamineId,E.LastDate from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer B on B.BOPId=BI.BOPId "&_
		"left join ExaPeriod P on P.ExaPerId=B.ExaPerId "&_
		"where 1=1 "&Query&" and (E.GradeState=3) and "&_
		"("&SeeEmp("E.ExaObjCode", "Sel")&" or "&IsManage()&" or E.DetailRight like '%"&userId&"%') "&_
		"order by E.GradeState,E.LastDate",G_DBConn,1,1,1
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
			strGroupName=""
			do while not rsMain.eof
				if strGroupName="" then
					strGroupName=strGroupName&rsMain("GroupName")
				else
					strGroupName=strGroupName&"<br>"&rsMain("GroupName")
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
    <td><%=strCheck(strGroupName)%></td>
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
</td></tr><tr>
  <td><span class="STYLE1">评分表提示:</span></td>
</tr>
<tr><td>
<table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
  <tr align="center" bgcolor="DDDDDD">
	<td>序号</td>
	<td>被考核期间</td>
	<td>被考核公司</td>
	<td>被考核部门</td>
	<td>被考核岗位</td>
	<td>被考核人员</td>
	<td>考核人员</td>
	<td>评分</td>
	<td>评分终止日</td>
	<td>评分状态</td>
  </tr>
<%
	RowIndex=0
	rsTemp.open "select min(BI.BOPIName) BOPIName,min(Year(BOPYear)) as BOPYear,Min(E.ExaObjType) ExaObjType, "&_
		"min(E.ExaObjCode) ExaObjCode,min(E.GradeState) GradeState,min(P.ExaPerName) ExaPerName, "&_
		"E.ExamineId,min(E.OverDate) OverDate,ET.ExaEmpCode from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer B on B.BOPId=BI.BOPId "&_
		"left join ExaPeriod P on P.ExaPerId=B.ExaPerId "&_
		"left join ExamineItem EI on EI.ExamineId=E.ExamineId "&_
		"left join ExaEmpTab ET on (ET.ExaItemId=EI.ExaItemId) "&_
		"where 1=1 "&Query&" and (E.GradeState=2 or E.GradeState=3) "&_
		"and ET.state<>0 and ET.state<>3 and ET.state<>4 group by E.ExamineId,ET.ExaEmpCode ",G_DBConn,1,1,1	
	do while not rsTemp.eof
	RowIndex=RowIndex+1
		ExamineId=rsTemp("ExamineId")
		BOPIName=rsTemp("BOPIName")
		BOPYear=rsTemp("BOPYear")
		ExaPerName=rsTemp("ExaPerName")
		ExaObjType=rsTemp("ExaObjType")
		ExaObjCode=rsTemp("ExaObjCode")
		GradeState=rsTemp("GradeState")
		OverDate=rsTemp("OverDate")
		CurEmpCode=rsTemp("ExaEmpCode")
		DeptName=""
		CorpNameChs=""
		EmpNameChs=""
		GroupName=""
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
			strGroupName=""
			do while not rsMain.eof
				GroupName=rsMain("GroupName")
				if strGroupName="" then
					strGroupName=GroupName
				else
					strGroupName=strGroupName&"<br>"&GroupName
				end if
				rsMain.movenext	
			loop
			rsMain.close
	end select
	rsMain.open "select EmpNameChs from Employee where EmpCode='"&CurEmpCode&"'",G_DBConn,1,1,1
	if not rsMain.eof then
		ExaEmpNameChs=rsMain("EmpNameChs")
	end if
	rsMain.close
	rsMain.open "select EF.ExaFactorName from ExamineItem EI "&_
		"left join ExaEmpTab EE on EE.ExaItemId=EI.ExaItemId "&_
		"left join ExaFactor EF on EF.ExaFactorId=EI.ExaFactorId "&_
		"left join Examine E on E.ExamineId=EI.ExamineId "&_
		"where EE.ExaEmpCode='"&CurEmpCode&"' and E.ExamineId="&ExamineId&"",G_DBConn,1,1,1
	strExaFactorName=""
	do while not rsMain.eof 
		ExaFactorName=rsMain("ExaFactorName")
		if strExaFactorName="" then
			strExaFactorName=ExaFactorName
		else
			strExaFactorName=strExaFactorName&"<br>"&ExaFactorName
		end if
		rsMain.movenext
	loop
	rsMain.close
%>
  <tr align="center" bgcolor="#FFFFFF">
	<td><a href="../Examine/GradeEdit.asp?ExamineId=<%=ExamineId%>&ExaEmpCode=<%=CurEmpCode%>"><%=RowIndex%></a>&nbsp;</td>
	<td><%=BOPYear&BOPIName%>&nbsp;</td>
	<td><%=CorpNameChs%>&nbsp;</td>
	<td><%=DeptName%>&nbsp;</td>
	<td><%=GroupName%>&nbsp;</td>
	<td><%=EmpNameChs%>&nbsp;</td>
	<td><%=ExaEmpNameChs%>&nbsp;</td>
	<td><%=ExaTotalScore(ExamineId,CurEmpCode)%>&nbsp;</td>
	<td><%=OverDate%>&nbsp;</td>
	<td><%=GetState(CurEmpCode,ExamineId)%>&nbsp;</td>
  </tr>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
</table>
</td></tr></table>
</body>
</html>
