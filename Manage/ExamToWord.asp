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
<title>绩效考核表编辑</title>
<script language="javascript">
function LeadOutWord(){
Layer1.style.border=0;
Wobj = new ActiveXObject('Word.Application');
Wobj.Application.Visible = true;
var mydoc=Wobj.Documents.Add('Normal',0,0);
Wobj.ActiveDocument.PageSetup
Wobj.ActiveDocument.PageSetup.TopMargin =35
Wobj.ActiveDocument.PageSetup.BottomMargin =35
Wobj.ActiveDocument.PageSetup.LeftMargin =40
Wobj.ActiveDocument.PageSetup.RightMargin =60
Wobj.ActiveDocument.PageSetup.Orientation=1
myRange =mydoc.Range(0,1);
var sel=Layer1.document.body.createTextRange();
sel.select();
Layer1.document.execCommand('Copy');
sel.moveEnd('character')
myRange.Paste();
var objView = mydoc.ActiveWindow.View
objView.Type = 3
mydoc.ActiveWindow.View.ShowParagraphs = false
mydoc.ActiveWindow.View.TableGridlines=false
//Wobj.ActiveDocument.Tables(1).Rows(1).Delete()
}
</script>
</head>
<%
	ExamineId=request("ExamineId")
	ExaFactorId=request("ExaFactorId")
	Submits=request("Submits")
	strSave=request("Save")
	
	Set rsMain = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	Set rsCorp = Server.CreateObject("ADODB.Recordset")

	'显示数据
	rsMain.open "select BI.BOPIName,Year(BOPYear) as BOPYear,E.ExaObjType,E.OverDate, "&_
		"E.ExaObjCode,E.GradeState,P.ExaPerName,E.DetailRight from Examine E "&_
		"left join BegOfPerItem BI on BI.BOPItemId=E.BOPItemId "&_
		"left join BegOfPer B on B.BOPId=BI.BOPId "&_
		"left join ExaPeriod P on P.ExaPerId=B.ExaPerId "&_
		"where E.ExamineId='"&ExamineId&"'",G_DBConn,1,1,1
	if not rsMain.eof then
		BOPIName=rsMain("BOPIName")
		BOPYear=rsMain("BOPYear")
		ExaPerName=rsMain("ExaPerName")
		ExaObjType=rsMain("ExaObjType")
		ExaObjCode=rsMain("ExaObjCode")
		GradeState=rsMain("GradeState")
		OverDate=rsMain("OverDate")
		DetailRight=rsMain("DetailRight")
	end if
	rsMain.close
	select case ExaObjType
		case "1" 
			rsMain.open "select CorpNameChs from CorpInfo where CorpCode='"&ExaObjCode&"'",G_DBConn,1,1,1
			if not rsMain.eof then
				CorpNameChs=rsMain("CorpNameChs")
				ObjName=CorpNameChs
			end if
			rsMain.close
		case "2"
			rsMain.open "select D.DeptName,C.CorpNameChs from Dept D "&_
				"left join CorpInfo	C On (C.CorpId=D.CorpId) "&_
				"where DeptCode='"&ExaObjCode&"'",G_DBConn,1,1,1
			if not rsMain.eof then
				DeptName=rsMain("DeptName")
				CorpNameChs=rsMain("CorpNameChs")
				ObjName=DeptName
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
				ObjName=EmpNameChs
			end if
			rsMain.close
			rsMain.open "select EG.GroupName from Employee E "&_
				"left join EmployeeRole ER on (ER.EmpCode=E.EmpCode) "&_
				"left join EmployeeGroup EG on (EG.GroupCode=ER.GroupCode) "&_
				"where E.EmpCode='"&EmpCode&"'",G_DBConn,1,1,1
			do while not rsMain.eof
				if GroupName<>"" then
					GroupName=GroupName&"<br>"&rsMain("GroupName")
				else
					GroupName=GroupName&rsMain("GroupName")
				end if
				rsMain.movenext	
			loop
			rsMain.close
	end select
	
	btnSave=""
	btnBack=""
	btnSaveEmp=""
	if ModuleCode="AB" then
		btnSave="disabled"
		btnBack="disabled"
		btnSaveEmp="disabled"
		btnDelete="disabled"
		btnSubmit="disabled"
	else
		if OverDate="" or isnull(OverDate) then
			btnSubmit="disabled"
		end if
		if GradeState<>1 then
			btnSubmit="disabled"
		end if
		if GradeState<>0 and  GradeState<>1 then
			btnSave="disabled"
		end if
		if GradeState<>3 then
			btnBack="disabled"
		end if
	end if
	
	Window_OffsetY=request("Window_OffsetY")
	if Window_OffsetY="" or isnull(Window_OffsetY) then Window_OffsetY=0
%>
<body onLoad="LeadOutWord()">
	<br>
  <Center>
    <h2><font color="#FF0000"><%=BOPYear%>年<%=BOPIName%>&nbsp;<%=ObjName%>&nbsp;</font>绩效考核表</h2>
  </Center>
  <br>
<table align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="95%" border="1" style="border-width:1px;">
<tr align="center">
	<td width="88" bgcolor="DDDDDD">考核期间</td>
	<td width="148" bgcolor="#FFFFFF"><%=BOPYear&BOPIName%>&nbsp;</td>
	<td width="74" bgcolor="DDDDDD">考核部门</td>
	<td width="201" bgcolor="#FFFFFF"><%=DeptName%>&nbsp;</td>
	<td width="127" bgcolor="DDDDDD">考核岗位</td>
	<td width="148" bgcolor="#FFFFFF"><%=GroupName%>&nbsp;</td>
</tr>
<tr align="center">
	<td bgcolor="DDDDDD">考核人员</td>
	<td bgcolor="#FFFFFF"><%=EmpNameChs%>&nbsp;</td>
	<td bgcolor="DDDDDD">终止日期</td>
	<td bgcolor="#FFFFFF"><%=OverDate%></td>
	<td bgcolor="DDDDDD">最终得分</td>
	<td bgcolor="#FFFFFF"><%=TotalScore(ExamineId)%>分&nbsp;</td>
</tr>
</table>
<%
	rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&" order by OrderNum",G_DBConn,1,1,1
	FactCount=0
	do while not rsMain.eof
		FactCount=FactCount+1
		ExaProdId=rsMain("ExaProdId")
		ExaFactorId=rsMain("ExaFactorId")
		Weighing=rsMain("Weighing")
		IsRepeat=rsMain("IsRepeat")
		MarksType=rsMain("MarksType")
		ExaItemId=rsMain("ExaItemId")
		
		isReadOnly=""
		if IsRepeat="" or isnull(IsRepeat) then
			isReadOnly="disabled"
		end if
		ScoreOnly=""
		if MarksType=true then
			ScoreOnly="disabled"
		end if
%>
<table align="center" cellpadding="0" cellspacing="0" width="95%" border="1" style="border-width:1px;">
<tr align="center" bgcolor="DDDDDD">
	<td nowrap>序号</td>
	<td bgcolor="DDDDDD">考核项目</td>
	<td >考核要素</td>
	<td bgcolor="DDDDDD">考核标准</td>
	<td >权重</td>
	<td>结果</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="3"><%=FactCount%></td>
	<td>
<%
	rsTemp.open "select * from ExaProdject where IsDel=0 and ExaProdId="&ExaProdId&" order by ExaProdName ",G_DBConn,1,1,1
	ExaProdName=rsTemp("ExaProdName")
%>
		<%=ExaProdName%>
<%
	rsTemp.close
%>	</td>
	<td>
<%

	rsTemp.open "select * from ExaFactor where ExaFactorId not in "&_
		"(select isnull(ExaFactorId,-1) from ExamineItem where ExaItemId<>"&ExaItemId&" and ExamineId="&ExamineId&" ) "&_
		"and IsDel=0 and ExaFactorId="&ExaFactorId&" order by ExaFactorName ",G_DBConn,1,1,1
	ExaFactorName=rsTemp("ExaFactorName")
%>
	<%=ExaFactorName%>
<%
	rsTemp.close
%></td>
	<td width="480" align="left">
	<%
		if ExaFactorId<>"" and not isnull(ExaFactorId) then
			rsTemp.open "select * from ExaFactor where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
			if not rsTemp.eof then
				response.Write(rsTemp("ExaNorm"))
			end if
			rsTemp.close
		end if
	%>	</td>
	<td>
	<%=Weighing%>%	</td>
	<td><%=FactorScore(ExaItemId)%>分</td>
</tr>
<tr align="center" bgcolor="DDDDDD">
	<td colspan="1" nowrap>可重复出现</td>
	<td bgcolor="DDDDDD">分值类型</td>
	<td colspan="3">评分办法及标准</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="1" align="center"><%if IsRepeat="True" then response.Write("是") else response.Write("否") end if%></td>
	<td align="center"><%if MarksType="True" then response.Write("数值") else response.Write("区间") end if%></td>
	<td width="480" colspan="3">
<%
	'评分办法及标准
	if MarksType<>"" and not isnull(MarksType) then
%>
		<table cellpadding="0" cellspacing="0" width="100%">
	<%
		rsTemp.open "select EM.*,EF.ExaFacItemName from ExaMark EM "&_
			"left join ExaFactorItem EF on(EF.ExaFacItemId=EM.ExaFacItemId) "&_
			"where ExaItemId="&ExaItemId&" ",G_DBConn,1,1,1
		Num=rsTemp.recordcount
		if not rsTemp.eof then
		if (Num mod 5)<>0 then
			Rows=cdbl(Num)/5+1 
		else
			Rows=cdbl(Num)/5
		end if
		for i=0 to Rows-1 
	%>
		  <tr align="center">
	<%
		for j=1 to 5
			if (i*5+j)<=Num then
			ExaFacItemId=rsTemp("ExaFacItemId")
			ExaFacItemName=rsTemp("ExaFacItemName")
			MinMarks=rsTemp("MinMarks")
			MaxMarks=rsTemp("MaxMarks")
	%>
			<td width="100"><%=ExaFacItemName%><br>
			<%if MarksType=0 then%><%=MaxMarks%>-<%=MinMarks%><%else%><%=MaxMarks%><%end if%></td>
	<%
			rsTemp.movenext
			end if
		next
	%>
		  </tr>
	<%
			
		next
		end if
		rsTemp.close
	%>
		</table>
<%
	end if
%>	</td>
</tr>
</table>
<%
		rsMain.movenext
	loop
	rsMain.close

%>
<div id="Layer1" ></div>
<meta http-equiv="refresh" content="0;url=ExamineEdit.asp?ExamineId=<%=ExamineId%>">
</body>
</html>
