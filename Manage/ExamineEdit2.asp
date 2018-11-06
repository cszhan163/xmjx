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
<script type="text/javascript" language="javascript">

var HoTooBbsLoad = 0;
var HoTooNavLoad = 0;

function loadBar(fl)
//fl is show/hide flag
{
  var x,y;
  if (self.innerHeight)
  {// all except Explorer
    x = self.innerWidth;
    y = self.innerHeight;
  }
  else 
  if (document.documentElement && document.documentElement.clientHeight)
  {// Explorer 6 Strict Mode
   x = document.documentElement.clientWidth;
   y = document.documentElement.clientHeight;
  }
  else
  if (document.body)
  {// other Explorers
   x = document.body.clientWidth;
   y = document.body.clientHeight; 
  }

    var el=document.getElementById('loader');
 if(null!=el)
 {
  var top = (y/2) - 50;
  var left = (x/2) - 150;
  if( left<=0 ) left = 10;
  el.style.visibility = (fl==1)?'visible':'hidden';
  el.style.display = (fl==1)?'block':'none';
  el.style.left = left + "px"
  el.style.top = top + "px";
  el.style.zIndex = 2;
 }
 HoTooBbsLoad = 1;
 HoTooNavLoad = 1;
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
	
	if Submits=" 返回 " then
		response.Redirect("ExamineList.asp")
		response.End()
	end if
	if Submits=" 打印 " then
		response.Redirect("ExamToWord.asp?ExamineId="&ExamineId&"")
		response.End()
	end if
	
	if Submits=" 删除 " then
		G_DBConn.execute("delete Examine where ExamineId="&ExamineId&"")
	end if
	
	if Submits=" 保存 " or strSave="ok" or Submits="添加考核要素" then
		rsMain.open "select GradeState from Examine where ExamineId="&ExamineId&"",G_DBConn,1,1,1
		if not rsMain.eof then
			GradeState=rsMain("GradeState")
		end if
		rsMain.close
		if (GradeState=0 or GradeState=1) and ModuleCode<>"AB" then 				'开始评分后不能在进行修改
			'删除考核办法
			varExaItemId=request("varExaItemId")
			if varExaItemId<>"" and not isnull(varExaItemId) then
				'response.Write(varExaItemId)
				G_DBConn.execute("update ExamineItem set MarksType=1,IsRepeat=1 "&_
					"where ExaItemId="&varExaItemId&"")
				G_DBConn.execute("delete ExaMark where ExaItemId="&varExaItemId&"")
			end if
			OverDate=request("OverDate")
			'保存考核表
			if OverDate<>"" and not isnull(OverDate) then
				if isdate(OverDate) then
					rsMain.open "select * from Examine where ExamineId="&ExamineId&"",G_DBConn,2,3,1
						rsMain("OverDate")=OverDate
						DetailRight=request("EmpScope"&ExamineId)
						if DetailRight<>"" and not isnull(DetailRight) then
							rsMain("DetailRight")=DetailRight
						end if
					rsMain.update
					rsMain.close
				else
					response.Write("<script language='javascript'>alert('日期填写错误！')</script>")
				end if
			end if
			'保存要素
			AllWeigh=0
			Rows=request("AllExaItemId").count
			for i=1 to Rows
				CurExaItemId=request("AllExaItemId")(i)
				rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&" and ExaItemId="&CurExaItemId&"",G_DBConn,2,3,1
					ExaProdId=request("ExaProdId"&CurExaItemId)
					if ExaProdId<>"" and not isnull(ExaProdId) then
						rsMain("ExaProdId")=ExaProdId
					end if
					ExaFactorId=request("ExaFactorId"&CurExaItemId)
					if ExaFactorId<>"" and not isnull(ExaFactorId) then				
						rsMain("ExaFactorId")=ExaFactorId
					end if
					MarksType=request("MarksType"&CurExaItemId)
					if MarksType<>"" and not isnull(MarksType) then
						rsMain("MarksType")=MarksType
					end if
					IsRepeat=request("IsRepeat"&CurExaItemId)
					if IsRepeat<>"" and not isnull(IsRepeat) then
						rsMain("IsRepeat")=IsRepeat
					end if
					Weighing=request("Weighing"&CurExaItemId)
					if Weighing<>"" and not isnull(Weighing) then
						rsMain("Weighing")=Weighing
						AllWeigh=cdbl(AllWeigh)+cdbl(Weighing)
					end if
				rsMain.update
				rsMain.close
				'保存考核办法及标准
				rsMain.open "select * from ExaMark where ExaItemId="&CurExaItemId&"",G_DBConn,1,1,1
					Rows=rsMain.recordcount
				rsMain.close
				if ExaFactorId<>"" and not isnull(ExaFactorId) then
					rsTemp.open "select * from ExaFactorItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
					do while not rsTemp.eof
						ExaFacItemId=rsTemp("ExaFacItemId")
						rsMain.open "select * from ExaMark where ExaItemId="&CurExaItemId&" and "&_
							"ExaFacItemId="&ExaFacItemId&"",G_DBConn,2,3,1
							if Rows=0 then 
								rsMain.addnew
								rsMain("ExaItemId")=CurExaItemId
								rsMain("ExaFacItemId")=ExaFacItemId
							end if
								MinMarks=request("Min"&CurExaItemId&"Marks"&ExaFacItemId)
								if MinMarks<>"" and not isnull(MinMarks) and IsNumeric(MinMarks) then
									rsMain("MinMarks")=MinMarks
								else
									rsMain("MinMarks")=0
								end if
								MaxMarks=request("Max"&CurExaItemId&"Marks"&ExaFacItemId)
								if MaxMarks<>"" and not isnull(MaxMarks) and IsNumeric(MaxMarks) then
									rsMain("MaxMarks")=MaxMarks
								else
									rsMain("MaxMarks")=0
								end if
						rsMain.update
						rsMain.close
						rsTemp.movenext
					loop
					rsTemp.close
				end if
			next
			if AllWeigh<>100 then
				strErr="请注意！要素权重相加不为100%。"
			else
			'更新考核状态
				G_DBConn.execute("update Examine set GradeState=1 where ExamineId="&ExamineId&"")
			end if
		end if
	end if
	
	if Submits="保存人员" then
		'保存考核表
		rsMain.open "select * from Examine where ExamineId="&ExamineId&"",G_DBConn,2,3,1
			DetailRight=request("EmpScope"&ExamineId)
			if DetailRight<>"" and not isnull(DetailRight) then
				rsMain("DetailRight")=DetailRight
			end if
		rsMain.update
		rsMain.close
	end if
	
	if Submits="添加考核要素" then
	rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&"",G_DBConn,2,3,1
	rsMain.addnew
		rsMain("ExamineId")=ExamineId
		rsMain("OrderNum")=cdbl(GetMaxOrder(ExamineId))+1
	rsMain.update
	rsMain.close
	end if
	
	if Submits="删除考核要素" then
		Rows=request("ExaItemId").count
		for i=1 to Rows
			CurExaItemId=request("ExaItemId")(i)
			G_DBConn.execute("delete ExamineItem where ExaItemId="&CurExaItemId&"")
		next
		'重新对数据编辑排序序号
		OrderIndex=0
		rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&" order by OrderNum ",G_DBConn,1,1,1
		do while not rsMain.eof
			OrderIndex=cdbl(OrderIndex)+1
			CurExaItemId=rsMain("ExaItemId")
			G_DBConn.execute("update ExamineItem set OrderNum="&OrderIndex&" where ExaItemId="&CurExaItemId&"")
			rsMain.movenext
		loop
		rsMain.close
	end if
	
	if Submits="开始评分" then
		rsMain.open "select sum(Weighing) as AllWeigh from ExamineItem where ExamineId="&ExamineId&" ",G_DBConn,1,1,1
		if not rsMain.eof then
			AllWeigh=rsMain("AllWeigh")
		end if
		rsMain.close
		rsMain.open "select count(*) as count from ExamineItem where ExamineId="&ExamineId&" and SumEmpWeigh=0 ",G_DBConn,1,1,1
		if not rsMain.eof then
			WeighCoun=rsMain("count")
		end if
		rsMain.close
		if cdbl(AllWeigh)<>100 then
			response.Write("<script language='javascript'>alert('要素权重相加不为100%，不能开始评分！')</script>")
		elseif WeighCoun<>0 then
			response.Write("<script language='javascript'>alert('人员权重相加不全为100%，不能开始评分！')</script>")
		else
			G_DBConn.execute("update Examine set GradeState=2 where ExamineId="&ExamineId&"")
		end if
	end if
	
	if Submits="终止评分" then
		rsMain.open "select EET.* from ExaEmpTab EET "&_
			"left join ExamineItem EI on EI.ExaItemId=EET.ExaItemId "&_
			"where EI.ExamineId='"&ExamineId&"'",G_DBConn,2,3,1
		do while not rsMain.eof
			Marks=rsMain("Marks")
			Weighing=rsMain("Weighing")
			rsMain("ResultMarks")=formatnumber(cdbl(Marks)*cdbl(Weighing)/100,2,-1)
			rsMain.update
			rsMain.movenext
		loop
		rsMain.close
		rsMain.open "select * from ExamineItem EI "&_
			"where EI.ExamineId='"&ExamineId&"'",G_DBConn,2,3,1
		do while not rsMain.eof 
			ExaItemId=rsMain("ExaItemId")
			rsMain("ResultMarks")=FactorScore(ExaItemId)
			rsMain.update
			rsMain.movenext
		loop
		rsMain.close
		G_DBConn.execute("update Examine set GradeState=4 where ExamineId='"&ExamineId&"'")
	end if
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
				GroupName=GroupName&"<br>"&rsMain("GroupName")
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
<body onLoad="window.scrollTo(0,<%=Window_OffsetY%>)">
<form method="post" action="ExamineEdit.asp?ExamineId=<%=ExamineId%>" name="Form1">
  <Center>
    <h2><font color="#FF0000"><%=BOPYear%>年<%=BOPIName%>&nbsp;<%=ObjName%>&nbsp;</font>绩效考核表</h2>
  </Center>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr><td align="right"><font color="#FF0000"><%=strErr%></font>
    <%if left(ModuleCode,1)="B" then%>
	<input type="submit" name="Submits" value=" 保存 " <%=btnSave%>>
    <input type="submit" name="Submits" value="开始评分" <%=btnSubmit%>>
    <input type="submit" name="Submits" value="终止评分" <%=btnBack%>>
    <input type="submit" name="Submits" value=" 删除 " <%=btnDelete%> onClick="return confirm('是否确定删除此表？')">
	<%else%>
	<input type="submit" name="Submits" value=" 打印 ">
	<%end if%>
    <input type="submit" name="Submits" value=" 返回 ">
  </td>
  </tr>
</table>
<table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
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
	<td bgcolor="#FFFFFF"><input name="OverDate" type="text" class="input" value="<%=OverDate%>"></td>
	<td bgcolor="DDDDDD">最终得分</td>
	<td bgcolor="#FFFFFF"><%=TotalScore(ExamineId)%>分&nbsp;</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr><td align="right">
<DIV   id=BtnDiv   style="border:0   solid   #808080; width:500; height:  30; position:   absolute; z-index:4; left: 402px; top: 146px;"> 
	<%if left(ModuleCode,1)="B" then%>
    <input type="submit" name="Submits" value="排序" onClick="javascript:DoOrder(<%=ExamineId%>)" <%=btnSave%>>
    <input name="Submits" type="button" id="Submits" value="批量添加" <%=btnSave%> onClick="javascript:SelFactor(<%=ExamineId%>)">
    <input type="submit" name="Submits" value="添加考核要素" <%=btnSave%> onClick="FactSave()">
    <input type="submit" name="Submits" value="删除考核要素" onClick="DelFactor()" <%=btnSave%>>
	<input type="submit" name="Submits" value=" 保存 " <%=btnSave%> onClick="FactSave()">
	<input type="hidden" name="Save" value="">
	<%end if%>
</DIV>
  </td>
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
<table border="1" align="center" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
<tr align="center" bgcolor="DDDDDD">
	<td width="36">序号</td>
	<td width="118" bgcolor="DDDDDD">考核项目</td>
	<td width="133">考核要素</td>
	<td width="220" bgcolor="DDDDDD">考核标准</td>
	<td width="85">权重</td>
	<td width="95">结果</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="3"><input type="checkbox" name="ExaItemId" value="<%=ExaItemId%>">&nbsp;<%=FactCount%>)&nbsp;
	<input type="hidden" name="AllExaItemId" value="<%=ExaItemId%>"></td>
	<td><select name="ExaProdId<%=ExaItemId%>">
		<option value="">选择考核项目</option>
<%
	rsTemp.open "select * from ExaProdject where IsDel=0 order by ExaProdName ",G_DBConn,1,1,1
	do while not rsTemp.eof
	CurExaProdId=rsTemp("ExaProdId")
	ExaProdName=rsTemp("ExaProdName")
%>
		<option value="<%=CurExaProdId%>" <%if trim(CurExaProdId)=trim(ExaProdId) then response.Write("selected") end if%> ><%=ExaProdName%></option>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
	</select>
	</td>
	<td><select name="ExaFactorId<%=ExaItemId%>" onChange="GetNorm(<%=ExaItemId%>)" style="width:200">
		<option value="">选择考核要素</option>
<%

	rsTemp.open "select * from ExaFactor where ExaFactorId not in "&_
		"(select isnull(ExaFactorId,-1) from ExamineItem where ExaItemId<>"&ExaItemId&" and ExamineId="&ExamineId&" ) "&_
		"and IsDel=0 order by ExaFactorName ",G_DBConn,1,1,1
	do while not rsTemp.eof
	CurExaFactorId=rsTemp("ExaFactorId")
	ExaFactorName=rsTemp("ExaFactorName")
%>
		<option value="<%=CurExaFactorId%>" <%if trim(CurExaFactorId)=trim(ExaFactorId) then response.Write("selected") end if%> ><%=ExaFactorName%></option>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
	</select></td>
	<td width="270" align="left">
	<%
		if ExaFactorId<>"" and not isnull(ExaFactorId) then
			rsTemp.open "select * from ExaFactor where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
			if not rsTemp.eof then
				response.Write(rsTemp("ExaNorm"))
			end if
			rsTemp.close
		end if
	%>
	</td>
	<td>
	
	<input name="Weighing<%=ExaItemId%>" type="text" class="priceinput" value="<%=Weighing%>">%
	<select name="ScoreKindId<%=ExaItemId%>" onChange="GetMark(<%=ExaItemId%>)" <%Response.Write(ScoreOnly)%>>
		<option value="-1">清空</option>
<%
	rsTemp.open "select * from ScoreKind ",G_DBConn,1,1,1
	do while not rsTemp.eof
		ScoreKindId=rsTemp("ScoreKindId")
		SWeighing=rsTemp("Weighing")
%>
	<option value="<%=ScoreKindId%>"><%=SWeighing%>%</option>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
	</select></td>
	<td><%=FactorScore(ExaItemId)%>分</td>
</tr>
<tr align="center" bgcolor="DDDDDD">
	<td colspan="1">是否可重复出现</td>
	<td bgcolor="DDDDDD">分值类型</td>
	<td colspan="3">评分办法及标准</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="1" align="center"><input name="IsRepeat<%=ExaItemId%>" type="radio" value="1" <%if IsRepeat="True" then response.Write("checked") end if%> onClick="AgeSave()">是
	<input name="IsRepeat<%=ExaItemId%>" type="radio" value="0" <%if IsRepeat="False" then response.Write("checked") end if%> onClick="AgeSave()">否</td>
	<td align="center"><input name="MarksType<%=ExaItemId%>" type="radio" value="1" <%if MarksType="True" then response.Write("checked") end if%> <%=isReadOnly%> onClick="AgeSave()">数值
	<input name="MarksType<%=ExaItemId%>" type="radio" value="0" <%if MarksType="False" then response.Write("checked") end if%> <%=isReadOnly%> onClick="AgeSave()">区间</td>
	<td width="480" colspan="3">
<%
	'评分办法及标准
	if MarksType<>"" and not isnull(MarksType) then
%>
		<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolordark="#FFFFFF" bordercolorlight="#999999">
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
			<%if MarksType=0 then%><input name="Max<%=ExaItemId%>Marks<%=ExaFacItemId%>" type="text" class="scoreinput" value="<%=MaxMarks%>" onChange="CheckNum(<%=ExaItemId%>,<%=ExaFacItemId%>,'Max')">-<input name="Min<%=ExaItemId%>Marks<%=ExaFacItemId%>" type="text" class="scoreinput" value="<%=MinMarks%>" onChange="CheckNum(<%=ExaItemId%>,<%=ExaFacItemId%>,'Min')"><%else%><input name="Max<%=ExaItemId%>Marks<%=ExaFacItemId%>" type="text" class="priceinput" value="<%=MaxMarks%>" onChange="CheckNum(<%=ExaItemId%>,<%=ExaFacItemId%>,'Max')"><%end if%></td>
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
%>
	</td>
</tr>
<%
	if IsDetail(userId,ExamineId)="True" or ModuleCode="BC" then
%>
<tr align="center" bgcolor="DDDDDD">
  <td colspan="3" bgcolor="DDDDDD">考核人员<%if ModuleCode="BC" and GradeState=0 or GradeState=1 then%>(<a href="javascript:LoadWindowUpDown(<%=ExaItemId%>,<%=ExamineId%>)">设置</a>)<%end if%></td>
  <td>权重</td>
<td colspan="2">评分</td>
<%
	rsTemp.open "select EE.*,E.EmpNameChs from ExaEmpTab EE "&_
		"left join Employee E on (E.EmpCode=EE.ExaEmpCode) "&_
		"where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
	do while not rsTemp.eof	
	EmpNameChs=rsTemp("EmpNameChs")
	Weighing=rsTemp("Weighing")
	Marks=rsTemp("Marks")
%>
</tr>
<tr align="center" bgcolor="#FFFFFF">
<td colspan="3"><%=EmpNameChs%>&nbsp;</td>
<td><%=Weighing%>%&nbsp;</td>
<td colspan="2"><%=Marks%>分</td>
</tr>
<%
		rsTemp.movenext
	loop
	rsTemp.close
%>
</table>
<%
		end if
		rsMain.movenext
	loop
	rsMain.close

%>
<%if ModuleCode<>"AB" then%>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
<tr><td valign="bottom"><span class="STYLE1">抄送以下人员查看：</span></td>
<td align="right"><input name="Submits" type="submit" value="保存人员" <%=btnSaveEmp%> onClick="EmpSave()"></td>
</tr>
</table>
<%
	rsCorp.open "select * from CorpInfo ",G_DBConn,1,1,1
	do while not rsCorp.eof
	CorpId=rsCorp("CorpId")
	CorpNameChs=rsCorp("CorpNameChs")
%>
<table border="1" align="center" width="800" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="0" cellspacing="0" style=" margin-top:5px; margin-bottom:5px;" >
  <tr>
    <td colspan="2" align="center" bgcolor="#DDDDDD" height="25"><%=CorpNameChs%>&nbsp;</td>
  </tr>
<%
rsMain.Open "Select DeptCode,DeptName From Dept where CorpId="&CorpId&" order by DeptID",G_DBConn,2,3,1
	Do While Not rsMain.Eof 
%> 
  <tr class=tdcss>
    <td nowrap bgcolor="#DDDDDD" align="center" height="25"><%=rsMain("DeptName")%></td>
    <td nowrap valign="top" bgcolor="#FFFFFF" align="left">&nbsp; 
<%
		rsTemp.Open "Select Grade,EmpCode,EmpNameChs From Employee Where DeptCode='"&rsMain("DeptCode")&"' ",G_DBConn,2,3,1
			i=1
		  	Do While Not rsTemp.Eof
%>
        <input type="checkbox" name="EmpScope<%=ExamineId%>" value="<%=rsTemp("EmpCode")%>" <% if InStr(DetailRight,rsTemp("EmpCode")) then Response.Write "Checked" %>>
<%
					Response.Write rsTemp("EmpNameChs")
				 		if Len(rsTemp("EmpNameChs"))<6 then
				   			Response.Write Replace(space(5-Len(rsTemp("EmpNameChs")))," ","&nbsp;&nbsp;")
				 		end if
						i=i+1
						if i=7 then
							i=1
							Response.Write "<br>&nbsp;"
						end if
               rsTemp.MoveNext
		  	Loop
		rsTemp.Close
%>      </td>
  </tr>
<%
		rsMain.MoveNext
	Loop
rsMain.Close
%> 
</table>
<%
		rsCorp.movenext
	loop
	rsCorp.close
end if
%>
<input type="hidden" name="varExaItemId" value="">
<input type="hidden" id="Window_OffsetY" name="Window_OffsetY" value="0">
<input type="hidden" id="Window_OffsetX" name="Window_OffsetX" value="0">
</form>
<script language="vbscript">
	sub GetNorm(strExaItemId)
		if strExaItemId<>"" and not isnull(strExaItemId) then
			Form1.varExaItemId.value=strExaItemId
			Form1.Window_OffsetY.value=document.body.scrollTop
			Form1.Save.value="ok"
			Form1.submit()
		end if
	end sub
	sub AgeSave()
		Form1.Window_OffsetY.value=document.body.scrollTop
		Form1.Save.value="ok"
		Form1.submit()
	end sub
	sub FactSave()
		Form1.Window_OffsetY.value=document.body.scrollTop
	end sub
	sub DelFactor()
		Mess=msgbox("是否确定删除？",1,"询问？")
		if Mess="1" then
			Form1.Window_OffsetY.value=document.body.scrollTop
		else
			window.event.returnValue=false
		end if
	end sub
	sub EmpSave()
		Form1.Window_OffsetY.value=document.body.scrollTop
		Form1.submit()
	end sub
	sub CheckNum(ItemId,FacItemId,Sel)
		strName=Sel&ItemId&"Marks"&FacItemId
		ScodeValue=eval("Form1."&strName&".value")
		if not IsNumeric(ScodeValue) then
			alert("请填写正确的数字格式！")
		end if
	end sub
	sub GetMark(ExaItemId)
		Num=0
		strName="ScoreKindId"&ExaItemId
		ScoreKindId=eval("Form1."&strName&".value")
		select case ScoreKindId
		<%
			rsMain.open "select * from ScoreKind ",G_DBConn,1,1,1
			do while not rsMain.eof
				ScoreKindId=rsMain("ScoreKindId")
				strWeighing=rsMain("Weighing")
				Max1=rsMain("Max1")
				Min1=rsMain("Min1")
				Max2=rsMain("Max2")
				Min2=rsMain("Min2")
				Max3=rsMain("Max3")
				Min3=rsMain("Min3")
				Max4=rsMain("Max4")
				Min4=rsMain("Min4")
				Max5=rsMain("Max5")
				Min5=rsMain("Min5")
		%>
			case "<%=ScoreKindId%>"
				eval("Form1.Weighing"&ExaItemId&"").value=<%=strWeighing%>
		<%
				rsCorp.open "select * from ExaMark where ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")",G_DBConn,1,1,1
				do while not rsCorp.eof
					ExaFacItemId=rsCorp("ExaFacItemId")
					ExaItemId=rsCorp("ExaItemId")
		%>
				select case ExaItemId
				case "<%=ExaItemId%>"
					Num=Num+1
					ExaFacItemId="<%=ExaFacItemId%>"
					if Num=1 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Max1%>
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Min1%>
					elseif Num=2 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Max2%>
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Min2%>
					elseif Num=3 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Max3%>
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Min3%>
					elseif Num=4 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Max4%>
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Min4%>
					elseif Num=5 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Max5%>
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=<%=Min5%>
					end if
				end select
		<%
						rsCorp.movenext
					loop
					rsCorp.close
				rsMain.movenext
			loop
			rsMain.close
		%>
		case else
			eval("Form1.Weighing"&ExaItemId&"").value=0
		<%
				rsCorp.open "select * from ExaMark where ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")",G_DBConn,1,1,1
				do while not rsCorp.eof
					ExaFacItemId=rsCorp("ExaFacItemId")
					ExaItemId=rsCorp("ExaItemId")
		%>
				select case ExaItemId
				case "<%=ExaItemId%>"
					Num=Num+1
					ExaFacItemId="<%=ExaFacItemId%>"
					if Num=1 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
					elseif Num=2 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
					elseif Num=3 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
					elseif Num=4 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
					elseif Num=5 then
						eval("Form1.Max"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
						eval("Form1.Min"&ExaItemId&"Marks"&ExaFacItemId&"").value=0
					end if
				end select
		<%
						rsCorp.movenext
					loop
					rsCorp.close
		%>
		end select
	end sub
</script>
<script language="javascript">
function LoadWindowUpDown(idd,ExamineId)
{
   var IDID=idd
   URL="ExaEmpEdit.asp"+"?ExaItemId="+IDID+"&ExamineId="+ExamineId+"&ClientY="+document.body.scrollTop;
   window.showModalDialog(URL,window,"dialogWidth:800px;dialogHeight:500px;");
}
function SelFactor(ExamineId)
{
   URL="SelFactor.asp"+"?ExamineId="+ExamineId+"&ClientY="+document.body.scrollTop;
   window.showModalDialog(URL,window,"dialogWidth:800px;dialogHeight:500px;");
}
function DoOrder(ExamineId)
{
   URL="OrderEdit.asp"+"?ExamineId="+ExamineId+"&ClientY="+document.body.scrollTop;
   window.showModalDialog(URL,window,"dialogWidth:800px;dialogHeight:500px;");
}
</script>
<script   language="javascript">   
  <!--   
  //window.onload   =   resizeDiv;   
  window.onresize   =   resizeDiv;   
  window.onscroll   =   resizeDiv;   
  window.onerror   =   function(){}   
  function   resizeDiv()   
  {   
  var   obj=document.getElementById("BtnDiv")   
  try{    
 	obj.style.top   =   parseInt(document.body.scrollTop,10)+145;  
  	obj.style.left =   parseInt(document.body.scrollLeft,10)+400;
  }   
  catch(e){}   
  }   
  -->   
  </script>
</body>
</html>
