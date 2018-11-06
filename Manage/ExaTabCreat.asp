<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BB"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>考核期间创建</title>
</head>
<%
	Submits=request("Submits")
	SelBOPYear=request("BOPYear")
	BOPItemId=request("BOPItemId")
	SelYear=request("SelYear")
	SelBOPId=request("SelBOPId")
	SelBOPItemId=request("SelBOPItemId")
	Set rsMain = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	rsMain.cursorlocation=3
	
	if Submits=" 添加 " and BOPItemId<>"" and not isnull(BOPItemId) then
		rsTemp.open "select * from BegOfPer where BOPId=(select BOPId from BegOfPerItem "&_
			"where BOPItemId="&BOPItemId&")",G_DBConn,1,1,1
		if not rsTemp.eof then
			ExaPerId=rsTemp("ExaPerId")
			EmpCode=rsTemp("EmpCode")
			DeptCode=rsTemp("DeptCode")
			CorpCode=rsTemp("CorpCode")
		end if
		rsTemp.close
		G_DBConn.BeginTrans
		if CorpCode<>"" and not isnull(CorpCode) then
			ArrCorpCode=split(CorpCode,", ")
			for C=0 to UBound(ArrCorpCode)
				CurCorpCode=ArrCorpCode(C)
				rsMain.open "select * from Examine ",G_DBConn,2,3,1
				rsMain.addnew
					rsMain("BOPItemId")=BOPItemId
					rsMain("ExaObjType")=1
					rsMain("ExaObjCode")=CurCorpCode
				rsMain.update
				CurExamineId=rsMain("ExamineId")
				rsMain.close
				ExamineId=MaxId(1,CurCorpCode,ExaPerId)
				if ExamineId>0 then
					CopyFactor CurExamineId,ExamineId
				end if
				if G_DBConn.Errors.count>0 then
					G_DBConn.RollBackTrans 
					response.Redirect("ExaTabCreat.asp")
				end if
			next
		end if
		if DeptCode<>"" and not isnull(DeptCode) then
			ArrDeptCode=split(DeptCode,", ")
			for D=0 to UBound(ArrDeptCode)
				CurDeptCode=ArrDeptCode(D)
				rsMain.open "select * from Examine ",G_DBConn,2,3,1
				rsMain.addnew
					rsMain("BOPItemId")=BOPItemId
					rsMain("ExaObjType")=2
					rsMain("ExaObjCode")=CurDeptCode
				rsMain.update
				CurExamineId=rsMain("ExamineId")
				rsMain.close
				ExamineId=MaxId(2,CurDeptCode,ExaPerId)
				if ExamineId>0 then
					CopyFactor CurExamineId,ExamineId
				end if
				if G_DBConn.Errors.count>0 then
					G_DBConn.RollBackTrans 
					response.Redirect("ExaTabCreat.asp")
				end if
			next
		end if
		if EmpCode<>"" and not isnull(EmpCode) then
			ArrEmpCode=split(EmpCode,", ")
			for E=0 to UBound(ArrEmpCode)
				CurEmpCode=ArrEmpCode(E)
				rsMain.open "select * from Examine ",G_DBConn,2,3,1
				rsMain.addnew
					rsMain("BOPItemId")=BOPItemId
					rsMain("ExaObjType")=3
					rsMain("ExaObjCode")=CurEmpCode
				rsMain.update
				CurExamineId=rsMain("ExamineId")
				rsMain.close
				ExamineId=MaxId(3,CurEmpCode,ExaPerId)
				if ExamineId>0 then
					CopyFactor CurExamineId,ExamineId
				end if
				if G_DBConn.Errors.count>0 then
					G_DBConn.RollBackTrans 
					response.Redirect("ExaTabCreat.asp")
				end if
			next
		end if
		G_DBConn.execute("update BegOfPerItem set CreatState=1 where BOPItemId="&BOPItemId&"")
		if G_DBConn.Errors.count>0 then
			G_DBConn.RollBackTrans
			 response.Redirect("ExaTabCreat.asp")
		end if
		G_DBConn.committrans
	end if
	
	if Submits=" 删除 " then
		rsTemp.open "select count(*) as count from Examine where BOPItemId='"&CurBOPItemId&"' "&_
			"and GradeState<>0 ",G_DBConn,1,1,1
		if not rsTemp.eof then
			num=rsTemp("count")
		end if
		rsTemp.close
		if num<1 then
			G_DBConn.BeginTrans
			Rows=request("BOPItemId").count
			for i=1 to Rows 
				CurBOPItemId=request("BOPItemId")(i)
				G_DBConn.execute("delete Examine where BOPItemId='"&CurBOPItemId&"'")
				if G_DBConn.Errors.count>0 then
					G_DBConn.RollBackTrans
					 response.Redirect("ExaTabCreat.asp")
				end if
			next
			G_DBConn.execute("update BegOfPerItem set CreatState=0 where BOPItemId="&CurBOPItemId&"")
			if G_DBConn.Errors.count>0 then
				G_DBConn.RollBackTrans
				 response.Redirect("ExaTabCreat.asp")
			end if
			G_DBConn.committrans
		else
			response.Write("<script language='javascript'>alert('已经使用,不能被删除!')</script>")
		end if
	end if
	
	if Submits=" 保存 " then
		Rows=request("AllBOPItemId").count
		for i=1 to Rows
			CurBOPItemId=request("AllBOPItemId")(i)
			Remarks=request("Remarks"&CurBOPItemId)
			G_DBConn.execute("update BegOfPerItem set Remarks='"&Remarks&"' where BOPItemId="&CurBOPItemId&"")
		next
	end if
%>
<body>
<form method="post" action="ExaTabCreat.asp" name="Form1">
  <Center>
    <h2>考 核 期 间 设 置</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr>
  	<td width="408" >条件：<select name="SelYear" onChange="Form1.submit()">
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
	  <select name="SelBOPItemId">
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
      <input type="submit" name="Submits" value="查询"></td>
  	<td width="392"  align="right">	<select name="BOPYear" onChange="Form1.submit()">
		<option value="">选择年份</option>
<%
	rsMain.open "select Year(BOPYear) as BOPYear from BegOfPer group by BOPYear",G_DBConn,1,1,1
	do while not rsMain.eof
		BOPYear=rsMain("BOPYear")
%>
		<option value="<%=BOPYear%>" <%if trim(SelBOPYear)=trim(BOPYear) then response.Write("selected") end if%>><%=BOPYear%>年</option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	  </select>
  	  <select name="BOPItemId">
	    <option value="">选择期间</option>
<%
	rsMain.open "select * from BegOfPerItem where BOPId in (select BOPId from BegOfPer "&_
		"where Year(BOPYear)='"&SelBOPYear&"' and CreatState=0)",G_DBConn,1,1,1
	do while not rsMain.eof
		BOPItemId=rsMain("BOPItemId")
		BOPIName=rsMain("BOPIName")
%>
		<option value="<%=BOPItemId%>"><%=BOPIName%></option>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
	    </select>
  	  <input type="submit" name="submits" value=" 添加 ">
  	  <input type="submit" name="Submits" value=" 保存 ">
  	  <input type="submit" name="Submits" value=" 删除 " onClick="return confirm('是否确定删除此期间？')">
  	  </td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
  	<tr align="center" bgcolor="DDDDDD" class="tdcss">
		<td width="156" >年份及考核期间名称</td>
		<td width="120" >打分结束日期</td>
		<td width="353" >被考核人员表</td>
		<td width="161" >备注</td>
	</tr>
<%
	if SelBOPItemId<>"" and not isnull(SelBOPItemId) then
		Query=" and BI.BOPItemId="&SelBOPItemId&" "
	else
		if SelBOPId<>"" and not isnull(SelBOPId) then
			Query=" and B.BOPId="&SelBOPId&" "
		else
			if SelYear<>"" and not isnull(SelYear) then
				Query=" and Year(B.BOPYear)='"&SelYear&"' "
			end if
		end if
	end if
	rsMain.open "select Year(B.BOPYear) as BOPYear,BI.* from BegOfPerItem BI "&_
		"left join BegOfPer B on(B.BOPId=BI.BOPId) where BI.CreatState=1 "&Query&" ",G_DBConn,1,1,1
	do while not rsMain.eof
		BOPItemId=rsMain("BOPItemId")
		BOPYear=rsMain("BOPYear")
		BOPIName=rsMain("BOPIName")
		LastDate=rsMain("LastDate")
		Remarks=rsMain("Remarks")
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input name="BOPItemId" type="checkbox" value="<%=BOPItemId%>"><%=BOPYear%>年<%=BOPIName%><input type="hidden" name="AllBOPItemId"  value="<%=BOPItemId%>"></td>
		<td><%=LastDate%>&nbsp;</td>
		<td align="left" width="353"><center>
		  <a href="ExamineList.asp"><strong>全部</strong></a>
		</center>
          <strong>公司：</strong>
<%
	rsTemp.open "select C.CorpCode,C.CorpNameChs,E.ExamineId from Examine E left join CorpInfo C on (C.CorpCode=E.ExaObjCode) "&_
		"where BOPItemId="&BOPItemId&" and ExaObjType=1 order by C.CorpNameChs ",G_DBConn,1,1,1
	do while not rsTemp.eof
		CorpCode=rsTemp("CorpCode")
		CorpNameChs=rsTemp("CorpNameChs")
		ExamineId=rsTemp("ExamineId")
		Color=GetColor(ExamineId)
		response.Write("<a href='ExamineEdit.asp?ExamineId="&ExamineId&"&ModuleCode=BC'><font color="&Color&">"&CorpNameChs&"</font></a>"&"&nbsp;&nbsp;")
		rsTemp.movenext
	loop
	rsTemp.close
%>
	<br>
    <strong>部门：</strong>
<%
	rsTemp.open "select D.DeptCode,D.DeptName,E.ExamineId from Examine E left join Dept D on (D.DeptCode=E.ExaObjCode) "&_
		"where BOPItemId="&BOPItemId&" and ExaObjType=2 order by D.DeptName ",G_DBConn,1,1,1
	do while not rsTemp.eof
		DeptCode=rsTemp("DeptCode")
		DeptName=rsTemp("DeptName")
		ExamineId=rsTemp("ExamineId")
		Color=GetColor(ExamineId)
		response.Write("<a href='ExamineEdit.asp?ExamineId="&ExamineId&"&ModuleCode=BC'><font color="&Color&">"&DeptName&"</font></a>"&"&nbsp;&nbsp;")
		rsTemp.movenext
	loop
	rsTemp.close
%>
	<br>
    <strong>人员：</strong>
<%
	rsTemp.open "select L.EmpCode,L.EmpNameChs,E.ExamineId from Examine E left join Employee L on (L.EmpCode=E.ExaObjCode) "&_
		"where BOPItemId="&BOPItemId&" and ExaObjType=3 order by L.EmpNameChs",G_DBConn,1,1,1
	do while not rsTemp.eof
		EmpCode=rsTemp("EmpCode")
		EmpNameChs=rsTemp("EmpNameChs")
		ExamineId=rsTemp("ExamineId")
		Color=GetColor(ExamineId)
		response.Write("<a href='ExamineEdit.asp?ExamineId="&ExamineId&"&ModuleCode=BC'><font color="&Color&">"&EmpNameChs&"</font></a>"&"&nbsp;&nbsp;")
		rsTemp.movenext
	loop
	rsTemp.close
%>
	  </td>
		<td><input type="text" class="midinput" value="<%=Remarks%>" name="Remarks<%=BOPItemId%>"></td>
	</tr>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
  </table>
  <p>&nbsp;</p>
</form>
<script language="vbscript">

</script>
</body>
</html>
