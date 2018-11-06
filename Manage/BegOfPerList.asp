<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BA"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>期初考核设置列表</title>
</head>
<%
	Submit=request("Submit")
	txtYear=request("txtYear")
	SearchStr=request("SearchStr")
	set RSDB=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	set rsPeriod=Server.CreateObject("ADODB.Recordset")
	set rsItem=Server.CreateObject("ADODB.Recordset")
	Set rsEmp = Server.CreateObject("ADODB.Recordset")
	Set rsDept = Server.CreateObject("ADODB.Recordset")
	Set rsCorp = Server.CreateObject("ADODB.Recordset")
	RSDB.cursorlocation=3
	
	if Submit=" 添加 " then
		strYear=txtYear&"-01-01"
		strError=""
		if not isdate(strYear) then
			strError="填写的年份格式不对，请检查后重新添加。"
		else		
			RSDB.open "select * from BegOfPer ",G_DBConn,2,3,1
			RSDB.addnew
				RSDB("BOPYear")=txtYear&"-01-01"
			RSDB.update
			BOPId=RSDB("BOPId")
			RSDB.close
			response.Redirect "BegOfPerEdit.asp?UrlYear="&txtYear&""
			response.End()
		end if
	end if
	
	if SearchStr<>"" and not isnull(SearchStr) then
		if isdate(SearchStr&"-01-01") then
			Query=" and Year(BOPYear)='"&SearchStr&"' "
		end if
	end if
%>
<body>
<form method="post" action="BegOfPerList.asp" name="Form1">
  <Center>
    <h2>期 初 考 核 设 置 列 表</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr>
  	<td >年份
      <input name="SearchStr" type="text" class=input title="年份" value="<%=SearchStr%>" maxlength="4">  
      <input type="submit" name="Submit" value="查询">  </td>
  	<td  align="right"><font color="#FF0000"><%=strError%></font><input name="txtYear" type="text" class=input title="年份" value="" maxlength="4">
  	  <input type="submit" name="submit" value=" 添加 "></td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
  <tr align="center" bgcolor="DDDDDD" class="tdcss">
		<td width="99">年份</td>
		<td width="117">考核方法</td>
		<td width="166" bgcolor="DDDDDD">考核期间</td>
		<td>被考核人员</td>
</tr>
  <%
  	RSDB.open "select distinct BOPYear from BegOfPer where 1=1 "&Query&" ",G_DBConn,1,1,1
	do while not RSDB.eof
	RowIndex=0
	BOPYear=RSDB("BOPYear")
	BOPYear=datepart("yyyy",BOPYear)
		rsTemp.open "select * from BegOfPer where Year(BOPYear)='"&BOPYear&"'",G_DBConn,1,1,1
		do while not rsTemp.eof
		Rows=rsTemp.recordCount
		BOPId=rsTemp("BOPId")
		ExaPerId=rsTemp("ExaPerId")
		EmpCode=rsTemp("EmpCode")
		DeptCode=rsTemp("DeptCode")
		CorpCode=rsTemp("CorpCode")
  %>
  	<tr align="center" bgcolor="#FFFFFF" class="tdcss">
	<%if RowIndex=0 then%>
		<td rowspan="<%=Rows%>"><a href="BegOfPerEdit.asp?UrlYear=<%=BOPYear%>"><%=BOPYear%></a></td>
	<%end if%>
		<td><%
			if ExaPerId<>"" and not isnull(ExaPerId) then
				rsPeriod.open "select * from ExaPeriod where ExaPerId="&ExaPerId&"",G_DBConn,1,1,1
				if not rsPeriod.eof then
					response.Write(rsPeriod("ExaPerName"))
				end if
				rsPeriod.close
			else
				response.Write("&nbsp;")
			end if			
		%></td>
		<td><%
			rsItem.open "select * from BegOfPerItem where BOPId="&BOPId&"",G_DBConn,1,1,1
			strBOPIName=""
			do while not rsItem.eof 
				if strBOPIName="" then
					strBOPIName=rsItem("BOPIName")
				else
					strBOPIName=strBOPIName&"<br>"&rsItem("BOPIName")
				end if
				
				rsItem.movenext
			loop
			rsItem.close
			response.Write(strCheck(strBOPIName))
		%></td>
		<td width="408" align="left">
		<strong>公司：</strong>
		<% 
			CorpNameChs=""
			if CorpCode<>"" and not isnull(CorpCode) then
				ArrCorpCode=split(CorpCode,", ")
				CorpNameChs=""
				for C=0 to UBound(ArrCorpCode)
					CorpCode=ArrCorpCode(C)
					rsCorp.open "select * from CorpInfo where CorpCode='"&CorpCode&"' ",G_DBConn,1,1,1
					do while not rsCorp.eof
						CorpNameChs=CorpNameChs&"&nbsp;&nbsp;"&rsCorp("CorpNameChs")&"("&AvgScore(1,CorpCode,BOPId)&")"
						rsCorp.movenext
					loop
					rsCorp.close	
				next
			end if
			response.Write(CorpNameChs)
			%>
			<br>
			<strong>部门：</strong>
			<%
			DeptName=""
			if DeptCode<>"" and not isnull(DeptCode) then
				ArrDeptCode=split(DeptCode,", ")
				DeptName=""
				for D=0 to UBound(ArrDeptCode)
					DeptCode=ArrDeptCode(D)
					rsDept.open "select * from Dept where DeptCode='"&DeptCode&"'",G_DBConn,1,1,1
					do while not rsDept.eof
						DeptName=DeptName&"&nbsp;&nbsp;"&rsDept("DeptName")&"("&AvgScore(2,DeptCode,BOPId)&")"
						rsDept.movenext
					loop
					rsDept.close
				next
			end if
			response.Write(DeptName)
			%>
			<br>
			<strong>人员：</strong>
	  <%
	  		EmpNameChs=""
			if EmpCode<>"" and not isnull(EmpCode) then
				ArrEmpCode=split(EmpCode,", ")
				EmpNameChs=""
				for E=0 to UBound(ArrEmpCode)
					EmpCode=ArrEmpCode(E)
					rsEmp.open "select * from Employee where EmpCode='"&EmpCode&"'",G_DBConn,1,1,1
					do while not rsEmp.eof
						EmpNameChs=EmpNameChs&"&nbsp;&nbsp;"&rsEmp("EmpNameChs")&"("&AvgScore(3,EmpCode,BOPId)&")"
						rsEmp.movenext
					loop
					rsEmp.close
				next
			end if
			response.Write(EmpNameChs)
		%>&nbsp;	  </td>
	</tr>
  <%
  			RowIndex=RowIndex+1
  			rsTemp.movenext
		loop
		rsTemp.close
  		RSDB.movenext
	loop
  	RSDB.close
  %>
</table>
</form>
</body>
</html>
