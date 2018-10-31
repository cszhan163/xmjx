<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BA"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>期初考核设置编辑</title>
</head>
<%
	Submit=Request("Submit")
	CurYear=request("UrlYear")
	SelBOPId=request("SelBOPId")

	Set RSDB = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	Set rsEmp = Server.CreateObject("ADODB.Recordset")
	Set rsDept = Server.CreateObject("ADODB.Recordset")
	Set rsCorp = Server.CreateObject("ADODB.Recordset")
	Set rsMain = Server.CreateObject("ADODB.Recordset")
    RSDB.cursorlocation=3
	
	if Submit=" 返回 " then
		response.Redirect "BegOfPerList.asp"
		response.End()
	end if
	
	if Submit=" 取消 " then
		response.Redirect "BegOfPerEdit.asp?UrlYear="&CurYear&""
	end if
	
	if Submit=" 添加 " then
		RSDB.open "select * from BegOfPer ",G_DBConn,2,3,1
		RSDB.addnew
			RSDB("BOPYear")=CurYear&"-01-01"
		RSDB.update
		RSDB.close
		response.Redirect "BegOfPerEdit.asp?UrlYear="&CurYear&" "
		response.End()
	end if
	
	if Submit=" 保存 " then
		Rows=request("AllBOPId").count
		for i=1 to Rows
			CurBOPId=request("AllBOPId")(i)
			RSDB.open "select * from BegOfPer where BOPId="&CurBOPId&"",G_DBConn,2,3,1
				if request("ExaPer"&CurBOPId)<>"" then
					RSDB("ExaPerId")=request("ExaPer"&CurBOPId)
				end if
				'if request("CorpCode"&CurBOPId).count<>0 then
					RSDB("CorpCode")=request("CorpCode"&CurBOPId)
				'end if
				'if request("DeptCode"&CurBOPId).count<>0 then
					RSDB("DeptCode")=request("DeptCode"&CurBOPId)
				'end if
				'if request("EmpCode"&CurBOPId).count<>0 then
					RSDB("EmpCode")=request("EmpCode"&CurBOPId)
				'end if
			RSDB.update
			RSDB.close
			'开始设置后又添加的人员
			RSDB.open "select BOPItemId from BegOfPerItem where BOPId="&CurBOPId&"",G_DBConn,1,1,1
			do while not RSDB.eof
				CurBOPItemId=RSDB("BOPItemId")
				rsTemp.open "select * from Examine where BOPItemId="&CurBOPItemId&"",G_DBConn,1,1,1
				if not rsTemp.eof then
					if request("CorpCode"&CurBOPId).count<>0 then
						ArrCorpCode=split(request("CorpCode"&CurBOPId),", ")
						for C=0 to UBound(ArrCorpCode)
							CurCorpCode=ArrCorpCode(C)
							rsMain.open "select * from Examine "&_
							"where BOPItemId="&CurBOPItemId&" and ExaObjType=1 and ExaObjCode='"&CurCorpCode&"'",G_DBConn,2,3,1
							if rsMain.eof then
								rsMain.addnew
								rsMain("BOPItemId")=CurBOPItemId
								rsMain("ExaObjType")=1
								rsMain("ExaObjCode")=CurCorpCode
								rsMain.update
							end if
							rsMain.close
						next
					end if
					if request("DeptCode"&CurBOPId).count<>0 then
						ArrDeptCode=split(request("DeptCode"&CurBOPId),", ")
						for D=0 to UBound(ArrDeptCode)
							CurDeptCode=ArrDeptCode(D)
							rsMain.open "select * from Examine "&_
							"where BOPItemId="&CurBOPItemId&" and ExaObjType=2 and ExaObjCode='"&CurDeptCode&"'",G_DBConn,2,3,1
							if rsMain.eof then
								rsMain.addnew
								rsMain("BOPItemId")=CurBOPItemId
								rsMain("ExaObjType")=2
								rsMain("ExaObjCode")=CurDeptCode
								rsMain.update
							end if
							rsMain.close
						next
					end if
					if request("EmpCode"&CurBOPId).count<>0 then
						ArrEmpCode=split(request("EmpCode"&CurBOPId),", ")
						for E=0 to UBound(ArrEmpCode)
							CurEmpCode=ArrEmpCode(E)
							rsMain.open "select * from Examine "&_
							"where BOPItemId="&CurBOPItemId&" and ExaObjType=3 and ExaObjCode='"&CurEmpCode&"'",G_DBConn,2,3,1
							if rsMain.eof then
								rsMain.addnew
								rsMain("BOPItemId")=CurBOPItemId
								rsMain("ExaObjType")=3
								rsMain("ExaObjCode")=CurEmpCode
								rsMain.update
							end if
							rsMain.close
						next
					end if
				end if
				rsTemp.close
				RSDB.Movenext
			loop
			RSDB.close
		next
	end if
	
	if Submit=" 删除 " then
		Rows=request("strSelBOPId").count
		for i=1 to Rows
			CurBOPId=request("strSelBOPId")(i)
			RSDB.open "select count(*) as count from BegOfPerItem where BOPID="&CurBOPId&"",G_DBConn,1,1,1
			if not RSDB.eof then
				num=RSDB("count")
			end if
			RSDB.close
			if num<=0 then
				G_DBConn.execute("delete BegOfPer where BOPId in("&CurBOPId&")")
			else
				response.Write("<script language='javascript'>alert('已经使用,不能被删除!')</script>")
			end if
		next
	end if
	
%>
<body>
<form method="post" action="BegOfPerEdit.asp?UrlYear=<%=CurYear%>" name="Form1">
  <Center>
    <h2>期 初 考 核 设 置 编 辑</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="600">
  <tr>
  	<td width="173" >年份:<%=CurYear%>年<input type="hidden" value="<%=CurYear%>" name="UrlYear"></td>
  	<td width="427"  align="right">	<input type="submit" name="submit" value=" 添加 ">
  	  <input type="submit" name="Submit" value=" 保存 ">
  	  <input type="submit" name="Submit" value=" 删除 " onClick="return confirm('是否确定删除此数据？')">
  	  <input type="submit" name="Submit" value=" 取消 ">
  	  <input type="submit" name="Submit" value=" 返回 ">
	  <input type="hidden" name="SelBOPId"  value="<%=SelBOPId%>"></td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="600" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
  	<tr align="center" bgcolor="DDDDDD" class="tdcss">
		<td width="90">序号</td>
		<td width="264">考核方法</td>
		<td width="238">详细期间设置</td>
	</tr>
<%
	RowIndex=0
	RSDB.open "select * from BegOfPer where Year(BOPYear)='"&CurYear&"'",G_DBConn,1,1,1
	do while not RSDB.eof 
	RowIndex=RowIndex+1
	BOPId=RSDB("BOPId")
	ExaPerId=RSDB("ExaPerId")
	strEmpCode=RSDB("EmpCode")
	strDeptCode=RSDB("DeptCode")
	strCorpCOde=RSDB("CorpCode")
%>
	<tr align="center" bgcolor="#FFFFFF" class="tdcss">
		<td><a href="BegOfPerEdit.asp?SelBOPId=<%=BOPId%>&UrlYear=<%=CurYear%>"><%=RowIndex%></a><input name="strSelBOPId" type="checkbox" value="<%=BOPId%>"><input type="hidden" name="AllBOPId" value="<%=BOPId%>"></td>
		<td>
		<select name="ExaPer<%=BOPId%>">
		<option value="">请选择期间</option>
		<%
			rsTemp.open "select * from ExaPeriod ",G_DBConn,1,1,1
			do while not rsTemp.eof
			CurExaPerId=rsTemp("ExaPerId")
			ExaPerName=rsTemp("ExaPerName")
		%>
		<option value="<%=CurExaPerId%>" <%if CurExaPerId=ExaPerId then response.Write("selected") end if%>><%=ExaPerName%></option>
		<%
				rsTemp.movenext
			loop
			rsTemp.close
		%>
		</select>
		</td>
		<td><a href="BegOfPerItem.asp?BOPId=<%=BOPId%>">设置详细期间</a></td>
	</tr>
<%
	if trim(BOPId)=trim(SelBOPId) then
%>
	<tr><td colspan="3" align="center" bgcolor="#FFFFFF">
		<%
			rsCorp.open "select * from CorpInfo",G_DBConn,1,1,1
			do while not rsCorp.EOF 
			CorpId=rsCorp("CorpId")
			CorpCode=rsCorp("CorpCode")
			CorpNameChs=rsCorp("CorpNameChs")
			if IsSelect(CurYear,"CorpCode",CorpCode)="true" then
				color="#FF0000"
			else
				color="#000000"
			end if
		%>
			<table border="1" align="center" width="560" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="0" cellspacing="0" style=" margin-top:5px; margin-bottom:5px;" >
		  <tr>
			<td colspan="2" align="center" bgcolor="#DDDDDD" height="25">
			<input name="CorpCode<%=SelBOPId%>" type="checkbox" value="<%=CorpCode%>" <%if InStr(strCorpCOde,CorpCode) then Response.Write "Checked" %>><font color=<%=color%>><%=CorpNameChs%></font></td>
		  </tr>
		<%
		RSDept.Open "Select DeptCode,DeptName From Dept where CorpId="&CorpId&" order by DeptID",G_DBConn,2,3,1
			Do While Not RSDept.Eof 
			DeptCode=rsDept("DeptCode")
			DeptName=rsDept("DeptName")
			if IsSelect(CurYear,"DeptCode",DeptCode)="true" then
				color="#FF0000"
			else
				color="#000000"
			end if
		%> 
		  <tr class=tdcss>
			<td nowrap bgcolor="#DDDDDD" align="center" height="25"><input name="DeptCode<%=SelBOPId%>" type="checkbox" value="<%=DeptCode%>" <%if InStr(strDeptCOde,DeptCode) then Response.Write "Checked" %>><font color=<%=color%>><%=DeptName%></font></td>
			<td nowrap valign="top" bgcolor="#FFFFFF" align="left">&nbsp;
		<%
				rsEmp.Open "Select Grade,EmpCode,EmpNameChs From Employee Where DeptCode='"&DeptCode&"' ",G_DBConn,2,3,1
					i=1
					Do While Not rsEmp.Eof
					EmpCode=rsEmp("EmpCode")
					EmpNameChs=rsEmp("EmpNameChs")
					if IsSelect(CurYear,"EmpCode",EmpCode)="true" then
						color="#FF0000"
					else
						color="#000000"
					end if
		%>
				<input type="checkbox" name="EmpCode<%=SelBOPId%>" value="<%=EmpCode%>" <% if InStr(strEmpCode,EmpCode) then Response.Write "Checked" %>><font color=<%=color%>>
		<%
							Response.Write EmpNameChs
								if Len(EmpNameChs)<5 then
									Response.Write Replace(space(4-Len(EmpNameChs))," ","&nbsp;&nbsp;")
								end if
								i=i+1
								if i=5 then
									i=1
									Response.Write "<br>&nbsp;"
								end if
					   rsEmp.MoveNext
					Loop
				rsEmp.Close
		%></font></td>
		  </tr>
		<%
				RSDept.MoveNext
			Loop
		RSDept.Close
		%> 
		</table>
		<%
			rsCorp.movenext
			loop
			rsCorp.close
		%>
	</td>
	</tr> 
<%
	end if
		RSDB.movenext
	loop
	RSDB.close
%>
</table>
</form>
</body>
</html>
