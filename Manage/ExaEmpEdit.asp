<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BC"%>
<html>
<head>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE3 {color: #0033FF}
-->
</style>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<base target="_self">
<title>考核人员设置</title>
</head>
<%
	Set rsMain = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	Set rsCorp = Server.CreateObject("ADODB.Recordset")
	Set rsWeigh = Server.CreateObject("ADODB.Recordset")
	
	ExaItemId=request("ExaItemId")
	Submits=request("Submits")
	ClientY=request("ClientY")
	ExamineId=request("ExamineId")
	
	if Submits=" 保存 " then
		'删除现有的人员
		G_DBConn.execute("delete ExaEmpTab where ExaItemId="&ExaItemId&"")
		'重新添加
		strErr=""
		AllWeigh=0
		Rows=request("EmpScope"&ExaItemId).count
		for i=1 to Rows
			EmpScope=request("EmpScope"&ExaItemId)(i)
			Weighing=request("Weighing"&EmpScope)
			if Weighing="" or isnull(Weighing) then Weighing=0
			if IsNumeric(Weighing) then
				if Weighing<>0 then
					rsMain.open "select * from ExaEmpTab where ExaItemId="&ExaItemId&"",G_DBConn,2,3,1
					rsMain.addnew
						rsMain("ExaItemId")=ExaItemId
						rsMain("ExaEmpCode")=EmpScope
						if Weighing<>"" and not isnull(Weighing) then
							rsMain("Weighing")=Weighing
							AllWeigh=cdbl(AllWeigh)+cdbl(Weighing)
						end if
					rsMain.update
					rsMain.close
				end if
			end if
		next
		if cdbl(AllWeigh)=100 then
			G_DBConn.execute("update ExamineItem set SumEmpWeigh=1 where ExaItemId="&ExaItemId&"")
		else
			G_DBConn.execute("update ExamineItem set SumEmpWeigh=0 where ExaItemId="&ExaItemId&"")
			Response.Write("<script language='javascript'>alert('请注意！权重相加不为100%。')</script>")
		end if
	end if
	
	'显示数据
	strExaEmpCode=""
	rsMain.open "select * from ExaEmpTab where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
	do while not rsMain.eof 
		strExaEmpCode=strExaEmpCode&rsMain("ExaEmpCode")&", "
		rsMain.movenext
	loop
	rsMain.close
%>
<body>
<form name="form1" method="post" action="ExaEmpEdit.asp?ExaItemId=<%=ExaItemId%>&ClientY=<%=ClientY%>&ExamineId=<%=ExamineId%>">
<Center>
    <h2>考 核 人 员 设 置</h2>
</Center>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="600">
<tr valign="bottom">
  <td><span class="STYLE3">在选中人员后面的方框内录入此人员权重。<br>
  </span><span class="STYLE1">权重必需为数据格式且不能为0，否则数据将不会被保存！<br><%=strErr%></span></td>
<td align="right">
  <input name="Submits" type="submit" id="Submits" value=" 保存 ">
  <input name="Submits" type="submit" id="Submits" value=" 关闭 " onClick="PageClose(<%=ExamineId%>)"></td>
</tr>
</table>
<%
	rsCorp.open "select * from CorpInfo ",G_DBConn,1,1,1
	do while not rsCorp.eof
	CorpId=rsCorp("CorpId")
	CorpNameChs=rsCorp("CorpNameChs")
%>
<table border="1" align="center" width="600" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="0" cellspacing="0" style=" margin-top:5px; margin-bottom:5px;" >
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
			EmpCode=rsTemp("EmpCode")
%>
        <input type="checkbox" name="EmpScope<%=ExaItemId%>" value="<%=EmpCode%>" <% if InStr(strExaEmpCode,rsTemp("EmpCode")) then Response.Write "Checked" %>>
<%
					Response.Write rsTemp("EmpNameChs")
					if Len(rsTemp("EmpNameChs"))<6 then
				   			Response.Write Replace(space(5-Len(rsTemp("EmpNameChs")))," ","&nbsp;&nbsp;")
				 	end if
					Weighing=""
					rsWeigh.open "select * from ExaEmpTab where ExaItemId="&ExaItemId&" and "&_
						"ExaEmpCode='"&EmpCode&"'",G_DBConn,1,1,1
					if not rsWeigh.eof then
						Weighing=rsWeigh("Weighing")
					end if
					rsWeigh.close
%>
<input name="Weighing<%=EmpCode%>" type="text" class="priceinput" value="<%=Weighing%>">%
<%
						i=i+1
						if i=4 then
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
		rsCorp.Movenext
	loop
	rsCorp.close
%>
<input type="hidden" name="txt0" value="<%=ClientY%>">
<input type="hidden" name="ExaId" value="<%=ExamineId%>">
</form>
<script language="javascript">
	function PageClose(ExamineId)
	{
		window.returnValue = document.getElementById("txt0").value; 
		window.dialogArguments.location.reload("ExamineEdit.asp"+"?Window_OffsetY="+window.returnValue+"&ExamineId="+ExamineId); 
		window.close();
	}
</script>
</body>
</html>
