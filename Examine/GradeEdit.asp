<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "**"%>
<html>
<head>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�������ֱ༭</title>
</head>
<%
	ExamineId=request("ExamineId")
	Submits=request("Submits")
	ExaEmpCode=request("ExaEmpCode")
	IsSubmit=request("IsSubmit")
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	set rsMarks=Server.CreateObject("ADODB.Recordset")

	if Submits=" ���� " then
		response.Redirect("GradeList.asp")
		response.End()
	end if

	if Submits=" ���� " then
		Rows=request("ALLExaItemId").count
		for i=1 to Rows
			CurExaItemId=request("ALLExaItemId")(i)
			Marks=request("Marks"&CurExaItemId)
			rsTemp.open "select EI.*,EF.ExaFactorName,EF.IsCanOver from ExamineItem EI "&_
				"left join ExaFactor EF on EF.ExaFactorId=EI.ExaFactorId "&_
				"where ExaItemId="&CurExaItemId&"",G_DBConn,1,1,1
			if not rsTemp.eof then
				MarksType=rsTemp("MarksType")
				IsRepeat=rsTemp("IsRepeat")
				ExaFactorName=rsTemp("ExaFactorName")
				IsCanOver=rsTemp("IsCanOver")
			end if
			'���жϵ�һ���Ƿ��ظ�
			'if IsRepeat="True" or SelFirst(CurExaItemId,MarksType,Marks)="False" then
			if IsNumeric(Marks) then
				if (CheckGrade(CurExaItemId,Marks)="True" or IsCanOver=true) then
					rsMarks.open "select * from ExaEmpTab where ExaItemId="&CurExaItemId&" "&_
						"and ExaEmpCode='"&ExaEmpCode&"'",G_DBConn,2,3,1
						rsMarks("Marks")=Marks
					rsMarks.update
					rsMarks.close
				else
					response.Write("<script language='javascript'>alert('Ҫ��:("&ExaFactorName&")   ����:("&Marks&")  �����Ͽ��˱�׼���������ֽ����ᱻ���棬�����������д��');</script>")
				end if
			end if
			'else
'				response.Write("<script language='javascript'>alert('Ҫ��:'"&ExaFactorName&" �����ظ�,������������֣�);
'			end if
			rsTemp.close
		next
	end if

	if IsSubmit="Ok" then
		G_DBConn.execute("update ExaEmpTab set State=4 where ExaEmpCode='"&ExaEmpCode&"' and "&_
			"ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")")
		rsMain.open "select Count(*) as count from ExaEmpTab ET "&_
			"left join ExamineItem EI on EI.ExaItemId=ET.ExaItemId "&_
			"where EI.ExamineId="&ExamineId&" and ET.State<>4 ",G_DBConn,1,1,1
		if not rsMain.eof then
			num=rsMain("count")
		end if
		rsMain.close
		if num<1 then
			G_DBConn.execute("update Examine set GradeState=3,LastDate=getdate() where ExamineId="&ExamineId&"")
		end if
	end if

	if Submits="�����˻�" then
		G_DBConn.execute("update ExaEmpTab set State=2 where ExaEmpCode='"&ExaEmpCode&"' and "&_
			"ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")")
	end if

	if Submits=" �˻� " then
		G_DBConn.execute("update ExaEmpTab set State=3 where ExaEmpCode='"&ExaEmpCode&"' and "&_
			"ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")")
	end if

	'if Submits=" ȷ�� " then
'		G_DBConn.execute("update ExaEmpTab set State=4 where ExaEmpCode='"&ExaEmpCode&"' and "&_
'			"ExaItemId in (select ExaItemId from ExamineItem where ExamineId="&ExamineId&")")
'		rsMain.open "select Count(*) as count from ExaEmpTab ET "&_
'			"left join ExamineItem EI on EI.ExaItemId=ET.ExaItemId "&_
'			"where EI.ExamineId="&ExamineId&" and ET.State<>4 ",G_DBConn,1,1,1
'		if not rsMain.eof then
'			num=rsMain("count")
'		end if
'		rsMain.close
'		if num<1 then
'			G_DBConn.execute("update Examine set GradeState=3,LastDate=getdate() where ExamineId="&ExamineId&"")
'		end if
'	end if
	'��ʾ����
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

	rsMain.open "select State from ExaEmpTab where ExaEmpCode='"&ExaEmpCode&"' and "&_
		"ExaItemId in (select ExaItemId from ExamineItem where ExamineId='"&ExamineId&"')",G_DBConn,1,1,1
	if not rsMain.eof then
		ExaGradeState=rsMain("State")
	end if
	rsMain.close

	if ExaGradeState<>0 and ExaGradeState<>3 then
		BtnSave="disabled"
	end if
	if ExaGradeState<>1 and ExaGradeState<>4  then
		BtnApplyBack="disabled"
	end if
	if ExaGradeState<>1 then
		BtnOk="disabled"
	end if
	if ExaGradeState<>1 and  ExaGradeState<>2 then
		BtnBack="disabled"
	end if
%>
<body>
<form method="post" action="GradeEdit.asp?ExamineId=<%=ExamineId%>&ExaEmpCode=<%=ExaEmpCode%>" name="Form1">
  <Center>
    <h2><font color="#FF0000"><%=BOPYear%>��<%=BOPIName%>&nbsp;<%=ObjName%>&nbsp;</font>��Ч�������ֱ�</h2>
  </Center>

  <table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
    <tr>
      <td align="right"><font color="#FF0000"><%=strErr%></font>
		  <%if ModuleCode="AA" then%>
		  <input type="submit" name="Submits" value=" ���� " <%=BtnSave%>>
          <input name="Submits" type="button" id="Submits" value=" �ύ " <%=BtnSave%> onClick="CheckSub()">
		  <input name="IsSubmit" type="hidden" value="">
          <%if ExaGradeState=3 then %>
          <span class="STYLE1">���˻�</span>
		  <%elseif ExaGradeState=2 then %>
		  <span class="STYLE1">�����˻�</span>
	      <%else%><input name="Submits" type="submit" id="Submits" value="�����˻�" <%=BtnApplyBack%>><%end if%>
		  <%end if%>
		  <%if ModuleCode="BD" then%>
          <!--input name="Submits" type="submit" id="Submits" value=" ȷ�� " <%'=BtnOk%>-->
          <input name="Submits" type="submit" id="Submits" value=" �˻� " <%=BtnBack%>>
		 <%end if%>
          <input name="Submits" type="submit" id="Submits" value=" ���� ">
      </td>
    </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
    <tr align="center">
      <td width="88" bgcolor="DDDDDD">�����ڼ�</td>
      <td width="148" bgcolor="#FFFFFF"><%=BOPYear&BOPIName%>&nbsp;</td>
      <td width="74" bgcolor="DDDDDD">���˲���</td>
      <td width="201" bgcolor="#FFFFFF"><%=DeptName%>&nbsp;</td>
      <td width="127" bgcolor="DDDDDD">���˸�λ</td>
      <td width="148" bgcolor="#FFFFFF"><%=GroupName%>&nbsp;</td>
    </tr>
    <tr align="center">
      <td bgcolor="DDDDDD">������Ա</td>
      <td bgcolor="#FFFFFF"><%=EmpNameChs%>&nbsp;</td>
      <td bgcolor="DDDDDD">��ֹ����</td>
      <td bgcolor="#FFFFFF"><%=OverDate%>&nbsp;</td>
      <td bgcolor="DDDDDD">���յ÷�</td>
      <td bgcolor="#FFFFFF"><%=ExaTotalScore(ExamineId,ExaEmpCode)%>��&nbsp;</td>
    </tr>
  </table>
<%
	FactCount=0
	rsMain.open "select EI.*,EP.ExaProdName,EF.ExaFactorName,EF.ExaNorm from ExamineItem EI "&_
		"left join ExaEmpTab ET on(ET.ExaItemId=EI.ExaItemId) "&_
		"left join ExaProdject EP on(EP.ExaProdId=EI.ExaProdId) "&_
		"left join ExaFactor EF on(EF.ExaFactorId=EI.ExaFactorId) "&_
		"where ExamineId="&ExamineId&" and ET.ExaEmpCode='"&ExaEmpCode&"'",G_DBConn,1,1,1
	do while not rsMain.eof
	FactCount=FactCount+1
	ExaItemId=rsMain("ExaItemId")
	Weighing=rsMain("Weighing")
	IsRepeat=rsMain("IsRepeat")
	MarksType=rsMain("MarksType")
	ExaProdName=rsMain("ExaProdName")
	ExaFactorName=rsMain("ExaFactorName")
	ExaNorm=rsMain("ExaNorm")
%>
  <table border="1" align="center" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999">
    <tr align="center" bgcolor="DDDDDD">
      <td width="36">���</td>
      <td width="118" bgcolor="DDDDDD">������Ŀ</td>
      <td width="133">����Ҫ��</td>
      <td width="220" bgcolor="DDDDDD">���˱�׼</td>
      <td width="85">Ȩ��</td>
      <td width="95">���</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
      <td rowspan="3"><%=FactCount%>)&nbsp;<input type="hidden" name="ALLExaItemId" value="<%=ExaItemId%>"></td>
      <td><%=ExaProdName%>&nbsp;</td>
      <td><%=ExaFactorName%>&nbsp;</td>
      <td width="270" align="left"><%=ExaNorm%>&nbsp;</td>
      <td><%=Weighing%>%</td>
      <td><%=ExaFactorScore(ExaItemId,ExaEmpCode)%>��&nbsp;</td>
    </tr>
    <tr align="center" bgcolor="DDDDDD">
      <td colspan="1">�Ƿ���ظ�����</td>
      <td bgcolor="DDDDDD">��ֵ����</td>
      <td colspan="3">���ְ취����׼</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td colspan="1" align="center">
	  <%if IsRepeat="True" then
	  		response.Write("��")
		else
			response.Write("��")
		end if%>
	  </td>
      <td align="center">
	   <%if MarksType="True" then
	  		response.Write("��ֵ")
		else
			response.Write("����")
		end if%>
	  </td>
      <td width="480" colspan="3"><%
	'���ְ취����׼
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
                  <%if MarksType=0 then%><%=MaxMarks%>-<%=MinMarks%>
                  <%else%><%=MaxMarks%><%end if%>��
			  </td>
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
    <tr align="center" bgcolor="DDDDDD">
      <td colspan="3" bgcolor="DDDDDD">������Ա</td>
      <td>Ȩ��</td>
      <td colspan="2">����</td>
      <%
	rsTemp.open "select EE.*,E.EmpNameChs from ExaEmpTab EE "&_
		"left join Employee E on (E.EmpCode=EE.ExaEmpCode) "&_
		"where ExaItemId="&ExaItemId&" and EE.ExaEmpCode='"&ExaEmpCode&"'",G_DBConn,1,1,1
	do while not rsTemp.eof
	EmpNameChs=rsTemp("EmpNameChs")
	Weighing=rsTemp("Weighing")
	Marks=rsTemp("Marks")
%>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="3"><%=EmpNameChs%>&nbsp;</td>
      <td><%=Weighing%>%&nbsp;</td>
      <td colspan="2"><input name="Marks<%=ExaItemId%>" type="text" class="priceinput" value="<%=Marks%>" onChange="CheckNum(<%=ExaItemId%>)">��</td>
    </tr>
    <%
		rsTemp.movenext
	loop
	rsTemp.close
%>
  </table>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
<input type="hidden" name="ExaEmpCode" value="<%=ExaEmpCode%>">
</form>
<script language="vbscript">
sub CheckNum(ExaItemId)
	strName="Marks"&ExaItemId
	ScodeValue=eval("Form1."&strName&".value")
	if not IsNumeric(ScodeValue) then
		alert("����д��ȷ�����ָ�ʽ��")
	end if
end sub
sub CheckSub()
	<%
		set rsSub=server.CreateObject("ADODB.Recordset")
		rsSub.open "select count(*) as ZNum from ExaEmpTab ET "&_
			" left join ExamineItem EI on EI.ExaItemId=ET.ExaItemId "&_
			" where ExamineId="&ExamineId&" and ExaEmpCode='"&ExaEmpCode&"' and Marks=0 ",G_DBConn,1,1,1
			ZNum=rsSub("ZNum")
	%>
	if Number("<%=ZNum%>")>0 then
		if MsgBox("��<%=ZNum%>�����Ϊ�㣬ȷ���ύ��",vbYesNo)=vbYes then
			Form1.IsSubmit.value="Ok"
			Form1.submit()
		end if
	else
		Form1.IsSubmit.value="Ok"
		Form1.submit()
	end if
	<%
		rsSub.close
	%>
end sub
</script>
</body>
</html>
