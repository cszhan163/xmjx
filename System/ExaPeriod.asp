<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "CC"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������������</title>
</head>

<body>
<%
	Submit=Request("Submit")
	EditNum=Request("EditNum")
	If ISNumeric(EditNum) then
		EditNum=Cint(EditNum)
	else
		EditNum=-1
	End if
%>
<%
	if Submit="����" then
%> 
<meta http-equiv="refresh" content="0;URL=AllMana.asp">   
<%
		Response.end
	end if
	Set RSDB = Server.CreateObject("ADODB.Recordset")
    RSDb.cursorlocation=3

	If Submit="  ɾ��  " then
		ExaPerIdX=Request("ExaPerId")
		If IsNumeric(ExaPerIdX) then
			ExaPerIdX=Cint(ExaPerIdX)
		else
			ExaPerIdX=-1
		end if
		RSDb.open "select Count(*) count from BegOfPer where ExaPerId="&ExaPerIdX&"",G_DBConn,1,1,1
		if not rsdb.eof then
			num=rsdb("count")
		end if
		RSDB.close
		if num<1 then
			RSDB.Open "Delete From ExaPeriod where ExaPerId="&ExaPerIdX,G_DBConn,2,3,1
		else
			response.Write("<script language='javascript'>alert('�Ѿ�ʹ��,���ܱ�ɾ��!')</script>")
		end if
	End if
	If Submit="  ����  " then
		XXStr=Request("BJStr")
		If XXStr="���" then
          		RSDB.Open "Select * From ExaPeriod ",G_DBConn,2,3,1
				RSDB.AddNew
			    RSDB("ExaPerName")=Request("ExaPerName")
		    	RSDB.UpDate
			RSDB.Close
		End if
		If XXStr="�༭" then
			ExaPerIdX=Request("ExaPerId")
			If IsNumeric(ExaPerIdX) then
				ExaPerIdX=Cint(ExaPerIdX)
			else
				ExaPerIdX=-1
			end if
			RSDB.Open "Select * From ExaPeriod Where ExaPerId="&ExaPerIdX,G_DBConn,2,3,1
				If Not RSDB.Eof then
					RSDB("ExaPerName")=Request("ExaPerName")
		     		RSDB.UpDate
				end if
			RSDB.Close
		End if
	End if
%>
<form method="post" action="ExaPeriod.asp" name="Form1">
  <Center>
    <p>
    <h2>�� �� �� ��</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="600">
  	<tr>
		<td align="right">	<input type="submit" name="submit" value="  ���  ">
    <input type="submit" name="Submit" value="  ����  ">
    <input type="submit" name="Submit" value="  ɾ��  " onClick="return confirm('�Ƿ�ȷ��ɾ�������ڣ�')"></td>
	</tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="600" bordercolordark="#FFFFFF" bordercolorlight="#999999">
    <tr bgcolor="DDDDDD" class=tdcss> 
      <td width="28%" height="25"> 
        <div align="center">�����������</div>
      </td>
      <td width="31%" height="25"> 
        <div align="center">������������</div>
      </td>
    </tr>
    <%
	RowIndex=0
	RSDB.Open "Select * From ExaPeriod order by ExaPerId", G_DBConn, 2, 3, 1
		Do While NOT RSDB.eof
		RowIndex=RowIndex+1
%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25"> 
        <div align="center"><font color="#0000FF">&nbsp; <a href="ExaPeriod.asp?EditNum=<%=RSDB("ExaPerId")%>"><%=RowIndex%></a></font></div>
      </td>
      <td height="25" align="center"> <%	If RSDB("ExaPerId")<>EditNum then	%> <font color="#0000FF">&nbsp;<%=RSDB("ExaPerName")%></font> 
        <%	Else	%> 
        <input type="text" name="ExaPerName" value="<%=RSDB("ExaPerName")%>" class=midinput>
        <input type="hidden" name="ExaPerId" value="<%=RSDB("ExaPerId")%>">
      <% 		BJStr="�༭"
	     End if %></td>
      <%
			RSDB.MoveNext
		Loop
	RSDB.Close
%> </tr>
    <%	If Submit="  ���  " then 
		BjStr="���" 	%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25" align="center"> <%
	RSDB.Open "Select Count(*) as Count From ExaPeriod ",G_DBConn,2,3,1
		Rows=RSDB("Count")+1
	RSDB.Close
	Response.Write Rows
%> 
        <div align="center"> 
          <Input type="hidden" name="ExaPerId" value="<%=ExaPerIdMax%>">
        </div>
      </td>
      <td height="25" align="center"> 
        <input type="text" name="ExaPerName" class=midinput>      </td>
    </tr>
    <%	End if 	%> 
  </table>
<p></p>
  <Center>
	<input type="hidden" name="BJStr" value="<%=BJStr%>">
  </Center>
</form>
</body>
<%
	Set RSDB=Nothing
	Set G_DBConn=Nothing
%>
</html>
