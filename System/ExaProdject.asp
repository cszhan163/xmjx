<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "CD"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������Ŀά��</title>
</head>

<body>
<%
	Submits=Request("Submits")
	EditNum=Request("EditNum")
	SearchStr=curselvalue("SearchStr")
	ProIsDel=curselvalue("ProIsDel")
	If ISNumeric(EditNum) then
		EditNum=Cint(EditNum)
	else
		EditNum=-1
	End if
%>
<%
	if Submits="����" then
%> 
<meta http-equiv="refresh" content="0;URL=AllMana.asp">   
<%
		Response.end
	end if
	Set RSDB = Server.CreateObject("ADODB.Recordset")
    RSDb.cursorlocation=3
	Query="" 
	if SearchStr<>"" and not isnull(SearchStr) then
		Query=Query&" and ExaProdName like '%"&SearchStr&"%' "
	end if
	if ProIsDel<>"" and not isnull(ProIsDel) then
		if ProIsDel<>2 then
			Query=Query&" and IsDel='"&ProIsDel&"' "
		end if
	else
		Query=Query&" and IsDel='0' "
	end if
	
	if Submits="  ����  " then
		ExaProdIdX=Request("ExaProdId")
		If IsNumeric(ExaProdIdX) then
			ExaProdIdX=Cint(ExaProdIdX)
		else
			ExaProdIdX=-1
		end if	
		RSDB.open "select IsDel from ExaProdject where ExaProdId="&ExaProdIdX&"",G_DBConn,1,1,1
		if not RSDB.eof then
			IsDel=RSDB("IsDel")
		end if
		RSDB.close
		if IsDel=1 then
			G_DBConn.execute("update ExaProdject set IsDel=0 where ExaProdId="&ExaProdIdX&"")
		end if
	end if
	
	If Submits="  ɾ��  " then
		ExaProdIdX=Request("ExaProdId")
		If IsNumeric(ExaProdIdX) then
			ExaProdIdX=Cint(ExaProdIdX)
		else
			ExaProdIdX=-1
		end if	
		RSDB.open "select IsDel from ExaProdject where ExaProdId="&ExaProdIdX&"",G_DBConn,1,1,1
		if not RSDB.eof then
			IsDel=RSDB("IsDel")
		end if
		RSDB.close
		if IsDel=1 then
			RSDb.open "select Count(*) count from ExamineItem where ExaProdId="&ExaProdIdX&"",G_DBConn,1,1,1
			if not rsdb.eof then
				num=rsdb("count")
			end if
			RSDB.close
			if num<1 then
				RSDB.Open "Delete From ExaProdject where ExaProdId="&ExaProdIdX,G_DBConn,2,3,1
				response.Redirect("ExaProdject.asp")
				response.End()
			else
				response.Write("<script language='javascript'>alert('�Ѿ�ʹ��,���ܱ�ɾ��!')</script>")
			end if
		else
			G_DBConn.execute("update ExaProdject set IsDel=1 where ExaProdId="&ExaProdIdX&"")
		end if
	End if
	If Submits="  ����  " then
		XXStr=Request("BJStr")
		ExaProdName=Request("ExaProdName")
		RSDB.open "select Count(*) count from ExaProdject where ExaProdName='"&ExaProdName&"'",G_DBConn,1,1,1
		if not RSDB.eof then
			ProdNum=RSDB("count")
		end if
		RSDB.close
		if ProdNum>0 then
			response.Write("<script language='javascript'>alert('������ͬ���Ƶ����ݣ�');</script>")
		end if
		If XXStr="���" then
          		RSDB.Open "Select * From ExaProdject ",G_DBConn,2,3,1
				RSDB.AddNew
			    RSDB("ExaProdName")=Request("ExaProdName")
				RSDB("Remarks")=Request("Remarks")
		    	RSDB.UpDate
			RSDB.Close
		End if
		If XXStr="�༭" then
			ExaProdIdX=Request("ExaProdId")
			If IsNumeric(ExaProdIdX) then
				ExaProdIdX=Cint(ExaProdIdX)
			else
				ExaProdIdX=-1
			end if
			RSDB.Open "Select * From ExaProdject Where ExaProdId="&ExaProdIdX,G_DBConn,2,3,1
				If Not RSDB.Eof then
					RSDB("ExaProdName")=Request("ExaProdName")
					RSDB("Remarks")=Request("Remarks")
		     		RSDB.UpDate
				end if
			RSDB.Close
		End if
	End if
%>
<form method="post" action="ExaProdject.asp" name="Form1">
  <Center>
    <p>
    <h2>�� �� �� Ŀ ά �� </h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="700px">
  <tr>
  <td width="351">��Ŀ����
      <input type="text" name="SearchStr" value="<%=SearchStr%>" class=input title="������Ŀ����">
      <select name="ProIsDel" onChange="Form1.submit()">
	    <option value="0" <%if ProIsDel="0" then response.Write("selected") end if%>>����</option>
		<option value="1" <%if ProIsDel="1" then response.Write("selected") end if%>>����</option>
		<option value="2" <%if ProIsDel="2" then response.Write("selected") end if%>>ȫ��</option>
      </select>  
      <input type="submit" name="Submits" value="��ѯ">  </td>
  	<td width="349" align="right">	<input type="submit" name="Submits" value="  ���  ">
    <input type="submit" name="Submits" value="  ����  ">
    <input name="Submits" type="Submit" id="Submit" value="  ����  ">
    <input type="submit" name="Submits" value="  ɾ��  " onClick="return confirm('�Ƿ�ȷ��ɾ������Ŀ��')"></td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="700" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
    <tr bgcolor="DDDDDD" class=tdcss> 
      <td width="28%" height="25"> 
        <div align="center">������Ŀ���</div>
      </td>
      <td width="31%" height="25"> 
        <div align="center">������Ŀ����</div>
      </td>
	   <td width="31%" height="25" nowrap> 
        <div align="center">��ע</div>
      </td>
    </tr>
    <%
	RowIndex=0
	RSDB.Open "Select * From ExaProdject where 1=1 "&Query&" order by ExaProdId ", G_DBConn, 2, 3, 1
		Do While NOT RSDB.eof
		RowIndex=RowIndex+1
		IsDel=RSDB("IsDel")
%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25">
	    
        <div align="center"><font color="#0000FF">&nbsp; <%if IsDel=1 then %><font color="#FF0000">*</font><%end if %><a href="ExaProdject.asp?EditNum=<%=RSDB("ExaProdId")%>"><%=RowIndex%></a></font></div>
      </td>
      <td height="25" align="center"> <%	If RSDB("ExaProdId")<>EditNum then	%> <font color="#0000FF">&nbsp;<%=RSDB("ExaProdName")%></font> 
        <%	Else	%> 
        <input type="text" name="ExaProdName" value="<%=RSDB("ExaProdName")%>" class=midinput>
        <input type="hidden" name="ExaProdId" value="<%=RSDB("ExaProdId")%>">
      <% 		BJStr="�༭"
	     End if %></td>
	  <td height="25" align="center" width="600"> <%	If RSDB("ExaProdId")<>EditNum then	%> <font color="#0000FF">&nbsp;<%=RSDB("Remarks")%></font> 
        <%	Else	%> 
        <input name="Remarks" type="text" class=midinput value="<%=RSDB("Remarks")%>" maxlength="50">
      <% 		BJStr="�༭"
	     End if %></td>
      <%
			RSDB.MoveNext
		Loop
	RSDB.Close
%> </tr>
    <%	If Submits="  ���  " then 
		BjStr="���" 	%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25" align="center"> <%
	RSDB.Open "Select Count(*) As Count From ExaProdject ",G_DBConn,2,3,1
		Rows=RSDB("Count")+1
	RSDB.Close
	Response.Write Rows
%> 
        <div align="center"> 
          <Input type="hidden" name="ExaProdId" value="<%=ExaProdIdMax%>">
        </div>
      </td>
      <td height="25" align="center"> 
        <input type="text" name="ExaProdName" class=midinput>      </td>
	  <td height="25" align="center"> 
        <input name="Remarks" type="text" class=midinput maxlength="50">      </td>
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
