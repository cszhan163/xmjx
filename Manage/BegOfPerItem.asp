<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BA"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ϸ�ڼ�����</title>
</head>
<%
	Submit=Request("Submit")
	EditNum=Request("EditNum")
	BOPId=request("BOPId")
	If ISNumeric(EditNum) then
		EditNum=Cint(EditNum)
	else
		EditNum=-1
	End if
%>
<%
	if Submit=" ���� " then
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		rsTemp.open "select BOPYear from BegOfPer where BOPId="&BOPId&"",G_DBConn,1,1,1
		if not rsTemp.eof then
			UrlYear=datepart("yyyy",rsTemp("BOPYear"))
		end if
		rsTemp.close
%> 
<meta http-equiv="refresh" content="0;URL=BegOfPerEdit.asp?UrlYear=<%=UrlYear%>">   
<%
		Response.end
	end if
	Set RSDB = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
    RSDb.cursorlocation=3

	If Submit=" ɾ�� " then
		BOPItemIdX=Request("BOPItemId")
		If IsNumeric(BOPItemIdX) then
			BOPItemIdX=Cint(BOPItemIdX)
		else
			BOPItemIdX=-1
		end if
		rsTemp.open "select count(*) as count from Examine where BOPItemId="&BOPItemIdX&"",G_DBConn,1,1,1
		if not rsTemp.eof then
			num=rsTemp("count")
		end if
		rsTemp.close
		if num<1 then
			RSDB.Open "Delete From BegOfPerItem where BOPItemId="&BOPItemIdX,G_DBConn,2,3,1
		else
			response.Write("<script language='javascript'>alert('�Ѿ�ʹ��,���ܱ�ɾ��!')</script>")
		end if
	End if
	If Submit=" ���� " then
		XXStr=Request("BJStr")
		If XXStr="���" then
          		RSDB.Open "Select * From BegOfPerItem ",G_DBConn,2,3,1
				RSDB.AddNew
			    RSDB("BOPIName")=Request("BOPIName")
				RSDB("BOPId")=request("BOPId")
		    	RSDB.UpDate
			RSDB.Close
		End if
		If XXStr="�༭" then
			BOPItemIdX=Request("BOPItemId")
			If IsNumeric(BOPItemIdX) then
				BOPItemIdX=Cint(BOPItemIdX)
			else
				BOPItemIdX=-1
			end if
			RSDB.Open "Select * From BegOfPerItem Where BOPItemId="&BOPItemIdX,G_DBConn,2,3,1
				If Not RSDB.Eof then
					RSDB("BOPIName")=Request("BOPIName")
		     		RSDB.UpDate
				end if
			RSDB.Close
		End if
	End if
%>
<form method="post" action="BegOfPerItem.asp?BOPId=<%=BOPId%>" name="Form1">
  <Center>
    <p>
    <h2>�� �� �� ��</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="600">
  	<tr>
	  <td align="right">	<input type="submit" name="submit" value=" ��� ">
    <input type="submit" name="Submit" value=" ���� ">
    <input type="submit" name="Submit" value=" ɾ�� " onClick="return confirm('�Ƿ�ȷ��ɾ����Ҫ�أ�')">
    <input type="submit" name="Submit" value=" ���� "></td>
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
	RSDB.Open "Select * From BegOfPerItem where BOPId="&BOPId&" order by BOPItemId", G_DBConn, 2, 3, 1
		Do While NOT RSDB.eof
		RowIndex=RowIndex+1
%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25"> 
        <div align="center"><font color="#0000FF">&nbsp; <a href="BegOfPerItem.asp?EditNum=<%=RSDB("BOPItemId")%>&BOPId=<%=BOPId%>"><%=RowIndex%></a></font></div>
      </td>
      <td height="25" align="center"> <%	If RSDB("BOPItemId")<>EditNum then	%> <font color="#0000FF">&nbsp;<%=RSDB("BOPIName")%></font> 
        <%	Else	%> 
        <input type="text" name="BOPIName" value="<%=RSDB("BOPIName")%>" class=midinput>
        <input type="hidden" name="BOPItemId" value="<%=RSDB("BOPItemId")%>">
      <% 		BJStr="�༭"
	     End if %></td>
      <%
			RSDB.MoveNext
		Loop
	RSDB.Close
%> </tr>
    <%	If Submit=" ��� " then 
		BjStr="���" 	%> 
    <tr bgcolor="#FFFFFF" class=tdcss> 
      <td height="25" align="center"> <%
	RSDB.Open "Select Count(*) as Count From BegOfPerItem where BOPId="&BOPId&" ",G_DBConn,2,3,1
		Rows=RSDB("Count")+1
	RSDB.Close
	Response.Write Rows
%> 
        <div align="center"> 
          <Input type="hidden" name="BOPItemId" value="<%=BOPItemIdMax%>">
        </div>
      </td>
      <td height="25" align="center"> 
        <input type="text" name="BOPIName" class=midinput>      </td>
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
</body>
</html>
