<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "BC"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<base target="_self">
<title>选择要素</title>
</head>
<%
	ExamineId=request("ExamineId")
	ClientY=request("ClientY")
	Submits=request("Submits")
	Set rsMain = Server.CreateObject("ADODB.Recordset")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	
	if Submits=" 确定 " then
		Counts=request("ExaFactorId").count
		maxOrderNum=GetMaxOrder(ExamineId)
		for i=1 to Counts
			CurExaFactorId=request("ExaFactorId")(i)
			maxOrderNum=cdbl(maxOrderNum)+1
			rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&"",G_DBConn,2,3,1
			rsMain.addnew
				rsMain("ExamineId")=ExamineId
				rsMain("ExaFactorId")=CurExaFactorId
				rsMain("OrderNum")=maxOrderNum
			rsMain.update
			rsMain.close
		next
%>
	<script language="javascript">
		window.returnValue = <%=ClientY%>; 
		window.dialogArguments.location.reload("ExamineEdit.asp"+"?Window_OffsetY="+window.returnValue+"&ExamineId="+<%=ExamineId%>); 
		window.close();
</script>
<%
	end if
%>
<body>
<form name="form1" method="post" action="SelFactor.asp?ClientY=<%=ClientY%>&ExamineId=<%=ExamineId%>">
<br>
<Center>
    <h2>选 择 要 素</h2>
</Center>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="600">
<tr valign="bottom">
<td align="right">
  <input name="Submits" type="submit" id="Submits" value=" 确定 "></td>
</tr>
</table>
<table border="1" align="center" width="600" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="0" cellspacing="0" style=" margin-top:5px; margin-bottom:5px;" >
<%
	rsMain.open "select * from ExaFactor where IsDel=0 and ExaFactorId not in "&_
		"(select ExaFactorId from ExamineItem where ExamineId="&ExamineId&") order by ExaFactorName ",G_DBConn,1,1,1
	Num=rsMain.recordcount
	if not rsMain.eof then
		if (Num Mod 2)=0 then
			Rows=cdbl(Num)/2
		else
			Rows=cdbl(Num)/2+1
		end if
		for i=0 to Rows-1
%>
<tr>
<%
			for j=0 to 1
				if (i*2+j)<num then
				ExaFactorName=rsMain("ExaFactorName")
				ExaFactorId=rsMain("ExaFactorId")
%>
<td bgcolor="DDDDDD"><input name="ExaFactorId" type="checkbox" value="<%=ExaFactorId%>"></td>
<td bgcolor="#FFFFFF"><%=ExaFactorName%>&nbsp;</td>
<%
				
					rsMain.movenext
				end if
			next
%>
</tr>
<%
		next
	end if
	rsMain.close
%>
</table>
<input type="hidden" name="txt0" value="<%=ClientY%>">
<input type="hidden" name="ExaId" value="<%=ExamineId%>">
</form>
<script language="javascript">
//	function PageClose(ExamineId)
//	{
//		window.returnValue = document.getElementById("txt0").value; 
//		window.dialogArguments.location.reload("ExamineEdit.asp"+"?Window_OffsetY="+window.returnValue+"&ExamineId="+ExamineId); 
//		window.close();
//	}
</script>
</body>
</html>
