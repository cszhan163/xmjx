<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>
<body>
<%
	Set g_DBConn = Server.CreateObject("ADODB.Connection")
	g_DBConn.Open Application("ConnectionString")
	
	set rs=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	rs.open "select * from Examine ",G_DBConn,1,1,1
	do while not rs.eof
		ExamineId=rs("ExamineId")
		Num=0
		rsTemp.open"select * from ExamineItem where ExamineId="&ExamineId&" and OrderNum is null order by ExaItemId",G_DBConn,2,3,1
		do while not rsTemp.eof
			ExaItemId=rsTemp("ExaItemId")
			Num=cdbl(Num)+1
			G_DBConn.execute("update ExamineItem set OrderNum="&Num&" where ExaItemId="&ExaItemId&"")
			rsTemp.movenext
		loop
		rsTemp.close
		rs.movenext
	loop
	rs.close 
%>
<form name="form1" method="post" action="">
  <input type="submit" name="Submit" value="初始化">
</form>
</body>
</html>
