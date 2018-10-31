<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "CF"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>分值种类编辑</title>
</head>
<%
	Submits=request("Submits")
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	
	if Submits=" 添加 " then
		rsMain.open "select * from ScoreKind ",G_DBConn,2,3,1
		rsMain.addnew
			rsMain("Weighing")=0
		rsMain.update
		rsMain.close
		Response.Redirect("ScoreKind.asp")
		Response.End()
	end if
	if Submits=" 保存 " then
		Rows=request("AllScoreKindId").count
		for i=1 to Rows 
			CurScoreKindId=request("AllScoreKindId")(i)
			rsMain.open "select * from ScoreKind where ScoreKindId="&CurScoreKindId&"",G_DBConn,2,3,1
				rsMain("Weighing")=request("Weighing"&CurScoreKindId)
				rsMain("SKName")=request("SKName"&CurScoreKindId)
				rsMain("Max1")=request("Max1"&CurScoreKindId)
				rsMain("Max2")=request("Max2"&CurScoreKindId)
				rsMain("Max3")=request("Max3"&CurScoreKindId)
				rsMain("Max4")=request("Max4"&CurScoreKindId)
				rsMain("Max5")=request("Max5"&CurScoreKindId)
				rsMain("Min1")=request("Min1"&CurScoreKindId)
				rsMain("Min2")=request("Min2"&CurScoreKindId)
				rsMain("Min3")=request("Min3"&CurScoreKindId)
				rsMain("Min4")=request("Min4"&CurScoreKindId)
				rsMain("Min5")=request("Min5"&CurScoreKindId)
			rsMain.update
			rsMain.close
		next
	end if
	if Submits=" 删除 " then
		strID=request("ScoreKindId")
		if strID<>"" and not isnull(strID) then
			G_DBConn.execute("Delete ScoreKind where ScoreKindId in ("&strID&")")
		end if
	end if
%>
<body>
<h2>分值种类编辑</h2>
<form name="form1" method="post" action="">
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td align="right">
  <input name="Submits" type="submit" id="Submits" value=" 添加 ">
  <input name="Submits" type="submit" id="Submits" value=" 保存 ">
  <input name="Submits" type="submit" id="Submits" value=" 删除 ">
</td>
</tr>
</table>
<table width="80%" border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999">
  <tr>
    <td width="9%" align="center" bgcolor="DDDDDD" class="tdcss">序号</td>
	<td width="9%" align="center" bgcolor="DDDDDD" class="tdcss">名称</td>
    <td width="13%" align="center" bgcolor="DDDDDD" class="tdcss">权重</td>
    <td width="18%" align="center" bgcolor="DDDDDD" class="tdcss">分值1</td>
    <td width="15%" align="center" bgcolor="DDDDDD" class="tdcss">分值2</td>
    <td width="15%" align="center" bgcolor="DDDDDD" class="tdcss">分值3</td>
    <td width="15%" align="center" bgcolor="DDDDDD" class="tdcss">分值4</td>
    <td width="15%" align="center" bgcolor="DDDDDD" class="tdcss">分值5</td>
  </tr>
<%
	RowIndex=0
	rsMain.open "select * from ScoreKind ",G_DBConn,1,1,1
	do while not rsMain.eof
		RowIndex=cdbl(RowIndex)+1
		ScoreKindId=rsMain("ScoreKindId")
		Weighing=rsMain("Weighing")
		SKName=rsMain("SKName")
		Max1=rsMain("Max1")
		Max2=rsMain("Max2")
		Max3=rsMain("Max3")
		Max4=rsMain("Max4")
		Max5=rsMain("Max5")
		Min1=rsMain("Min1")
		Min2=rsMain("Min2")
		Min3=rsMain("Min3")
		Min4=rsMain("Min4")
		Min5=rsMain("Min5")
%>
  <tr bgcolor="#FFFFFF">
    <td align="center" bgcolor="#FFFFFF">
	<input type="checkbox" name="ScoreKindId" value="<%=ScoreKindId%>"><%=RowIndex%>
	<input type="hidden" name="AllScoreKindId" value="<%=ScoreKindId%>"></td>
	<td>
		<input name="SKName<%=ScoreKindId%>" type="text" class="priceinput" value="<%=SKName%>">
	</td>
    <td align="center" bgcolor="#FFFFFF">
	<input name="Weighing<%=ScoreKindId%>" type="text" class="priceinput" value="<%=Weighing%>" onChange="CheckNum('Weighing',<%=ScoreKindId%>)">%</td>
    <td align="center" bgcolor="#FFFFFF">
	<input name="Max1<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Max1%>" onChange="CheckNum('Max1',<%=ScoreKindId%>)">-<input name="Min1<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Min1%>" onChange="CheckNum('Min1',<%=ScoreKindId%>)"></td>
    <td align="center" bgcolor="#FFFFFF"><input name="Max2<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Max2%>" onChange="CheckNum('Max2',<%=ScoreKindId%>)">-<input name="Min2<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Min2%>" onChange="CheckNum('Min2',<%=ScoreKindId%>)"></td>
    <td align="center" bgcolor="#FFFFFF"><input name="Max3<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Max3%>" onChange="CheckNum('Max3',<%=ScoreKindId%>)">-<input name="Min3<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Min3%>" onChange="CheckNum('Min3',<%=ScoreKindId%>)"></td>
    <td align="center" bgcolor="#FFFFFF"><input name="Max4<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Max4%>" onChange="CheckNum('Max4',<%=ScoreKindId%>)">-<input name="Min4<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Min4%>" onChange="CheckNum('Min4',<%=ScoreKindId%>)"></td>
    <td align="center" bgcolor="#FFFFFF"><input name="Max5<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Max5%>" onChange="CheckNum('Max5',<%=ScoreKindId%>)">-<input name="Min5<%=ScoreKindId%>" type="text" class="scoreinput" value="<%=Min5%>" onChange="CheckNum('Min5',<%=ScoreKindId%>)"></td>
  </tr>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
</table>
</form>
<script language="vbscript">
sub CheckNum(Str,ScoreKindId)
	ScodeValue=eval("Form1."&Str&ScoreKindId&".value")
	if not IsNumeric(ScodeValue) then
		eval("Form1."&Str&ScoreKindId).value=0
		alert("请填写正确的数字格式！")
	end if
end sub
</script>
</body>
</html>
