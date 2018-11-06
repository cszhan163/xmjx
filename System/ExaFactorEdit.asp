<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "CE"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>考核要素维护编辑</title>
</head>
<%
	Submit=Request("Submit")
	ExaFactorId=request("ExaFactorId")
	ExaFactorName=request("ExaFactorName")
	ExaNorm=request("ExaNorm")
	CurPage=request("CurPage")
	IsCanOver=request("IsCanOver")
	
	if Submit=" 返回 " then
%> 
<meta http-equiv="refresh" content="0;URL=ExaFactorList.asp?CurPage=<%=CurPage%>">   
<%
		Response.end
	end if
	
	Set RSDB = Server.CreateObject("ADODB.Recordset")
    RSDB.cursorlocation=3
	
	if Submit="New" then
		RSDB.open "select * from ExaFactor ",G_DBConn,2,3,1
		RSDB.addnew
			RSDB("ExaFactorName")=""
			RSDB("ExaNorm")=""
		RSDB.update
		ExaFactorId=RSDB("ExaFactorId")
		RSDB.close
	end if
	
	if Submit="添加办法" then
		RSDb.open "select Count(*) count from ExamineItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
		if not rsdb.eof then
			num=rsdb("count")
		end if
		RSDB.close
		if num<1 then
			RSDB.open "select * from ExaFactorItem where ExaFactorId="&ExaFactorId&"",G_DBConn,2,3,1
			RSDB.addnew
				RSDB("ExaFactorId")=ExaFactorId
				RSDB("ExaFacItemName")=""
			RSDB.update
			RSDB.close
		else
			response.Write("<script language='javascript'>alert('已经使用,不能被修改!');</script>")
		end if
	end if
	
	if Submit=" 保存 " then
		if ExaFactorName<>"" and not isnull(ExaFactorName) then
			RSDB.open "select * from ExaFactor where ExaFactorId="&ExaFactorId&" ",G_DBConn,2,3,1
				RSDB("ExaFactorName")=ExaFactorName
				RSDB("ExaNorm")=coder(ExaNorm)
				RSDB("IsCanOver")=IsCanOver
			RSDB.update
			RSDB.close
			
			Rows=request("AllExaFacItemId").count
			for i=1 to Rows
				CurrFacItemId=request("AllExaFacItemId")(i)
				RSDB.open "select * from ExaFactorItem where ExaFacItemId="&CurrFacItemId&"",G_DBConn,2,3,1
					RSDB("ExaFacItemName")=request("FacItemName"&CurrFacItemId)
				RSDB.update	
				RSDB.close
			next
			
			response.Redirect "ExaFactorEdit.asp?ExaFactorId="&ExaFactorId&""
			response.End()
		else
			strWarn="请将考核要素名称填写完整！"
		end if
	end if
	
	if Submit="删除要素" then
		RSDB.open "select IsDel from ExaFactor where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
		if not RSDB.eof then
			IsDel=RSDB("IsDel")
		end if
		RSDB.close
		if IsDel=1 then
			RSDb.open "select Count(*) count from ExamineItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
			if not rsdb.eof then
				num=rsdb("count")
			end if
			RSDB.close
			if num<1 then
				G_DBConn.execute("delete ExaFactor where ExaFactorId="&ExaFactorId&"")
				response.Redirect "ExaFactorList.asp"
				response.End()
			else
				response.Write("<script language='javascript'>alert('已经使用,不能被删除!')</script>")
			end if
		else
			G_DBConn.execute("update ExaFactor set IsDel=1 where ExaFactorId="&ExaFactorId&"")
			response.Redirect "ExaFactorList.asp"
			response.End()
		end if
	end if
	
	if Submit="启用要素" then
		RSDB.open "select IsDel from ExaFactor where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
		if not RSDB.eof then
			IsDel=RSDB("IsDel")
		end if
		RSDB.close
		if IsDel=1 then
			G_DBConn.execute("update ExaFactor set IsDel=0 where ExaFactorId="&ExaFactorId&"")
			response.Redirect "ExaFactorList.asp"
			response.End()
		end if
	end if
	
	if Submit="删除办法" then
		RSDb.open "select Count(*) count from ExamineItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
		if not rsdb.eof then
			num=rsdb("count")
		end if
		RSDB.close
		if num<1 then
			strFacItemId=request("ExaFacItemId")
			if strFacItemId<>"" and not isnull(strFacItemId) then
				G_DBConn.execute("delete ExaFactorItem where ExaFacItemId in ("&strFacItemId&")")
			end if
		else
			response.Write("<script language='javascript'>alert('已经使用,不能被删除!')</script>")
		end if
	end if
	
	RSDB.open "select * from ExaFactor where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
		ExaFactorName=RSDB("ExaFactorName")
		ExaNorm=RSDB("ExaNorm")
		IsCanOver=RSDB("IsCanOver")
	RSDB.close
	'检查是否有重复
	RSDB.open "select count(*) as count from ExaFactor where ExaFactorName='"&ExaFactorName&"'",G_DBConn,1,1,1
	if not RSDB.eof then
		FacNum=RSDB("count")
	end if
	RSDB.close
	
	if FacNum>1 then
		strWarn="提示：已有一条记录名称与此数据重复!"
	end if
%>
<body>
<form method="post" action="ExaFactorEdit.asp?ExaFactorId=<%=ExaFactorId%>&CurPage=<%=CurPage%>" name="Form1">
  <Center>
    <h2>考 核 要 素 维 护 编 辑</h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr>
  <td width="265"><font color="#FF0000"><%=strWarn%>&nbsp;</font></td>
  	<td width="535"  align="right">
	<input type="submit" name="submit" value=" 保存 ">
	<input type="submit" name="submit" value="删除要素" onClick="return confirm('是否确定删除此要素？')">
	<input type="submit" name="Submit" value="启用要素">
	<input type="submit" name="submit" value="添加办法">
	<input type="submit" name="submit" value="删除办法">
	<input type="submit" name="submit" value=" 返回 "></td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
  	<tr>
		<td width="121" align="center" bgcolor="DDDDDD">考核要素名称</td>
		<td width="673" bgcolor="#FFFFFF">
	  <input name="ExaFactorName" type="text" class="longinput" value="<%=ExaFactorName%>"></td>
	</tr>
	<tr>
		<td align="center" bgcolor="DDDDDD">考核要素标准</td>
		<td bgcolor="#FFFFFF"><textarea name="ExaNorm" class="mutiinput"><%=Htmlcoder(ExaNorm)%></textarea></td>
	</tr>
	<tr>
		<td align="center" bgcolor="DDDDDD">考核要素办法</td>
		<td bgcolor="#FFFFFF" width="673">
		<%
			RSDB.open "select * from ExaFactorItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
			do while not RSDB.eof
				ExaFacItemId=RSDB("ExaFacItemId")
				ExaFacItemName=RSDB("ExaFacItemName")
		%>
		<input type="checkbox" name="ExaFacItemId" value="<%=ExaFacItemId%>"><textarea name="FacItemName<%=ExaFacItemId%>" class="midtextarea" rows="3"><%=ExaFacItemName%></textarea>
		<input type="hidden" name="AllExaFacItemId" value="<%=ExaFacItemId%>">
		<%
				RSDB.movenext
			loop
			RSDB.close
		%>&nbsp;
	  </td>
	</tr>
	<tr><td bgcolor="DDDDDD">是否可大于最大分数</td><td bgcolor="#FFFFFF">
	  <input type="radio" name="IsCanOver" value="1" <%if IsCanOver=true then response.Write("checked") end if%>>
	  是
	  <input type="radio" name="IsCanOver" value="0" <%if IsCanOver=false then response.Write("checked") end if%>>
	      否</td>
	</tr>
  </table>
<input type="hidden" name="CurPage" value="<%=CurPage%>">
</form>
</body>
</html>
