<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%ModuleCode = "CE"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����Ҫ��ά���б�</title>
</head>
<%
	Submits=request("Submits")
	SearchStr=CurSelValue("SearchStr")
	Id=CurSelValue("Id")
	FactState=CurSelValue("FactState")
	set RSDB=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	RSDB.cursorlocation=3
	
	if SearchStr<>"" and not isnull(SearchStr) then
		Query=Query&" and ExaFactorName like '%"&SearchStr&"%' "
	end if
	if FactState<>"" and not isnull(FactState) then
		if FactState<>2 then
			Query=Query&" and IsDel='"&FactState&"' "
		end if
	else
		Query=Query&" and IsDel='0'"
	end if
 		
	if Submits="  ���  " then
		response.Redirect "ExaFactorEdit.asp?Submit=New"
		response.End()
	end if
	
	if Id<>"" and not isnull(Id) then 
		if Id="ExaFactorId" then
			strOrder=" order by ExaFactorId desc "
		end if
		if Id="ExaFactorName" then
			strOrder=" order by ExaFactorName asc "
		end if
	end if
%>
<body>
<form method="post" action="ExaFactorList.asp" name="Form1">
  <Center>
    <h2>�� �� Ҫ �� ά �� �� �� </h2>
  </Center>
  <table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
  <tr>
  	<td >Ҫ������
      <input type="text" name="SearchStr" value="<%=SearchStr%>" class=input title="����Ҫ������">
      <select name="FactState" onChange="Form1.submit()">
	    <option value="0" <%if FactState="0" then response.Write("selected") end if%>>����</option>
		<option value="1" <%if FactState="1" then response.Write("selected") end if%>>����</option>
		<option value="2" <%if FactState="2" then response.Write("selected") end if%>>ȫ��</option>
      </select>  
      <input type="submit" name="Submits" value="��ѯ">  </td>
  	<td  align="right">	<input type="submit" name="Submits" value="  ���  "></td>
  </tr>
  </table>
  <table border="1" align="center" bgcolor="#FFE8D0" cellpadding="5" cellspacing="0" width="800" bordercolordark="#FFFFFF" bordercolorlight="#999999" name="EmpGrid">
  	<tr align="center" bgcolor="DDDDDD" class="tdcss">
		<td><a href="ExaFactorList.asp?Id=ExaFactorId">���</a></td>
		<td><a href="ExaFactorList.asp?Id=ExaFactorName">Ҫ������</a></td>
		<td>���˱�׼</td>
		<td>���˰취</td>
	</tr>
<%
	RSDB.open "select * from ExaFactor where 1=1 "&Query&" "&strOrder&" ",G_DBConn,1,1,1
	if Not RSDB.eof then
		RSDB.PageSize=10
  		Submits=Request("Submits")
		CurPage=Request("CurPage")
  		if CurPage="" or not IsNumeric(CurPage) then
    		CurPage=1
  		else
    		CurPage=CInt(Request("CurPage"))
  		end if
  		if CurPage<1 or CurPage>RSDB.PageCount then
    		CurPage=1
  		end if
  		if Submits="��һҳ" then
  			CurPage=1
  		end if
  		if Submits="��һҳ" and CurPage<RSDB.PageCount then
  			CurPage=CurPage+1
  		end if
  		if Submits="��һҳ" and CurPage>1 then
  			CurPage=CurPage-1
  		end if
  		if Submits="���һҳ" then
  			CurPage=RSDB.PageCount
  		end if
  		RSDB.AbsolutePage=CurPage
	end if
	RecCount=0
	do while (not RSDB.eof) and (RecCount<RSDB.PageSize)
		ExaFactorId=RSDB("ExaFactorId")
		ExaFactorName=RSDB("ExaFactorName")
		ExaNorm=RSDB("ExaNorm")
		IsDel=RSDB("IsDel")
		if ExaFactorName="" or isnull(ExaFactorName) then
			G_DBConn.execute("delete ExaFactor where ExaFactorId="&ExaFactorId&"")
			response.Redirect("ExaFactorList.asp")
			response.End()
		end if
%>
	<tr align="center" bgcolor="#FFFFFF" class="tdcss">
		<td><%		if IsDel=1 then %>
        <font color="#FF0000">*</font>
<%		end if %> <%=RecCount+1%>&nbsp;</td>
		<td><a href="ExaFactorEdit.asp?ExaFactorId=<%=ExaFactorId%>&CurPage=<%=CurPage%>"><%=ExaFactorName%>&nbsp;</a></td>
	  <td align="left" width="400"><%=ExaNorm%>&nbsp;</td>
		<td width="300" align="left"><%
			rsTemp.open "select * from ExaFactorItem where ExaFactorId="&ExaFactorId&"",G_DBConn,1,1,1
			num=0
			do while not rsTemp.eof
				num=num+1
				ExaFacItemName=rsTemp("ExaFacItemName")
				response.Write(num&". "&ExaFacItemName&"&nbsp;&nbsp;")
				rsTemp.movenext
			loop
			rsTemp.close
		%>&nbsp;</td>
	
	</tr>
<%
		RecCount=RecCount+1
		RSDB.movenext
	loop
	
%>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="800">
	<tr> 
    <td width="85%"  valign="top"> 
        <div align="center">
            <input type="hidden" name="CurPage" value="<%=CurPage%>">
            <%if CurPage>1 then%> 
            <input type="submit" name="Submits" value="��һҳ">
            <input type="submit" name="Submits" value="��һҳ">
            <%end if%> 
            <%if CurPage<RSDB.PageCount then%> 
            <input type="submit" name="Submits" value="��һҳ">
            <input type="submit" name="Submits" value="���һҳ">
            <%end if%> 
        </div>
    </td>
      <td width="13%"> 
		<div align="left"><font size="2">��<%=CurPage%>ҳ����<%=RSDB.PageCount%>ҳ</font></div>
    </td>
  </tr>
  <tr> 
    <td height="2" colspan="2"></td>
  </tr> 
</table>
<%
RSDB.close
%>
</form>
</body>
</html>
