<%@ LANGUAGE = VBScript %>
<%ModuleCode = "CB"%>
<html>
<head>
<title>ְԱ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func_Censor.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet> 
<%
	Submit=Request("Submit")
	SearchStr=curselvalue("SearchStr")
	UserState=curselvalue("UserState")
%>
<%
	if Submit="����ְԱ" then
%>
    <meta http-equiv="refresh" content="0;URL=EmployeeEdit.asp?EmpId=-2">
<%
		Response.end
	end if
%>
<%
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	Set RSTemp= Server.CreateObject("ADODB.RecordSet")
%>
<body background="Images/gback.jpg">
<form name="ToSearch" method="post" > 
<center><h2>ְԱ��Ϣ�б�</h2></center>
  <table border="0" width="760" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td width="50%">
<%
	EmpCount=0
	RSTemp.Open "Select Count(*) as EmpCount from Employee where IsDel=0 and Isadmin=0",G_DBConn, 0, 1, 1
	if RSTemp("EmpCount")<>"" and not isNull(RSTemp("EmpCount")) then
		EmpCount=RSTemp("EmpCount")
	end if
	RSTemp.Close
%>
        <input type="submit" name="Submit" value="����ְԱ">

      </td> 
      <td width="50%" align="right">
        ��&nbsp;�� 
        <input type="text" name="SearchStr" class=shortinput value="<%=SearchStr%>">
		<select name="UserState" onChange="ToSearch.submit()">
		<option value="1" <%if UserState="1" then response.Write("selected") end if%>>��ְ</option>
		<option value="2" <%if UserState="2" then response.Write("selected") end if%>>��ְ</option>
		<option value="0" <%if UserState="0" then response.Write("selected") end if%>>ȫ��</option>
		</select>  
        <input type="submit" name="Submit" value="��ѯ">  
      </td>  
    </tr>  
  </table>  
  <table border="1" align="center" name="EmpGrid" bgcolor="#DDDDDD" width="760" cellspacing="0" cellpadding="0" bordercolorlight="#999999" bordercolordark="#FFFFFF">
    <tr align="center" class=tdcss> 
      <td nowrap>���</td>
      <td nowrap height="30">ְԱ����</td>
      <td nowrap height="30">ְԱӢ����</td>
      <td nowrap height="30">ְԱ������</td>
      <!--<td nowrap height="30">���Ŵ���</td>-->
      <td nowrap height="30">�� &nbsp; &nbsp;��</td>
      <td nowrap height="30">��������</td>
      <td nowrap height="30">����˾ʱ��</td>
    </tr>
<%
	if UserState=1 then
		Query=" and IsDel=0 "
	elseif UserState=2 then
		Query=" and IsDel=1 "
	elseif UserState="" or isnull(UserState) then
		Query=" and IsDel=0 "
	end if
	if SearchStr="" then
  		RSEmp.Open "SELECT A.*,B.DeptName FROM Employee A,Dept B where A.DeptCode*=B.DeptCode and IsDel<>2 "&Query&" ORDER BY A.DEPTCode,A.EmpId", G_DBConn,2,3,1
	else
  		RSEmp.Open "SELECT A.*,B.DeptName FROM Employee A,Dept B WHERE A.DeptCode*=B.DeptCode and (A.EmpCode LIKE '%"&SearchStr&"%' OR A.EmpNameEng LIKE '%"&SearchStr&"%' OR A.EmpNameChs LIKE '%"&SearchStr&"%' or A.DeptCode='"&SearchStr&"') and IsDel<>2 "&Query&" ORDER BY A.EMPID", G_DBConn, 2, 3, 0
	end if
%>
<%  
	ListNo=0
    Do While NOT RSEmp.eof  
		ListNo=ListNo+1
%> 
    <tr bgcolor="#FFFFFF" class=tdcss align="center"> 
	  <td><a href="EmployeeEdit.asp?EmpId=<%=RSEmp("EmpId")%>"><%=listNo%></a></td>
      <td> 
<%		if RsEmp("IsDel")=1 then %>
        <font color="#FF0000">*</font>
<%		end if %> 
        <a href="EmployeeEdit.asp?EmpId=<%=RSEmp("EmpId")%>"><%=RSEmp("EmpCode")%></a> 
      </td>
      <td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("EmpNameEng")%></font>
      </td>
      <td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("EmpNameChs")%></font>
      </td>
      <!--<td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("DeptCode")%></font>
      </td>-->
      <td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("DeptName")%></font>
      </td>
      <td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("BirthDate")%></font>
      </td>
      <td height="25">&nbsp
        <font color="#3333CC" size="2"><%=RSEmp("HireDate")%></font>
      </td>
    </tr>
<%  
  		RSEmp.MoveNext  
	Loop  
RsEmp.close
%> 
  </table>    
</form>  
<%
  set RSEmp=nothing
  Set G_DBConn=Nothing
%>
</body>
</html>