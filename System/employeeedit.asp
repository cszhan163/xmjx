<%@ LANGUAGE = VBScript %>
<%ModuleCode = "CB"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ְԱ����༭</title>
</head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<!--#include file = "../secret/upload.asp"-->
<link href="/secret/style.css" type=text/css rel=stylesheet> 
<%
	EmpID=Request("EmpID")
	Submit=Request("Submit")
%>
<%
	Set RSEmp = Server.CreateObject("ADODB.Recordset")
	Set RSDept = Server.CreateObject("ADODB.Recordset")
	
	'ȡ�õ�ǰ�û��Ĺ���Ա���
	RSEmp.Open "SELECT IsAdmin FROM Employee WHERE EmpCode = '"& UserId &"'", G_DBConn, 0, 1, 1
		if not RSEmp.EOF then
			UserIsAdmin = RSEmp("IsAdmin")
		end if
	RSEmp.Close 

%>
<%
	if Submit="����" then
%>
	<meta http-equiv="refresh" content="0;URL=Employeelist.asp">
<%
  		Response.end
	end if

	If Submit = "��ϸȨ��" Then
		Response.Redirect "EmployeeRightList.asp?EmpID="& EmpID &"&IsView=1"
  		Response.End
	End If
	
	if Submit = "����Ȩ��" then
		Response.Redirect "EmployeeDateRight.asp?EmpID="& EmpID
  		Response.End
	end if

	
	If Submit="ɾ����Ƭ" then 
		G_DBConn.execute "Update Employee set Photo=Null where empId='"&EmpId&"'"
		response.redirect "employeeEdit.asp?EmpId="&EmpId
		response.end
	end if

	if Submit="��������" then
		G_DBConn.execute "Update Employee set IsDel=0 where empId='"&EmpId&"'"
		response.redirect "employeeEdit.asp?EmpId="&EmpId
		response.end
	end if

	if Submit="ɾ��" then
		'�����һ��ɾ������IsDelΪ�棬���IsDelΪ��������ɾ��
		RSEmp.Open "Select IsDel,Isadmin From Employee Where EmpId='"&EmpID&"'",G_DBConn,2,3,1
			if not RSEmp.eof then
				if RSEmp("Isadmin") then 
					ErrMsg("������Ա�����ܱ�ɾ����")
				else 
					if RSEmp("IsDel")=1 then
						'RSEmp.Delete
						RSEmp("IsDel")=2
						RSEmp.UpDate
					else
						RSEmp("IsDel")=1
						RSEmp.UpDate
					end if
				end if
			end if
		RSEmp.Close
%>
   		<meta http-equiv="refresh" content="0;URL=Employeelist.asp">
<%
   		Response.End
	end if


	if Submit = "" then
	     Server.ScriptTimeOut=999999
	     set UF = new UpFile_Class
	     UF.GetDate()
	     for each FileName in UF.File		'���û��ϴ���ͼƬ�ļ��󣬻�ִ�и�ѭ��
			set File = UF.File(FileName)
     
			RSEmp.Open "SELECT Photo FROM Employee WHERE EmpId = '"& EmpId &"'", g_dBConn, 1, 3, 1
			if not RSEmp.EOF then
				RSEmp("Photo").AppendChunk File.FileData()
				RSEmp.Update
				
				Randomize		'��ʾ������ͼƬ
				RandValue = Rnd
			end if
			RSEmp.Close	
	    next
	end if
	
	
	if Submit="����" then
		EmpCode=Request("EmpCode")
		PassWord=Request("PassWord")
		ConfirmPWD=Request("ConfirmPWD")
		EmpNameEng=Request("EmpNameEng")
		EmpNameChs=Request("EmpNameChs")
		DeptCode=Request("DeptCode")
		EmpEmail=Request("EmpEmail")
		EmpPhone=Request("EmpPhone")
		EmpFax=Request("EmpFax")
		BirthDate=Request("BirthDate")
		aResume=Request("aResume")
		Grade=CDBL(Request("Grade"))
		MainPageCode=Request("MainPageCode")
		HireDate=Request("HireDate")
		if EmpCode="" then               
				StrErr=StrErr&"<center><font color=red>������ְԱ����!</font></center>"
				ErrEmp=1
		end if
		If EmpCode<>"" and EmpId="-2" Then
		RSEmp.Open "Select * from Employee Where EmpCode='"&EmpCode&"'",G_DBConn,2,3,1
			if not RSEmp.eof  then				
				StrErr=StrErr&"<center><font size=4 color=red>��ְԱ�����Ѿ�����</font></center>"
				EmpCode=""
			end if
		RSEmp.Close
		End If
			if PassWord="" or ConfirmPWD="" or ConfirmPWD<>PassWord then 							
					StrErr=StrErr&"<font color=red>����ȷ���롰���롱�͡�����ȷ�ϡ�</font>"
			end if 
			if Request("EmpNameEng")="" or Request("EmpNameChs")="" then               
				StrErr=StrErr&"<center><font color=red>ְԱ�С�Ӣ������Ϊ�գ�������¼��!</font></center>"
			end if	  			  			
  			if BirthDate<>"" then 
				if not IsDate(BirthDate) then					
					'StrErr="<center><font color=red>����ȷ����'��������'���ڸ�ʽ��</font></center>"				
				end if 			
  			end if  			
  			if HireDate<>"" then 
				if not IsDate(HireDate) then					
					'StrErr="<center><font color=red>����ȷ����'����˾����'���ڸ�ʽ��</font></center>"				
				end if 			
  			end if
		If 	StrErr="" Then			
   		RSEmp.CursorLocation=3		
   		RSEmp.Open "SELECT * FROM employee WHERE EmpId='"&EmpId&"'", G_DBConn,3,3,1
			if EmpId="-2" then
				RSEmp.AddNew
				RSEmp("EmpCode")=EmpCode				
			else
				if RSEmp.eof then
					Response.Write "<center><font color=red>�������ݴ���������!</font></center>"
               		Response.WRite "<body onclick='history.back()'>"
					Response.End
				end if
			end if
			
			if Password<>"********" then 
					RSEmp("PassWord")=md5(Password)
			end if 
			if HireDate<>"" and IsDate(HireDate) then 				
    			RSEmp("HireDate")=CDate(HireDate)				
			else 
				RSEmp("HireDate")=Null
  			end if
			if BirthDate<>"" and IsDate(BirthDate) then 				
    			RSEmp("BirthDate")=CDate(BirthDate)				
			else 
				RSEmp("BirthDate")=Null
  			end if
			RSEmp("EmpNameEng")=EmpNameEng
  			RSEmp("EmpNameChs")=EmpNameChs
  			RSEmp("DeptCode")=DeptCode	
  			RSEmp("EmpEmail")=EmpEmail
  			RSEmp("EmpPhone")=EmpPhone
  			RSEmp("EmpFax")=EmpFax
  			RSEmp("Resume")=aResume   			
  			RSEmp("Grade")=Grade
			RSEmp("MainPageCode") =MainPageCode
  			RSEmp.Update
  			EmpID=RSEmp("EmpID")  	
		
					'�����û������Ϣ
			if UserIsAdmin = True then
				RSDept.Open "SELECT GroupCode, GroupName FROM EmployeeGroup", G_DBConn, 0, 1, 1
				do while not RSDept.EOF
					for each Group in Request("GroupCode")
						if RSDept("GroupCode") = Group then
							Finded = 1
							exit for
						end if
					next
					
					if Finded = 1 then
						Sql = "IF NOT EXISTS(SELECT * FROM EmployeeRole WHERE GroupCode = '"& Valid(RSDept("GroupCode")) &"' AND EmpCode = '"& Valid(RSEmp("EmpCode")) &"') "&_
							  "INSERT INTO EmployeeRole(EmpCode, GroupCode) VALUES('"& Valid(RSEmp("EmpCode")) &"', '"& Valid(RSDept("GroupCode")) &"')"
						G_DBConn.Execute Sql
					else
						G_DBConn.Execute "DELETE FROM EmployeeRole WHERE EmpCode = '"& Valid(RSEmp("EmpCode")) &"' AND GroupCode = '"& Valid(RSDept("GroupCode")) &"'"
					end if
					RSDept.MoveNext
					Finded = 0
				loop
				RSDept.Close 
			end if
			RSEmp.Close
		Else
	Response.Write(StrErr)
	End If	

 	end if
%>
<%
	'��ȡ��ʾ��Ϣ
if EmpId<>"-2" then
  	RSEmp.Open "SELECT * FROM employee WHERE EmpID='"&EmpID&"' and IsDel<>2", G_DBConn, 2, 3, 1
		if RSEmp.eof then
			Response.Write "<center><font size=5>������ȷ��ʾ��Ϣ</font></center>"
		  	Response.end
  		end if
  		EmpCode=RSEmp("EmpCode")		
  		EmpNameEng=RSEmp("EmpNameEng")
  		EmpNameChs=RSEmp("EmpNameChs")
  		DeptCode=RSEmp("DeptCode")
  		BirthDate=RSEmp("BirthDate")
  		HireDate=RSEmp("HireDate")
  		Grade=RSEmp("Grade")
  		PassWord=RSEmp("PassWord")
  		EmpEmail=RSEmp("EmpEmail")
  		EmpPhone=RSEmp("EmpPhone")
  		EmpFax=RSEmp("EmpFax")
		IsAdmin=RSEmp("IsAdmin")
		aResume=RSEmp("Resume")
        IsDel=RSEmp("IsDel")
		MainPageCode = RSEmp("MainPageCode")
		ModuleRight = RSEmp("ModuleRight")
		DenyModuleRight = RSEmp("DenyModuleRight")
  	RSEmp.Close
end if
%> 
<body background="Images/gback.jpg">
<form method="post" name="Qemp"  action="EmployeeEdit.asp" >
<center><img src="images/empEdit.gif" width="236" height="24"></center>
<table border="1" align="center" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" bgcolor="#FFFFFF"  width="740">
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD" height="30">ְԱ����<%=EmphasisTag%></td>
    <td bgcolor="#FFFFFF"> 
      <input type="hidden" name="EmpId" size="20" value="<%=EmpId%>">
<%		if EmpCode=""  or  (StrErr <>"" and ErrEmp=1 ) then	%> 
      <input type="text" name="EmpCode" size="20" value="<%=EmpCode%>" maxlength="8" class=input>
<%		else 
			Response.Write Empcode
%>
<input type="hidden" name="EmpCode" size="20" value="<%=Empcode%>">
<%			
		end if
%>
    </td>
	<td nowrap align="center"> 
		<%if EmpId <> "-2" and UserIsAdmin = True then%><a href="EmployeeGroupList.asp?EmpId=<%=EmpId%>"><b>�û����</b></a><%else%>�û����<%end if%>
	</td>
    <td rowspan="12" bgcolor="#FFFFFF" align="center"> 
       <img src="showphoto.asp?SelStr=Photo&TabStr=employee&FldStr=EmpID&ValStr=<%=EmpID%>&RandValue=<%=RandValue%>" width="207" height="181">
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">�� &nbsp; &nbsp;��</td>
    <td bgcolor="#FFFFFF"> 
      <input type="password" name="PassWord" size="8" value="********" class=input maxlength="8">
    </td>
    <td rowspan="11" nowrap valign="middle"> 
				<div nowrap style="width:170; height:100%; overflow:auto" <%if EmpId = "-2" then%>disabled<%end if%>>
<%
	if IsAdmin = True then					'��ʾϵͳ����Ա��Ϣ
%>
					<menu><li>ϵͳ����Ա</li></menu>
<%
	else									'��ʾһ���û���Ϣ
		if UserIsAdmin = True then
			Sql = "SELECT A.GroupId, A.GroupCode, A.GroupName, R.EmpCode FROM EmployeeGroup A LEFT JOIN EmployeeRole R ON A.GroupCode = R.GroupCode AND R.EmpCode = '"& Valid(EmpCode) &"' "&_
				  "ORDER BY A.GroupCode ASC"
		else
			Sql = "SELECT G.GroupId, G.GroupName FROM EmployeeRole A LEFT JOIN EmployeeGroup G ON A.GroupCode = G.GroupCode WHERE A.EmpCode = '"& Valid(EmpCode) &"' ORDER BY A.GroupCode ASC"
			Response.Write "<dir>"
		end if
		RSEmp.Open Sql, G_DBConn, 0, 1, 1
		do while not RSEmp.EOF 
			if UserIsAdmin = True then
%>
					<input type="checkbox" id="<%=RSEmp("GroupId")%>" name="GroupCode" value="<%=RSEmp("GroupCode")%>" <%If RSEmp("EmpCode") <> "" Then%>checked<%end if%>><label for="<%=RSEmp("GroupId")%>"><%=RSEmp("GroupName")%></label><br>
<%
			else
%>
					<li><label><%=RSEmp("GroupName")%></label></li><br>
<%
			end if
			RSEmp.MoveNext
		loop
		RSEmp.Close
		if UserIsAdmin = False then
			Response.Write "</dir>"
		end if
	end if
%>
				</div>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">����ȷ��</td>
    <td bgcolor="#FFFFFF"> 
      <input type="password" name="ConfirmPWD" size="8" value="********" class=input maxlength="8">
    </td>
  </tr>
  <tr class=tdcss> 
    <td bgcolor="DDDDDD" align="center">ְԱӢ����<%=EmphasisTag%></td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="EmpNameEng" size="20" value="<%=EmpNameEng%>" maxlength="50" class=midinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">ְԱ������<%=EmphasisTag%></td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="EmpNameChs" size="20" value="<%=EmpNameChs%>" maxlength="12" class=midinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">��������
    </td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="BirthDate" size="20" value="<%=BirthDate%>" class=shortinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">����˾����</td>
      <td bgcolor="#FFFFFF"> 
        <input type="text" name="HireDate" size="20" value="<%=HireDate%>" class=shortinput>
      </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">��������</td>
    <td bgcolor="#FFFFFF">
      <select size="1" name="DeptCode" width=50>
        <option value="">---ѡ����---</option>
<%
     RSDept.open "Select * from Dept ",G_DBConn,2,3,1
		Do While NOT RSDept.eof
%> 
          <option value="<%=RSDept("DeptCode")%>"<%if DeptCode=RSDept("DeptCode") then Response.write "selected"%>><%=RSDept("DeptName")%></option>
<%
			RSDept.MoveNext
        Loop
	RSDept.Close
%> 
      </select>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">��ϵ�绰</td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="EmpPhone" size="20" maxlength="20" value="<%=EmpPhone%>" class=midinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">����</td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="EmpFax" size="20" maxlength="20" value="<%=EmpFax%>" class=midinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">E_mail</td>
    <td bgcolor="#FFFFFF"> 
      <input type="text" name="EmpEmail" size="30" maxlength="50" value="<%=EmpEmail%>" class=midinput>
    </td>
  </tr>
  <tr class=tdcss> 
    <td align="center" bgcolor="DDDDDD">ϵͳ��ҳ</td>
    <td bgcolor="#FFFFFF"> 
      <select name="MainPageCode" style="width:150px">
<%
	RSDept.Open "SELECT A.ModuleCode, (CASE WHEN LEFT(A.ModuleName, 1) = 'S' THEN RIGHT(A.ModuleName, LEN(A.ModuleName) -1) ELSE A.ModuleName END) ModuleName "&_
				"FROM Sys_Module A WHERE "&_
				"NOT EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
				"			WHERE R.EmpCode = '"& EmpCode &"' AND ISNULL(G.DenyModuleRight, '') LIKE '%'+ A.ModuleCode +',%') "&_
				"		AND '"& DenyModuleRight &"' NOT LIKE '%'+ A.ModuleCode +',%' "&_
				"		AND (EXISTS(SELECT * FROM EmployeeRole R LEFT JOIN EmployeeGroup G ON R.GroupCode = G.GroupCode "&_
				"			WHERE R.EmpCode = '"& EmpCode &"' AND ISNULL(G.ModuleRight, '') LIKE '%'+ A.ModuleCode +',%') "&_
				"			OR '"& ModuleRight &"' LIKE '%'+ A.ModuleCode +',%') "&_
				"ORDER BY A.MenuPos, A.MenuItemPos ASC", G_DBConn, 0, 1, 1
	do while not RSDept.EOF
%>
					<option value="<%=RSDept("ModuleCode")%>" <%if MainPageCode = RSDept("ModuleCode") then%>selected<%end if%>><%=RSDept("ModuleName")%></option>
<%
		RSDept.MoveNext
	loop
	RSDept.Close 
%>
	</select>
    </td>
  </tr>
  <tr>
    <td align="center" bgcolor="#DDDDDD">���˷�ʽ</td>
    <td colspan="3"><%
		RSEmp.open "select distinct EP.ExaPerName from Examine E "&_
			"left join BegOfPerItem BPI on BPI.BOPItemId=E.BOPItemId "&_
			"left join BegOfPer BP on BP.BOPId=BPI.BOPId "&_
			"left join ExaPeriod EP on EP.ExaPerId=BP.ExaPerId "&_
			"where E.ExaObjType=3 and E.ExaObjCode='"&EmpCode&"' ",G_DBConn,1,1,1
		do while not RSEmp.eof 
			response.Write(RSEmp("ExaPerName")&"&nbsp;&nbsp;")
			RSEmp.movenext
		loop
		RSEmp.close
	%>&nbsp;</td></tr>
</table>
<table border="1" align="center" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" bgcolor="#FFFFFF"  width="740">
  <tr class=tdcss>
    <td align="center" bgcolor="DDDDDD"> 
      <div align="right">�� &nbsp; &nbsp;��</div>
    </td>
    <td bgcolor="#FFFFFF"> 
      <textarea rows="5" name="aResume" cols="76" class=multiinput><%=aResume%></textarea>
    </td>
  </tr>
</table>
<table border="1" align="center" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" bgcolor="#FFFFFF"  width="740">
  <tr bgcolor="#FFFFFF"> 
    <td align="center"> 
      <input type="Submit" name="Submit" value="����">
      <%if UserIsAdmin = True and IsAdmin = False and IsDel=0 then%><input type="Submit" name="Submit" value="ɾ��"><%end if%>
	  <%if UserIsAdmin = True and IsAdmin = False and IsDel=1 then%><input type="submit" name="Submit" value="��������"><%end if%>
	  <%if EmpId <> "-2" then%>
		<input type="Submit" name="Submit" value="��ϸȨ��">
		<%if UserIsAdmin then%><input type="Submit" name="Submit" value="����Ȩ��"><%end if%>
      <%end if%>
	  <input type="submit" name="Submit" value="ɾ����Ƭ">
      <input type="Submit" name="Submit" value="����">
    </td>
  </tr>
</table>
</form>
<%	if EmpID<>"-2" then	%> 
  	<form name="UpLoadPhoto" method="post" enctype="multipart/form-data" action="employeeedit.asp?EmpId=<%=EmpId%>" align="center">
      <center>
  		<input type="file" name="PhotoFile" value="�������ѡ����Ƭ�ļ�">
  		<input type="submit" name="Submit" value="�����Ƭ">
      </center>
  	</form>
<%	end if	%>
</body>
<%
 	Set RSEmp=nothing
 	Set RsDept=nothing
 	Set G_DBConn=Nothing
%>
</html>