<%@ LANGUAGE = VBScript %>
<!--#include virtual="/secret/crypt.asp"-->
<title>�Ϻ����ɽ��������޹�˾</title><body background="/secret/images/gback.jpg">
<%
	'����
	Submit=Request("Submit")
	if Submit="  ��  ��  " then
%>
	<meta http-equiv="refresh" content="0;URL=/">
<%
		Response.End
	end if
%>
<%
	Set RSDB = Server.CreateObject("ADODB.Recordset")
	Set G_DBConn = Server.CreateObject("ADODB.Connection")

	UID=Request("UID")
	pwd=md5(Request("pwd"))
	newpwd=Request("newpwd")
	conpwd=Request("conpwd")
	B=0
	if UID<>"" and pwd<>"" and newpwd<>"" and conpwd<>"" then
	  	if newpwd=conpwd then
    		SVRName=Request.ServerVariables("SERVER_NAME")
			DataBase="chemrole"
			G_DBConn.Open Application("ConnectionString")
		    RSDB.Open "SELECT PassWord FROM employee WHERE EmpCode='"&UID&"' AND PassWord='"&pwd&"'", G_DBConn, 2, 3, 1
    			if RSDB.eof then
%>
<table border="0" width="100%">
  <tr> 
    <td> 
      <div align="center"><b><FONT size="4" face="����_GB2312">������û�ID���������!</FONT></b></div>
    </td></tr></table>
<%
    			else
      				RSDB("PassWord")=md5(newpwd)
      				RSDB.Update
      				B=1
    			end if
    		RSDB.Close
    		G_DBConn.Close
  		else
%>
<table border="0" width="100%">
  <tr> 
    <td> 
      <div align="center"><b><FONT size="4" face="����_GB2312">��������ȷ�����벻��!</FONT></b></div>
    </td></tr></table>
<%
  		end if
	end if
	if B=1 then
%>
<meta http-equiv="refresh" content="0;URL=/">
<%
		Response.end
	end if
%>
<div align="center">
  <table width="75%" border="1" height="334">
    <tr> 
      <td> 
        <p align="center"><font size="6" color="#000066" face="����_GB2312"><b>�������û���������</b></font></p>
        <form name="CheckPwd" method="post" action="changepwd.asp" >
          <table width="100%" border="0">
            <tr> 
              <td> 
                <div align="right"><font size="4" face="����_GB2312" color="#000066"><b>�û���:</b></font></div>
              </td>
              <td>
                <input type="text" name="UID">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="����_GB2312" color="#000066"><b>��&nbsp;&nbsp;��:</b></font></div>
              </td>
              <td> 
                <input type="password" name="pwd">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="����_GB2312" color="#000066"><b>������:</b></font></div>
              </td>
              <td>
                <input type="password" name="newpwd">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="����_GB2312" color="#000066"><b>ȷ������:</b></font></div>
              </td>
              <td> 
                <input type="password" name="conpwd">
              </td>
            </tr>
          </table>
          <p align="center"> 
            <input type="submit" name="Submit" value="  ��  ��  ">
            <input type="reset" name="Reset" value="  ��  ��  ">
            <input type="submit" name="Submit" value="  ��  ��  ">
          </p>
          </form>
      </td>
    </tr>
  </table>
</div>
</body>