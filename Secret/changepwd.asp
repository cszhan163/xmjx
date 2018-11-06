<%@ LANGUAGE = VBScript %>
<!--#include virtual="/secret/crypt.asp"-->
<title>上海迈可进出口有限公司</title><body background="/secret/images/gback.jpg">
<%
	'程序
	Submit=Request("Submit")
	if Submit="  返  回  " then
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
      <div align="center"><b><FONT size="4" face="楷体_GB2312">输入的用户ID或密码错误!</FONT></b></div>
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
      <div align="center"><b><FONT size="4" face="楷体_GB2312">新密码与确认密码不符!</FONT></b></div>
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
        <p align="center"><font size="6" color="#000066" face="楷体_GB2312"><b>请输入用户名和密码</b></font></p>
        <form name="CheckPwd" method="post" action="changepwd.asp" >
          <table width="100%" border="0">
            <tr> 
              <td> 
                <div align="right"><font size="4" face="楷体_GB2312" color="#000066"><b>用户名:</b></font></div>
              </td>
              <td>
                <input type="text" name="UID">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="楷体_GB2312" color="#000066"><b>密&nbsp;&nbsp;码:</b></font></div>
              </td>
              <td> 
                <input type="password" name="pwd">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="楷体_GB2312" color="#000066"><b>新密码:</b></font></div>
              </td>
              <td>
                <input type="password" name="newpwd">
              </td>
            </tr>
            <tr> 
              <td> 
                <div align="right"><font size="4" face="楷体_GB2312" color="#000066"><b>确认密码:</b></font></div>
              </td>
              <td> 
                <input type="password" name="conpwd">
              </td>
            </tr>
          </table>
          <p align="center"> 
            <input type="submit" name="Submit" value="  修  改  ">
            <input type="reset" name="Reset" value="  清  除  ">
            <input type="submit" name="Submit" value="  返  回  ">
          </p>
          </form>
      </td>
    </tr>
  </table>
</div>
</body>