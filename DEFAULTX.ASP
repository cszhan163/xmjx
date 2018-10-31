<% @language="vbscript" %>
<!--#include file = "secret/checkpwd.asp"--> 
<%
	MenuGroup = Request.Cookies("MenuGroup")				'取得要显示的菜单组(Export, Domestic)

    Set RStemp=Server.CreateObject("ADODB.RecordSet")

    RSTemp.Open "Select empNameChs from Employee Where EmpCode='"&userId&"'",g_DBConn,2,3,1
        if not RStemp.eof then
           EmpNameChs=RStemp("EmpNameChs")
        end if
    RSTemp.Close

	'默认外贸子系统菜单
	if MenuGroup <> "Export" then
		MenuGroup = "Export"
		Response.Cookies("MenuGroup") = "Export"
		Response.Cookies("MenuGroup").Expires = DateAdd("yyyy", 1, Date())
	end if

	RSTemp.Open "SELECT MenuGroupName FROM Sys_MenuGroup WHERE MenuGroupCode = '"& MenuGroup &"'", g_DBConn, 0, 1, 1
	if not RSTemp.EOF then
		SubSystem = RSTemp("MenuGroupName")
	end if
	RSTemp.Close
%>
<HTML>
<HEAD>
<TITLE>上海迈可进出口有限公司 / <%=SubSystem%> &nbsp;-- <%="当前操作人员: "&EmpNameChs%></TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<%
	Set RSEmp = Server.CreateObject("ADODB.Recordset")

	'初始页面
	RSEmp.Open "SELECT M.PageName FROM Employee A LEFT JOIN Sys_Module M ON A.MainPageCode = M.ModuleCode "&_
			   "WHERE A.EmpCode = '"& UserId &"'", G_DBConn, 0, 1, 1
	if not RSEmp.EOF then
		hintFile = RSEmp("PageName")
	end if
	RSEmp.Close 

%>
<FRAMESET border=0 frameSpacing=0 rows=34,470* frameBorder=NO cols=*>
  <FRAME name=topFrame src="TopMenu.asp" noResize scrolling=no>
  <FRAME name=mainFrame src="<%=hintFile%>">
</FRAMESET>
<noframes></noframes>
</HTML>
