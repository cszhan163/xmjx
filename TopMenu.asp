<% @language="vbscript" %>
<!--#include file = "secret/checkpwd.asp"-->
<%
	mainItemCount = 0
	MenuGroup = Request("MenuGroup")
	if MenuGroup <> "" then
		Response.Cookies("MenuGroup") = MenuGroup
		Response.Cookies("MenuGroup").Expires = DateAdd("yyyy", 1, Date())
	else
		MenuGroup = Request.Cookies("MenuGroup")				'取得要显示的菜单组(Export, Domestic)
	end if


	Set RSMenu = Server.CreateObject("ADODB.Recordset")
	Set RSItem = Server.CreateObject("ADODB.Recordset")

	'取得主菜单数
	RSMenu.Open "SELECT Count(*) MenuCount FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"'", G_DBConn, 0, 1, 1
	if not RSMenu.EOF then
		MainMenuCount = RSMenu("MenuCount")
	end if
	RSMenu.Close

	'取得子菜单的最大项数
	RSMenu.Open "SELECT MAX(A.MenuItemCount) MaxItemCount FROM "&_
				"(SELECT Count(*) MenuItemCount FROM Sys_Module B LEFT JOIN Sys_Group_Module C ON B.ModuleCode = C.ModuleCode "&_
				"WHERE B.MenuPos <> 0 AND C.GroupCode = '"& MenuGroup &"' GROUP BY MenuPos) A"
	if not RSMenu.EOF then
		MaxItemCount = RSMenu("MaxItemCount")
	end if
	RSMenu.Close

	Dim menucoutArr(10)
    dim menuTextArr(10,10)
		dim menuRefArr(10,10)
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<BODY leftMargin="0" topMargin="0" background="img/gback.jpg">

<script language="vbscript">

	'定义所有子菜单项数组
	'dim menuArr(<%=MainMenuCount%>, <%=MaxItemCount%>)
	'dim menuTextArr(<%=MainMenuCount%>, <%=MaxItemCount%>)

<%	'添加主菜单按钮
	RSMenu.Open "SELECT A.*, B.ShowPos FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"' ORDER BY B.ShowPos ASC", G_DBConn, 1, 1, 1
	do while not RSMenu.EOF
%>
	'设置所有主菜单标题, 设置主菜单宽度
	 <%
		mainItemCount = 0
	 %>

%>
	<%=RSMenu("MenuCode")%>.text = "<%=RSMenu("MenuName")%>"
	<%=RSMenu("MenuCode")%>.height = 90
<%
		RSItem.Open "SELECT A.ModuleName, B.ShowPos, A.PageName FROM Sys_Module A LEFT JOIN Sys_Group_Module B ON A.ModuleCode = B.ModuleCode "&_
					"WHERE A.MenuPos = '"& RSMenu("MenuPos") &"' AND A.MenuItemPos <> 0 AND B.GroupCode = '"& MenuGroup &"' "&_
					"ORDER BY B.ShowPos ASC", G_DBConn, 0, 1, 1
		do while not RSItem.EOF
		'设置子菜单名称, 设置子菜单页面
%>
	 <%
	 	mainItemCount = mainItemCount+1
		menuTextArr(RSMenu("ShowPos"),RSItem("ShowPos"))= RSItem("ModuleName")
		menuRefArr(RSMenu("ShowPos"), RSItem("ShowPos")) = RSItem("PageName")
	 %>
	 menuArr(<%=RSMenu("ShowPos")%>, <%=RSItem("ShowPos")%>) = "<%=RSItem("PageName")%>"

<%
			RSItem.MoveNext
		loop
		RSItem.Close
		menucoutArr(RSMenu("ShowPos")) = mainItemCount
%>



sub <%=RSMenu("MenuCode")%>_OnClick()
<%		'添加菜单事件
		if RSMenu("PageName") <> "" then			'主菜单没有子菜单时
%>
	window.open "<%=RSMenu("PageName")%>", "mainFrame"
<%
		else										'主菜单有子菜单时
%>
	CurMenuPageName = menuArr(<%=RSMenu("ShowPos")%>,(<%=RSMenu("MenuCode")%>.CurMenuNo+1))
	Pos = InStr(1, CurMenuPageName, "/", 0)
	MenuGroup = Mid(CurMenuPageName, Pos + 1)

	'如果菜单项以 UpdateMenu/ 开头，更新当前窗口的菜单，否则在主窗口打开要求的页面
	if Left(CurMenuPageName, Pos - 1) = "UpdateMenu" then
		OldTitle = window.parent.document.title
		Pos = InStr(1, OldTitle, "/", 0)
		select case MenuGroup
<%
	'根据要显示的菜单组代码得到用于设置标题栏的菜单组名称
	RSItem.Open "SELECT MenuGroupCode, MenuGroupName FROM Sys_MenuGroup", g_DBConn, 0, 1, 1
	do while not RSItem.EOF
%>
			case "<%=RSItem("MenuGroupCode")%>"
				MenuGroupName = "<%=RSItem("MenuGroupName")%>"
<%
		RSItem.MoveNext
	loop
	RSItem.Close
%>
		end select

		window.parent.document.title = Left(OldTitle, Pos) &" "& MenuGroupName & Mid(OldTitle, Pos + 6)
		window.parent.focus					'把焦点转移到父窗口，如不转移执行下面语句后窗口失去焦点
		window.location.replace "TopMenu.asp?MenuGroup="& MenuGroup
	else
		window.open menuArr(<%=RSMenu("ShowPos")%>,(<%=RSMenu("MenuCode")%>.CurMenuNo+1)), "mainFrame"
	end if
<%
		end if
%>
end sub
<%
		RSMenu.MoveNext
	loop
	RSMenu.Close
%>

function openNewPage(row,col)

	window.open menuArr(row,col)

end function

function openNewPage1(sel)

	sel.options[sel.selectedIndex].value;
	window.open menuArr(row,col)

end function

</script>
<table border=0 cellpadding="2" cellspacing="0" align="center">
  <tr align="center">

<%	'添加主菜单按钮
	RSMenu.Open "SELECT A.*, B.ShowPos FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"' ORDER BY B.ShowPos ASC", G_DBConn, 1, 1, 1
	do while not RSMenu.EOF
%>
<td nowrap>
<input id="<%=RSMenu("MenuCode")%>_<%=RSMenu("ShowPos")%>" type="button" onclick="onclickbutton(this)" value="<%=RSMenu("MenuName")%>">
</input>
<select id="<%=RSMenu("MenuCode")%>_s" onchange="selected(this)">
<%
menum = RSMenu("ShowPos")
for i=0 to menucoutArr(menum)
  strRef = menuRefArr(menum,(i+1))
	response.write("<option value="&""""&strRef&""">"&menuTextArr(menum,(i+1)))
	response.write("</option>")
next
%>
</select>
</td>
<%
		RSMenu.MoveNext
	loop

	if not RSMenu.BOF then
		RSMenu.MoveFirst
	end if
%>
</tr>
</table>
</BODY>
<script language="jscript" src="Script/Remind.js"></script>
<SCRIPT language="JavaScript1.2">
<!--
document.body.scroll = "no";
if (document.all)
	document.body.style.cssText="border:4 ridge #0099ff";

//提醒功能
var Timer = window.setTimeout("QueryRemind()", "2000");		//设置初次向服务器查询提醒延时2秒
var MyReminder = new Reminder("Assist/CheckRemind.asp", "Assist/Remind.asp");
function onclickbutton(butItem){

	var str = butItem['id']
	var array = str.split('_')
	inputid  = array[0]+'_s'
	menum = array[1]
	alert('请选择右边的下拉选项')
	//alert(menum)
	//AddOption(array[0]+'_s',array[1])
}
function selected(sel){

var str=sel.options[sel.selectedIndex].value;
//alert(str);
html = str
var array = str.split('_')
var row = array[0]
var col = array[1]
//openNewPage(row,col);
//alert(menuArr(row,col))
window.open(html,'mainFrame')
}

function QueryRemind()
{
	MyReminder.CheckRemind();
}
-->
</SCRIPT>
</HTML>
