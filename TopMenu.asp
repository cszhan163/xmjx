<% @language="vbscript" %>
<!--#include file = "secret/checkpwd.asp"-->
<%
	mainItemCount = 0
	MenuGroup = Request("MenuGroup")
	if MenuGroup <> "" then
		Response.Cookies("MenuGroup") = MenuGroup
		Response.Cookies("MenuGroup").Expires = DateAdd("yyyy", 1, Date())
	else
		MenuGroup = Request.Cookies("MenuGroup")				'ȡ��Ҫ��ʾ�Ĳ˵���(Export, Domestic)
	end if


	Set RSMenu = Server.CreateObject("ADODB.Recordset")
	Set RSItem = Server.CreateObject("ADODB.Recordset")

	'ȡ�����˵���
	RSMenu.Open "SELECT Count(*) MenuCount FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"'", G_DBConn, 0, 1, 1
	if not RSMenu.EOF then
		MainMenuCount = RSMenu("MenuCount")
	end if
	RSMenu.Close

	'ȡ���Ӳ˵����������
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

	'���������Ӳ˵�������
	'dim menuArr(<%=MainMenuCount%>, <%=MaxItemCount%>)
	'dim menuTextArr(<%=MainMenuCount%>, <%=MaxItemCount%>)

<%	'������˵���ť
	RSMenu.Open "SELECT A.*, B.ShowPos FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"' ORDER BY B.ShowPos ASC", G_DBConn, 1, 1, 1
	do while not RSMenu.EOF
%>
	'�����������˵�����, �������˵����
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
		'�����Ӳ˵�����, �����Ӳ˵�ҳ��
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
<%		'��Ӳ˵��¼�
		if RSMenu("PageName") <> "" then			'���˵�û���Ӳ˵�ʱ
%>
	window.open "<%=RSMenu("PageName")%>", "mainFrame"
<%
		else										'���˵����Ӳ˵�ʱ
%>
	CurMenuPageName = menuArr(<%=RSMenu("ShowPos")%>,(<%=RSMenu("MenuCode")%>.CurMenuNo+1))
	Pos = InStr(1, CurMenuPageName, "/", 0)
	MenuGroup = Mid(CurMenuPageName, Pos + 1)

	'����˵����� UpdateMenu/ ��ͷ�����µ�ǰ���ڵĲ˵��������������ڴ�Ҫ���ҳ��
	if Left(CurMenuPageName, Pos - 1) = "UpdateMenu" then
		OldTitle = window.parent.document.title
		Pos = InStr(1, OldTitle, "/", 0)
		select case MenuGroup
<%
	'����Ҫ��ʾ�Ĳ˵������õ��������ñ������Ĳ˵�������
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
		window.parent.focus					'�ѽ���ת�Ƶ������ڣ��粻ת��ִ���������󴰿�ʧȥ����
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

<%	'������˵���ť
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

//���ѹ���
var Timer = window.setTimeout("QueryRemind()", "2000");		//���ó������������ѯ������ʱ2��
var MyReminder = new Reminder("Assist/CheckRemind.asp", "Assist/Remind.asp");
function onclickbutton(butItem){

	var str = butItem['id']
	var array = str.split('_')
	inputid  = array[0]+'_s'
	menum = array[1]
	alert('��ѡ���ұߵ�����ѡ��')
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
