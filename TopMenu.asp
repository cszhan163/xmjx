<% @language="vbscript" %>
<!--#include file = "secret/checkpwd.asp"-->
<%
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
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<BODY leftMargin="0" topMargin="0" background="img/gback.jpg">
<table border=0 cellpadding="2" cellspacing="0" align="center">
  <tr align="center" valign="middle" > 
    <td nowrap>
<%	'������˵���ť
	RSMenu.Open "SELECT A.*, B.ShowPos FROM Sys_Menu A LEFT JOIN Sys_Group_Menu B ON A.MenuCode = B.MenuCode "&_
				"WHERE B.GroupCode = '"& MenuGroup &"' ORDER BY B.ShowPos ASC", G_DBConn, 1, 1, 1
	do while not RSMenu.EOF
%>
<object
	  classid="clsid:3D1B5945-11CD-11D8-B7DC-00E04C40A1DE"
	  codebase="Controls/MenuOcxProj1.inf"
	  width=90
	  height=24
	  align=center
	  hspace=0
	  vspace=0
      id="<%=RSMenu("MenuCode")%>">
</object>
<%
		RSMenu.MoveNext
	loop

	if not RSMenu.BOF then
		RSMenu.MoveFirst
	end if
%>
    </td>
  </tr>
</table>

<script language="vbscript">
	'���������Ӳ˵�������
	dim menuArr(<%=MainMenuCount%>, <%=MaxItemCount%>)
<%
	do while not RSMenu.EOF
	'�����������˵�����, �������˵����
%>
	<%=RSMenu("MenuCode")%>.setBtnTxt("<%=RSMenu("MenuName")%>")
	<%=RSMenu("MenuCode")%>.btnWidth = 90
<%
		RSItem.Open "SELECT A.ModuleName, B.ShowPos, A.PageName FROM Sys_Module A LEFT JOIN Sys_Group_Module B ON A.ModuleCode = B.ModuleCode "&_
					"WHERE A.MenuPos = '"& RSMenu("MenuPos") &"' AND A.MenuItemPos <> 0 AND B.GroupCode = '"& MenuGroup &"' "&_
					"ORDER BY B.ShowPos ASC", G_DBConn, 0, 1, 1
		do while not RSItem.EOF
		'�����Ӳ˵�����, �����Ӳ˵�ҳ��
%>
	<%=RSMenu("MenuCode")%>.AddMenuItem("<%=RSItem("ModuleName")%>")
	menuArr(<%=RSMenu("ShowPos")%>, <%=RSItem("ShowPos")%>) = "<%=RSItem("PageName")%>"
<%
			RSItem.MoveNext
		loop
		RSItem.Close
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
</script>
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

function QueryRemind()
{
	MyReminder.CheckRemind();
}
-->
</SCRIPT>
</HTML>