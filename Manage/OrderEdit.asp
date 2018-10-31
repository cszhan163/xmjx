<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<!--#include virtual="/secret/checkpwd.asp"-->
<!--#include virtual="/secret/Func.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="/secret/style.css" type=text/css rel=stylesheet>
<base target="_self">
<title>排序要素及项目</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 18pt;
	font-weight: bold;
}
-->
</style>
</head>
<%
	set rsMain=Server.CreateObject("ADODB.Recordset")
	set rsTemp=Server.CreateObject("ADODB.Recordset")
	
	ExamineId=request("ExamineId")
	UpOrDown=request("UpOrDown")
	ExaItemId=request("ExaItemId")
	ClientY=request("ClientY")
	
	if UpOrDown="Up" then
		rsMain.open "select * from ExamineItem where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
		if not rsMain.eof then
			OrderNum=rsMain("OrderNum")
		end if
		rsMain.close
		UOrderNum=cdbl(OrderNum)-1
		rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&" and OrderNum="&UOrderNum&""
		if not rsMain.eof then
			DExaItemId=rsMain("ExaItemId")
		end if
		rsMain.close
		if DExaItemId<>"" and not isnull(DExaItemId) then
			G_DBConn.execute("update ExamineItem set OrderNum=-1 where ExaItemId="&ExaItemId&"")
			G_DBConn.execute("update ExamineItem set OrderNum="&OrderNum&" where ExaItemId="&DExaItemId&"")
			G_DBConn.execute("update ExamineItem set OrderNum="&UOrderNum&" where ExaItemId="&ExaItemId&"")
		end if
	end if
	if UpOrDown="Down" then
		rsMain.open "select * from ExamineItem where ExaItemId="&ExaItemId&"",G_DBConn,1,1,1
		if not rsMain.eof then
			OrderNum=rsMain("OrderNum")
		end if
		rsMain.close
		DOrderNum=cdbl(OrderNum)+1
		rsMain.open "select * from ExamineItem where ExamineId="&ExamineId&" and OrderNum="&DOrderNum&""
		if not rsMain.eof then
			UExaItemId=rsMain("ExaItemId")
		end if
		rsMain.close
		if UExaItemId<>"" and not isnull(UExaItemId) then
			G_DBConn.execute("update ExamineItem set OrderNum=-1 where ExaItemId="&ExaItemId&"")
			G_DBConn.execute("update ExamineItem set OrderNum="&OrderNum&" where ExaItemId="&UExaItemId&"")
			G_DBConn.execute("update ExamineItem set OrderNum="&DOrderNum&" where ExaItemId="&ExaItemId&"")
		end if
	end if
	
	Window_OffsetY=request("Window_OffsetY")
	if Window_OffsetY="" or isnull(Window_OffsetY) then Window_OffsetY=0
%>
<body onLoad="window.scrollTo(0,<%=Window_OffsetY%>)">
<center>
<h2>要素排序编辑</h2>
</center>
<form name="form1" action="OrderEdit.asp?ClientY=<%=ClientY%>&ExamineId=<%=ExamineId%>" method="post">
<table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td align="right">
<DIV   id=BtnDiv   style="border:0   solid   #808080; width:500; height:  30; position:   absolute; z-index:4; left: 164px; top: 45px;"> 
<input name="Submits" type="submit" id="Submits" value="确定" onClick="PageClose(<%=ExamineId%>)">
<input name="Up" type="button" id="Up" value="上↑" onClick="getUp()">
<input name="DOWN" type="button" id="DOWN" value="下↓" onClick="Down()">
</DIV>
</td>
</tr>
</table>
<table width="70%" border="1" align="center" bgcolor="#FFE8D0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999">
  <tr>
    <td width="10%" align="center" bgcolor="DDDDDD" class="tdcss">序号</td>
    <td width="32%" align="center" bgcolor="DDDDDD" class="tdcss">项目名称</td>
    <td width="47%" align="center" bgcolor="DDDDDD" class="tdcss">要素名称</td>
  </tr>
<%
	RowIndex=0
	rsMain.open "select EI.ExaItemId,EP.ExaProdName,EF.ExaFactorName from ExamineItem EI "&_
		"left join ExaProdject EP on EP.ExaProdId=EI.ExaProdId "&_
		"left join ExaFactor EF on EF.ExaFactorId=EI.ExaFactorId "&_
		"where ExamineId="&ExamineId&" order by OrderNum ",G_DBConn,1,1,1
	do while not rsMain.eof
		RowIndex=cdbl(RowIndex)+1
		CurExaItemId=rsMain("ExaItemId")
		ExaProdName=rsMain("ExaProdName")
		ExaFactorName=rsMain("ExaFactorName")
		if cdbl(CurExaItemId)=cdbl(ExaItemId) then
			strColor="#FFCCCC"
		else
			strColor="#FFFFFF"
		end if
%>
  <tr onClick="SelRow(<%=CurExaItemId%>)" bgcolor="<%=strColor%>">
    <td align="center" >&nbsp;<%=RowIndex%></td>
    <td align="center" >&nbsp;<%=ExaProdName%></td>
    <td align="center" >&nbsp;<%=ExaFactorName%></td>
  </tr>
<%
		rsMain.movenext
	loop
	rsMain.close
%>
</table>
<input type="hidden" name="UpOrDown" value="">
<input type="hidden" name="ExaItemId" value="<%=ExaItemId%>">
<input type="hidden" name="txt0" value="<%=ClientY%>">
<input type="hidden" name="ExaId" value="<%=ExamineId%>">
<input type="hidden" id="Window_OffsetY" name="Window_OffsetY" value="0">
<input type="hidden" id="Window_OffsetX" name="Window_OffsetX" value="0">
</form>
<script language="vbscript">
	sub getUp()
		ExaItemId=form1.ExaItemId.value
		if ExaItemId<>"" and not isnull(ExaItemId) then
			form1.UpOrDown.value="Up"
			Form1.Window_OffsetY.value=document.body.scrollTop
			form1.submit()
		end if
	end sub
	sub Down()
		ExaItemId=form1.ExaItemId.value
		if ExaItemId<>"" and not isnull(ExaItemId) then
			form1.UpOrDown.value="Down"
			Form1.Window_OffsetY.value=document.body.scrollTop
			form1.submit()
		end if
	end sub
	sub SelRow(ExaItemId)
		form1.ExaItemId.value=ExaItemId
		Form1.Window_OffsetY.value=document.body.scrollTop
		form1.submit()
	end sub
</script>
<script language="javascript">
	function PageClose(ExamineId)
	{
		window.returnValue = document.getElementById("txt0").value; 
		window.dialogArguments.location.reload("ExamineEdit.asp"+"?Window_OffsetY="+window.returnValue+"&ExamineId="+ExamineId); 
		window.close();
	}
</script>
<script   language="javascript">   
  <!--   
  //window.onload   =   resizeDiv;   
  window.onresize   =   resizeDiv;   
  window.onscroll   =   resizeDiv;   
  window.onerror   =   function(){}   
  function   resizeDiv()   
  {   
  var   obj=document.getElementById("BtnDiv")   
  try{    
 	obj.style.top   =   parseInt(document.body.scrollTop,10)+45;  
  	obj.style.left =   parseInt(document.body.scrollLeft,10)+164;
  }   
  catch(e){}   
  }   
  -->   
  </script>
</body>
</html>
