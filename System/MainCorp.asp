<%@ LANGUAGE = VBScript %>
<html>
<head>
<title>��˾��Ϣ��±�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<%ModuleCode = "CA"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<%
Set RSProd = Server.CreateObject("ADODB.Recordset")
Set RSDB = Server.CreateObject("ADODB.Recordset")
SearchStr=curselvalue("SearchStr")
Submit=Request("Submit")

if Submit = "��ӷֹ�˾" then
	Response.Redirect "CorpInfo.asp?CorpId=-2"
	Response.End 
end if

if SearchStr="" then
  RSProd.Open "SELECT * FROM CorpInfo ORDER BY CorpID", G_DBConn, 0, 1, 1
else
  RSProd.Open "SELECT * FROM CorpInfo WHERE CorpID LIKE '%"& SearchStr &"%' OR CorpNameChs LIKE '%"&SearchStr&"%' OR CorpNameEng LIKE '%"&SearchStr&"%' ORDER BY CorpID", G_DBConn, 0, 1, 1
end if
    '��־��Ϣ
    'LogInfo "ϵͳά��","MainCorp.asp","CorpInfo","",SearchStr,9
%>
<body background="images/gback.jpg">
<form name="forma" action="MainCorp.asp" method="post">
  <div align="center"><b><font size="5">��˾��Ϣά��</font></b></div>
<div align="center"> 
  <table border="0" width="97%" cellspacing="0" cellpadding="0" height="32" style="font:13px">
    <tr>
      <td width="46%" height="36"><input type="submit" name="Submit" value="��ӷֹ�˾"></td> 
      <td width="54%" height="36"><div align="right"><b>�ֹ�˾����</b> 
            <input type="text" name="SearchStr" value="<%=SearchStr%>">  
            <input type="submit" name="Submit" value="��ѯ">  
          </div>  
      </td>  
    </tr>  
  </table>  
    <table border="1" align="center" width="100%" name="EmpGrid" bgcolor="#FFFFFF" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" style="font:13px">
      <tr height="25" align="center" bgcolor="DDDDDD"> 
        <td width="26%" nowrap>�ֹ�˾����</td>
        <td width="19%" nowrap>����</td>
        <td width="21%" nowrap>�绰</td>
        <td width="34%" nowrap>��ַ</td>
    </tr>
    <%Do While NOT RSProd.eof%> 
      <tr bgcolor="#FFFFFF"> 
        <td width="26%" height="25" nowrap> 
          <div align="center"><font color="#3333CC" ><a href="CorpInfo.Asp?CorpId=<%=RSProd("CorpId")%>"><%=RSProd("CorpNameChs")%></a></font></div>
      </td>
        <td width="19%" height="25" nowrap><font color="#3333CC" ><%=RSProd("FaxNo")%>&nbsp;</font></td>
        <td width="21%" height="25"> 
          <div align="left"><font color="#3333CC"><%=RSProd("TelNo")%>&nbsp;</font></div>
      </td>
        <td width="34%" height="25"> 
          <div align="left"><font color="#3333CC"><%=RSProd("CorpAddressChs")%>&nbsp;</font></div>
      </td>
    </tr>
    <%
  RSProd.MoveNext
Loop  
%> 
  </table> 
</div>
</form>
<%G_DBConn.Close%>
</body>  
</html>
<html></html>