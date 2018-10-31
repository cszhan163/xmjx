<html>
<head>
<title>公司信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript">
function checkdata()
{
var CorpChs=QCorp.CorpNameChs.value
var CorpCode=QCorp.CorpCode.value
	if(CorpCode=="")
	{
	alert("请输入公司代码")
	window.event.returnValue=false;
	}
	else
	{
		if(CorpChs=="")
		{
		alert("请录入公司中文名称")
		window.event.returnValue=false;
		}
	}
}
</script>
</head>
<%ModuleCode = "CA"%>
<!--#include virtual="/secret/checkpwd.asp"-->
<link href="/secret/style.css" rel=stylesheet type=text/css>
<body background="images/gback.jpg">
<%
	CorpId = Request("CorpId")
	
	Set RSCorp = Server.CreateObject("ADODB.RecordSet")
	Set RSDept= Server.CreateObject("ADODB.RecordSet")

	Submit=Request("Submit")
	
	if Submit="返回" then
		'Server.Transfer "MainCorp.asp"
		response.Redirect "MainCorp.asp"
	end if
	
   	if Submit="删除部门" then
  		DelID=Request("DelID")
  		if DelID<>"" then
            RSDept.Open "SELECT * FROM dept WHERE CorpId='"& CorpId &"' AND DeptCode='"& CorpId & DelID&"'",G_DBConn, 1, 3, 1
            if not RSDept.eof then
			    RSDept.DELETE 
                RSDept.Update
            end if
            RSDept.Close
  		end if
	end if
	
%>
<%
if submit="保存" then
	'保存公司基本信息
    RSCorp.Open "SELECT * FROM Corpinfo WHERE CorpId='"& CorpId &"'", G_DBConn, 1, 3, 1
	    if RSCorp.eof then
	       RSCorp.AddNew
	    end if
	    if Trim(Request("CorpCode"))="" then
			ErrMsg("请录入公司代码.")
		else
			RSCorp("CorpCode")=Request("CorpCode")
		end if
	    if Trim(Request("CorpNameChs"))="" then
			ErrMsg("请录入公司中文名称.")
		else
	  		RSCorp("CorpNameChs")=Request("CorpNameChs")
	  	end if
	  	RSCorp("CorpNameEng")=Request("CorpNameEng")
	  	RSCorp("CorpAddressChs")=Request("CorpAddressChs")
	  	RSCorp("CorpAddressEng")=Request("CorpAddressEng")
 		 	RSCorp("TelNo")=Request("TelNo")
  		RSCorp("FaxNo")=Request("FaxNo")
  		RSCorp("Telex")=Request("Telex")
  		RSCorp("Email")=Request("Email")
  		RSCorp("www")=Request("www")
		RSCorp("province")=Request("province")
		RSCorp("City")=Request("City")
		RSCorp("Postalcode")=Request("Postalcode")
		RSCorp("recordsCode")=Request("recordsCode")
		RSCorp("OrgnizeCode")=Request("OrgnizeCode")
		RSCorp("HxCode")=Request("hxCode")
		RSCorp("BillPrefix") = UCase(Request("BillPrefix"))
  		RSCorp.Update
  		CorpId = RSCorp("CorpId")
    RSCorp.Close

    '保存部门
    For i=1 To Request("DeptSerialNo").Count
       	DeptSerialNo=Request("DeptSerialNo")(i)
       	DeptID=Request("DeptID")(i)
       	if DeptID="" then
    	  	Exit For
       	end if
       	RSDept.Open "Select * FROM Dept WHERE CorpId='"& CorpId &"' AND DeptCode='"& CorpId & DeptId &"'", G_DBConn, 2, 3, 1
       		if DeptSerialNo="0" then
       	  		if not RSDept.eof then
					Response.Write "<body onclick='history.back()'>"
					Response.Write "<center><font color=red>此代码已经被使用,点击左键返回!</font></center>"
					Response.End
        		else
     	   			RSDept.AddNew
        		end if
      		end if 
      		RSDept("DeptCode") = CorpId & DeptID
      		RSDept("CorpId") = CorpId
	  		RSDept("DeptName") = request("Dept")(i) 
			RSDept("DeptTelNo") = Request("DeptTelNo")(i)
			RSDept("DeptFax") = Request("DeptFax")(i)
      		Rsdept.update  
      	RsDept.Close
   	next
 end if
%>
<%
	'显示数据
   	RSCorp.Open "SELECT * FROM CorpInfo WHERE CorpId='"& CorpId &"'", G_DBConn, 3,2,1
   		if not RSCorp.eof then
     		CorpCode=RSCorp("CorpCode")
     		CorpNameChs=RSCorp("CorpNameChs")
     		CorpNameEng=RSCorp("CorpNameEng")
     		CorpAddressChs=RSCorp("CorpAddressChs")
     		CorpAddressEng=RSCorp("CorpAddressEng")
     		TelNo=RSCorp("TelNo")
     		FaxNo=RSCorp("FaxNo")
     		Telex=RSCorp("Telex")
     		Email=RSCorp("Email")
     		www=RSCorp("www")
			province=RSCorp("province")
			City=RSCorp("City")
			Postalcode=RSCorp("Postalcode")
			recordsCode=RSCorp("recordsCode")
			OrgnizeCode=RSCorp("OrgnizeCode")
			HxCode=RSCorp("Hxcode")
			BillPrefix = RSCorp("BillPrefix")
		end if
	  RSCorp.close
%>
<form method="post"  name="QCorp" action="Corpinfo.asp">
  <center>
    <img src="images/CorpInfo.gif" width="236" height="24">
  </center>
  <table border="1" bgcolor="#FFE8D0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" width="555" align="center">
    <tr class=tdcss>
      <td width="109"  height="26" bgcolor="DDDDDD"><div align="right">公司代码<%=EmphasisTag%> ：</div></td>
      <td  height="26" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CorpCode" size="60" value="<%=CorpCode%>" class=longinput>
      </td>
    </tr >
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">名 &nbsp; &nbsp;称<%=EmphasisTag%> ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CorpNameChs" size="60" value="<%=CorpNameChs%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">英 文 名 ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CorpNameEng" size="60" value="<%=CorpNameEng%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">地 &nbsp; &nbsp;址<%=EmphasisTag%> ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CorpAddressChs" size="60" value="<%=CorpAddressChs%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">英文地址 ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="CorpAddressEng" size="60" value="<%=CorpAddressEng%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">电 &nbsp; &nbsp;话 ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="TelNo" size="60" value="<%=TelNo%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">传 &nbsp; &nbsp;真 ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="FaxNo" size="60" value="<%=FaxNo%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td height="30"  bgcolor="DDDDDD"><div align="right">电 &nbsp; &nbsp;传 ：</div></td>
      <td height="30" colspan="3"  bgcolor="#FFFFFF"><input type="text" name="Telex" size="60" value="<%=Telex%>" class=longinput>
      </td>
    </tr>
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">E-mail &nbsp;：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="Email" size="60" value="<%=Email%>" class=longinput>
      </td>
    </tr >
    <tr class=tdcss>
      <td  height="30" bgcolor="DDDDDD"><div align="right">U R L&nbsp; ：</div></td>
      <td  height="30" colspan="3" bgcolor="#FFFFFF"><input type="text" name="www" size="60" value="<%=www%>" class=longinput>
      </td>
    </tr>
    <tr bgcolor="DDDDDD" align="center">
      <td height="30">省&nbsp; &nbsp;份：</td>
      <td bgcolor="FFFFFF"><input type="text" name="province" value="<%=province%>" class=shortinput></td>
      <td height="30">城市</td>
      <td bgcolor="FFFFFF"><input type="text" name="City" value="<%=City%>" class=shortinput></td>
    </tr>
    <tr bgcolor="DDDDDD" align="center" class=tdcss>
      <td>邮政编码</td>
      <td bgcolor="FFFFFF"><input type="text" name="Postalcode" value="<%=Postalcode%>" class=shortinput></td>
      <td>备案号</td>
      <td bgcolor="FFFFFF"><input type="text" name="recordsCode" value="<%=recordsCode%>" class=shortinput></td>
    </tr>
    <tr bgcolor="DDDDDD" align="center" class=tdcss>
      <td height="30">组织机构代码</td>
      <td bgcolor="FFFFFF"><input type="text" name="OrgnizeCode" value="<%=OrgnizeCode%>" class="shortinput"></td>
      <td>核销代号</td>
      <td bgcolor="FFFFFF"><input type="text" name="HxCode" value="<%=HxCode%>" class=shortinput></td>
    </tr>
    <tr align="center" class=tdcss bgcolor="white">
      <td bgcolor="DDDDDD">单据编号前缀</td>
      <td><input type="text" name="BillPrefix" value="<%=BillPrefix%>" class="shortinput">
      </td>
      <td bgcolor="DDDDDD">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr class=tdcss bgcolor="DDDDDD">
      <td  height="30"><div align="center">部门代码</div></td>
      <td width="213"  height="30"><div align="center">部 门 名 称</div></td>
      <td width="219" align="center" valign="middle">电话号码</td>
      <td width="219" align="center" valign="middle">传&nbsp;真</td>
    </tr>
    <%
	RSDept.Open "Select A.*, RIGHT(A.DeptCode, LEN(A.DeptCode) - LEN(A.CorpId)) DeptCode  FROM Dept A WHERE CorpId='"& CorpId &"'", G_DBConn, 0, 1, 1
      	i=1 
      	Do While NOT RSDept.eof
%>
    <tr class=tdcss>
      <td  bgcolor="#FFFFFF" align="center"><%=RSDept("DeptCode")%>
        <input type="hidden" name="DeptID" value="<%=RSDept("DeptCode")%>" >
      </td>
      <td  bgcolor="#FFFFFF"><div align="left">
          <input type="hidden" name="DeptSerialNo" value="<%=RSDept("DeptCode")%>">
          <input type="text" name="Dept" size="30" value="<%=RSDept("DeptName")%>" class=midinput>
        </div></td>
      <td align="left" valign="middle"  bgcolor="#FFFFFF"><input name="DeptTelNo" type="text" size="20" value="<%=RSDept("DeptTelNo")%>" class=input style="width:100">
      </td>
      <td  bgcolor="#FFFFFF"><input type="text" name="DeptFax" size="20" value="<%=RSDept("DeptFax")%>" class=input style="width:100">
      </td>
    </tr>
    <%
     		RSDept.MoveNext
     		i=i+1
     	Loop
  	RSDept.close
%>
    <tr class=tdcss>
      <td  bgcolor="#FFFFFF"><div align="right">
          <input name="DeptId" type="text" class=input maxlength="4">
        </div></td>
      <td  bgcolor="#FFFFFF"><div align="left">
          <input type="hidden" name="DeptSerialNo" value="0">
          <input type="text" name="Dept" size="30" class=midinput>
        </div></td>
      <td align="left" valign="middle"  bgcolor="#FFFFFF"><input name="DeptTelNo" type="text" size="20" class=input style="width:100">
      </td>
      <td align="left" valign="middle"  bgcolor="#FFFFFF"><input name="DeptFax" type="text" size="20" class=input style="width:100">
      </td>
    </tr>
    <tr>
      <td  bgcolor="DDDDDD" colspan="4" align="center"><input type="text" name="DelID" value="请输入部门代码" class=midinput>
        <input type="submit" name="Submit" value="删除部门">
        <input type="submit" name="Submit" value="保存"  onClick="checkdata()">
        <input type="submit" name="Submit" value="返回">
        <input type="hidden" name="CorpId" value="<%=CorpId%>">
      </td>
    </tr>
  </table>
</form>
</body>
<%
   	Set RSDept=Nothing
   	Set G_DBConn=Nothing
%>
</html>
