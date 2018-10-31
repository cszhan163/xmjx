<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上海迈可进出口有限公司</title>

<style type="text/css">
<!--
#Layer1 {
	position:absolute;
	width:100px;
	height:47px;
	z-index:1;
	left: 280px;
	top: 320px;
}
-->
</style>
</head>
<body topmargin="0" leftmargin="0" scroll="no" onkeydown="if(event.keyCode==13) keyDown()" onLoad="init()"> 

<img width="100%" height="100%" src="../IMG/KFBackground.jpg" style="position:absolute; left:0px; top:0px">
<form name="qform" action="Defaultx.asp" method="post">
	<div style="width:100%; height:100%; position:absolute; background-image:url(../IMG/MKTITLE.gif); background-repeat:no-repeat; background-position:10% 25%"></div>
	<div style="width:100%; height:100%; position:absolute; background-image:url(../IMG/YHHM.gif); background-repeat:no-repeat; background-position:70% 70%"></div>
	<div style="width:100%; height:100%; position:absolute; background-image:url(../IMG/YYMM.gif); background-repeat:no-repeat; background-position:70% 75%"></div>
	<div style="position:absolute; left:65%; top:78%"><Input type="Button" name="Button2" value="修改密码" OnClick="document.location='secret/changepwd.asp'"></div>
  <div style="position:absolute; left:75%; top:78%"><input type="Button" name="Button1" value=" 登 录 " onClick="qform.submit()">
	</div>
	<div style="position:absolute; left:73%; top:68%"><input type="text" name="Name" value="" style="width:100px; border:solid 1px #000000"></div>
	<div style="position:absolute; left:73%; top:72.5%"><input type="password" name="pwd" style="width:100px; border:solid 1px #000000"></div>
</form>
<img style="position:absolute; left:3%; bottom:8%" src="../IMG/hongRui.gif" WIDTH="250" HEIGHT="35">
<!--<a href="/BBS/BBS.asp"><img src="IMG/GGXX.gif" width="113" height="43" style="border:0px solid gray;position:absolute; left:30%; bottom:40%"></a>-->
</body>
<script language="vbscript">
    function init()
        qform.name.focus()
        if qform.name.value<>"" then
            qform.pwd.focus()
        end if
    end function 
</script>
<script language="javascript">
    function keyDown()
    {
        Name=document.qform.Name.value;
        if (Name!="")
       {
           pwd=document.qform.pwd.value;
           if (pwd=="")
           {
                 document.qform.pwd.focus();
            }
           else
           {
                 qform.submit();
           }
       }
       else
       {
           document.qform.Name.focus();
       }
     }
</script>
</html>
