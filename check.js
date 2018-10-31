//javascript验证客户端数据。是否为空，是否数字，是否日期时间
//strname验证对象名称如"姓名"
function CheckEmpty(data,strname)
{	
	var msg
	if(data=="")
	{
		msg=strname+"不允许为空"
		alert(msg);
		window.event.returnValue=false;
	}
}
function CheckNumber(number,strname)
{	
	var msg
	CheckEmpty(number,strname)
	if(isNaN(number))
	{	
		msg=strname+"不是数字"
		alert(msg)
		window.event.returnValue=false;
	} 
}
function CheckDate(date)
{
	var r = date.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/); 
	if(r==null)
	{
		alert("请输入正确的日期时间格式")
		window.event.returnValue=false;
	} 
	var d = new Date(r[1], r[3]-1,r[4]);
	if(d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4])
	{
		window.event.returnValue=false;
	}
	else
	{
		alert("请输入正确的日期时间格式")
		window.event.returnValue=false;
	}
}