//javascript��֤�ͻ������ݡ��Ƿ�Ϊ�գ��Ƿ����֣��Ƿ�����ʱ��
//strname��֤����������"����"
function CheckEmpty(data,strname)
{	
	var msg
	if(data=="")
	{
		msg=strname+"������Ϊ��"
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
		msg=strname+"��������"
		alert(msg)
		window.event.returnValue=false;
	} 
}
function CheckDate(date)
{
	var r = date.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/); 
	if(r==null)
	{
		alert("��������ȷ������ʱ���ʽ")
		window.event.returnValue=false;
	} 
	var d = new Date(r[1], r[3]-1,r[4]);
	if(d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4])
	{
		window.event.returnValue=false;
	}
	else
	{
		alert("��������ȷ������ʱ���ʽ")
		window.event.returnValue=false;
	}
}