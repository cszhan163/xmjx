//----------------------------------------------------------------------------------------------------------------------------
//	���캯��
//----------------------------------------------------------------------------------------------------------------------------
function Reminder(CheckFile, ShowFile)
{
	this.IsFirst = true;				//�Ƿ�Ϊ�û��ĵ�һ������
	this.CheckFile = CheckFile;			//����������������������ļ�
	this.ShowFile = ShowFile;			//��ʾ���Ѵ���ʱ�������������ļ�
	this.RemindWindow = null;			//�������Ѵ��ڵ�����
	this.CheckRemind = CheckRemind;
}

//-------------------------------------------------------------------------------------------------------------------------
//	�ڿͻ�����������������󣬼��������
//-------------------------------------------------------------------------------------------------------------------------
function CheckRemind()
{
	var Dom = new ActiveXObject("MSXML2.DOMDocument");
	var Href = window.location.href;
	var CheckDate = new Date();
	var bLoadXml;
	
	//�������������ѵ��ļ���������,���Ϳͻ���ʱ��
//		debugger;
	Dom.async = false;
//	alert("CheckDate: "+ CheckDate);
//	alert(Href.substring(0, Href.lastIndexOf("/") + 1) + this.CheckFile +"?IsFirst="+ this.IsFirst +"&CheckDate="+ GetReqDate(CheckDate));
	bLoadXml = Dom.load(Href.substring(0, Href.lastIndexOf("/") + 1) + this.CheckFile +"?IsFirst="+ this.IsFirst +"&CheckDate="+ GetReqDate(CheckDate));
//	alert(bLoadXml +"\n"+ Dom.xml)
	if (!bLoadXml)
	{
		if (this.RemindWindow != null)
		{
			this.RemindWindow.close();
		}
//		alert("����ʧ�ܣ�\nMSXML.load()");
		window.setTimeout(CheckRemind.caller, (60 - CheckDate.getSeconds()) * 1000);
	}
	else
	{
		if (Dom.parseError.errorCode != 0)
		{
			if (this.RemindWindow != null)
			{
				this.RemindWindow.close();
			}
			alert("����ʧ�ܣ�\nParseError: "+ Dom.parseError.reason);
		}
		else
		{
			var Result = Dom.documentElement.firstChild.text;
			if (Result == "SessionEnd")
			{
				if (this.RemindWindow != null)
				{
					this.RemindWindow.close();
				}
				window.focus();
				alert("��������Ự��ʱ�� �����µ�½��������");
				window.top.location.replace("default.asp");
			}
			else
			{
//				alert("ok"+ Dom.documentElement.childNodes[1].text +", "+ Result);
				var Sql = Dom.documentElement.childNodes[1].text;
//				debugger;
				if (Result == "1")				//��������
				{
					this.RemindWindow = window.open(this.ShowFile +"?IsFirst="+ this.IsFirst + "&CheckDate="+ GetReqDate(CheckDate),
										"HrReminder", "left="+ (window.screen.availWidth - 450) / 2 +", "+
										"top="+ ((window.screen.availHeight - 250) / 2 + 50) +", width=450, height=250"+
										"menubar=no, toolbar=no, resizable=no", true);
				}
				else							//��������
				{
				}

				window.setTimeout(CheckRemind.caller, (60 - CheckDate.getSeconds()) * 1000);
//				window.State.innerText = CheckDate;
				this.IsFirst = false;
			}
		}
	}
}

//-------------------------------------------------------------------------------------------------------------------------
//	��JScript�е�ʱ���ʽת��Ϊ '2006/11/20 15:30:04'�ĸ�ʽ����Sqlʹ��
//-------------------------------------------------------------------------------------------------------------------------
function GetReqDate(CheckDate)
{
	var Str = new String("00");
	var rtStr;		//ȡ�õ�ǰʱ��
	
	rtStr = CheckDate.getFullYear().toString();
	
	rtStr = rtStr +"/"+ Str.substr(0, 2 - (CheckDate.getMonth() + 1).toString().length) + (CheckDate.getMonth() + 1).toString();
	rtStr = rtStr +"/"+ Str.substr(0, 2 - CheckDate.getDate().toString().length) + CheckDate.getDate().toString();
	rtStr = rtStr +" "+ CheckDate.getHours() +":"+ CheckDate.getMinutes() +":"+ CheckDate.getSeconds();

	return rtStr;
}

//--------------------------------------------------------------------------------------------------------------------------------
//	��ͣ����ʱ���ж����������ʱ���Ƿ���Ч������ȡ���ѵ�ʱ������ʱ�䵥λ
//--------------------------------------------------------------------------------------------------------------------------------
function GetDelayTime(DelayTime)
{
	var Reg, Arr
	var rt = new Object();

	Reg = new RegExp("^([\\d.]*?) *?(��|����|Сʱ|��|��)?$", "i");
	Arr = Reg.exec(DelayTime);

	if (Arr == null)
	{
		rt.Error = "��ͣʱ����Ч�� ��������Ч����ͣʱ�䡣"
	}
	else
	{
		if (isNaN(Arr[1]))
		{
			rt.Error = "��ͣʱ��������Ч�� ��������Ч����ͣʱ�䡣"
		}
		else
		{
			if (Arr[1] <= 0)
			{
				rt.Error = "��ͣʱ�䲻��Ϊ�㣡 ��������Ч����ͣʱ�䡣"
			}
			else
			{
				rt.Number = Arr[1];
				if (Arr[2] == "")
				{
					rt.Unit = "����";			//δ¼����ͣʱ�䵥λʱ��Ĭ�Ϸ���
				}
				else
				{
					rt.Unit = Arr[2];
				}
			}
		}
	}
	
	return rt;
}

//--------------------------------------------------------------------------------------------------------------------------------
//	���ݵ�ǰʱ�䣬�û��������ͣʱ�䣬�����´ε�����ʱ��
//--------------------------------------------------------------------------------------------------------------------------------
function GetExpireTime(CheckDate, Number, Unit)
{
	var iNum = Math.floor(Number);			//ʱ����ֵ����������
	var dNum = Number - iNum;				//ʱ����ֵ��С������

	switch (Unit)
	{
		case "��" :
			CheckDate.setDate(CheckDate.getDate() + iNum * 7);
			CheckDate = GetExpireTime(CheckDate, dNum * 7, "��");
			break;

		case "��" :
			CheckDate.setDate(CheckDate.getDate() + iNum);
			CheckDate = GetExpireTime(CheckDate, dNum * 24, "Сʱ");
			break;

		case "Сʱ" :
			CheckDate.setHours(CheckDate.getHours() + iNum);
			CheckDate = GetExpireTime(CheckDate, dNum * 60, "����");
			break;

		case "����" :
			CheckDate.setMinutes(CheckDate.getMinutes() + iNum);
			break;
	}

	return CheckDate;
}

//--------------------------------------------------------------------------------------------------------------------------------
//	���ݵ�ǰʱ�䣬�����б���ʾ�ĵ���ʱ��
//--------------------------------------------------------------------------------------------------------------------------------
function GetExpiredNum(List, col)
{
	var CurTime = new Date();
	var RemindDate;
	var Elapse;
	var MinMilli = 1000 * 60;
	var HrMilli = MinMilli * 60;
	var DayMilli = HrMilli * 24;
	var WeekMilli = DayMilli * 7;
	
	for (var i = 0; i < aRemindDate.length; i++)
	{

		RemindDate = new Date(Date.parse(aRemindDate[i]));
		Elapse = CurTime.valueOf() - RemindDate.valueOf();
		if (Elapse >= WeekMilli)
		{
			List.rows(i+1).cells(col).innerText = "�ѹ��� "+ Math.floor(Elapse / WeekMilli) +" ��";
		}
		else
		{
			if (Elapse > DayMilli)
			{
				List.rows(i+1).cells(col).innerText = "�ѹ��� "+ Math.floor(Elapse / DayMilli) +" ��";
			}
			else
			{
				if (Elapse > HrMilli)
				{
					List.rows(i+1).cells(col).innerText = "�ѹ��� "+ Math.floor(Elapse / HrMilli) +" Сʱ";
				}
				else
				{
					if (Elapse > MinMilli)
					{
						List.rows(i+1).cells(col).innerText = "�ѹ��� "+ Math.floor(Elapse / MinMilli) +" ��";
					}
					else
					{
						List.rows(i+1).cells(col).innerText = "����";
					}
				}
			}
		}
	}
	
	window.setTimeout(Function("GetExpiredNum("+ List.id +", "+ col +")"), (60 - CurTime.getSeconds()) * 1000);
}