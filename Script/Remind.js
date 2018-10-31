//----------------------------------------------------------------------------------------------------------------------------
//	构造函数
//----------------------------------------------------------------------------------------------------------------------------
function Reminder(CheckFile, ShowFile)
{
	this.IsFirst = true;				//是否为用户的第一次提醒
	this.CheckFile = CheckFile;			//检测新提醒向服务器请求的文件
	this.ShowFile = ShowFile;			//显示提醒窗口时向服务器请求的文件
	this.RemindWindow = null;			//保存提醒窗口的引用
	this.CheckRemind = CheckRemind;
}

//-------------------------------------------------------------------------------------------------------------------------
//	在客户端向服务器发送请求，检测新提醒
//-------------------------------------------------------------------------------------------------------------------------
function CheckRemind()
{
	var Dom = new ActiveXObject("MSXML2.DOMDocument");
	var Href = window.location.href;
	var CheckDate = new Date();
	var bLoadXml;
	
	//向服务器检测提醒的文件发送请求,发送客户端时间
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
//		alert("提醒失败！\nMSXML.load()");
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
			alert("提醒失败！\nParseError: "+ Dom.parseError.reason);
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
				alert("与服务器会话超时！ 请重新登陆服务器。");
				window.top.location.replace("default.asp");
			}
			else
			{
//				alert("ok"+ Dom.documentElement.childNodes[1].text +", "+ Result);
				var Sql = Dom.documentElement.childNodes[1].text;
//				debugger;
				if (Result == "1")				//有新提醒
				{
					this.RemindWindow = window.open(this.ShowFile +"?IsFirst="+ this.IsFirst + "&CheckDate="+ GetReqDate(CheckDate),
										"HrReminder", "left="+ (window.screen.availWidth - 450) / 2 +", "+
										"top="+ ((window.screen.availHeight - 250) / 2 + 50) +", width=450, height=250"+
										"menubar=no, toolbar=no, resizable=no", true);
				}
				else							//无新提醒
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
//	把JScript中的时间格式转换为 '2006/11/20 15:30:04'的格式，供Sql使用
//-------------------------------------------------------------------------------------------------------------------------
function GetReqDate(CheckDate)
{
	var Str = new String("00");
	var rtStr;		//取得当前时间
	
	rtStr = CheckDate.getFullYear().toString();
	
	rtStr = rtStr +"/"+ Str.substr(0, 2 - (CheckDate.getMonth() + 1).toString().length) + (CheckDate.getMonth() + 1).toString();
	rtStr = rtStr +"/"+ Str.substr(0, 2 - CheckDate.getDate().toString().length) + CheckDate.getDate().toString();
	rtStr = rtStr +" "+ CheckDate.getHours() +":"+ CheckDate.getMinutes() +":"+ CheckDate.getSeconds();

	return rtStr;
}

//--------------------------------------------------------------------------------------------------------------------------------
//	暂停提醒时，判断输入的提醒时间是否有效，并提取提醒的时间间隔和时间单位
//--------------------------------------------------------------------------------------------------------------------------------
function GetDelayTime(DelayTime)
{
	var Reg, Arr
	var rt = new Object();

	Reg = new RegExp("^([\\d.]*?) *?(分|分钟|小时|天|周)?$", "i");
	Arr = Reg.exec(DelayTime);

	if (Arr == null)
	{
		rt.Error = "暂停时间无效！ 请输入有效的暂停时间。"
	}
	else
	{
		if (isNaN(Arr[1]))
		{
			rt.Error = "暂停时间数字无效！ 请输入有效的暂停时间。"
		}
		else
		{
			if (Arr[1] <= 0)
			{
				rt.Error = "暂停时间不能为零！ 请输入有效的暂停时间。"
			}
			else
			{
				rt.Number = Arr[1];
				if (Arr[2] == "")
				{
					rt.Unit = "分钟";			//未录入暂停时间单位时，默认分钟
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
//	根据当前时间，用户输入的暂停时间，计算下次的提醒时间
//--------------------------------------------------------------------------------------------------------------------------------
function GetExpireTime(CheckDate, Number, Unit)
{
	var iNum = Math.floor(Number);			//时间数值的整数部分
	var dNum = Number - iNum;				//时间数值的小数部分

	switch (Unit)
	{
		case "周" :
			CheckDate.setDate(CheckDate.getDate() + iNum * 7);
			CheckDate = GetExpireTime(CheckDate, dNum * 7, "天");
			break;

		case "天" :
			CheckDate.setDate(CheckDate.getDate() + iNum);
			CheckDate = GetExpireTime(CheckDate, dNum * 24, "小时");
			break;

		case "小时" :
			CheckDate.setHours(CheckDate.getHours() + iNum);
			CheckDate = GetExpireTime(CheckDate, dNum * 60, "分钟");
			break;

		case "分钟" :
			CheckDate.setMinutes(CheckDate.getMinutes() + iNum);
			break;
	}

	return CheckDate;
}

//--------------------------------------------------------------------------------------------------------------------------------
//	根据当前时间，计算列表显示的到期时间
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
			List.rows(i+1).cells(col).innerText = "已过期 "+ Math.floor(Elapse / WeekMilli) +" 周";
		}
		else
		{
			if (Elapse > DayMilli)
			{
				List.rows(i+1).cells(col).innerText = "已过期 "+ Math.floor(Elapse / DayMilli) +" 天";
			}
			else
			{
				if (Elapse > HrMilli)
				{
					List.rows(i+1).cells(col).innerText = "已过期 "+ Math.floor(Elapse / HrMilli) +" 小时";
				}
				else
				{
					if (Elapse > MinMilli)
					{
						List.rows(i+1).cells(col).innerText = "已过期 "+ Math.floor(Elapse / MinMilli) +" 分";
					}
					else
					{
						List.rows(i+1).cells(col).innerText = "立即";
					}
				}
			}
		}
	}
	
	window.setTimeout(Function("GetExpiredNum("+ List.id +", "+ col +")"), (60 - CurTime.getSeconds()) * 1000);
}