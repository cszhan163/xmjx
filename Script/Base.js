//=============================================
//  版本: 1.0.5.0
//  作者: 韩基伟
//=============================================
var $=function(id){return document.getElementById(id);}
window.onerror=function(){return false;}

document.write("<style type=\"text/css\">.TextEdit{border:none}</style>");

//编辑文本框
   var EditTextBox =function(obj)
   {    
       obj.contentEditable="true";  
       obj.className="";
       obj.style.width=99;
       obj.style.textAlign="left";
       if($("Submit1"))
       {       
            if($("Submit1").innerText==" 保 存 ")
               $("Submit1").innerText=" 添 加 ";
            else 
              $("Submit1").innerText=" 保 存 ";
              
      }
       obj.onblur=function()
        {
          if($("Submit1")) $("Submit1").innerText=" 添 加 ";
            obj.contentEditable="false"; 
            obj.style.textAlign="center"; 
            obj.className="TextEdit";
        }
   }
   //去除两端空格
  String.prototype.Trim = function() 
{ 
    return this.replace(/(^\s*)|(\s*$)/g, ""); 
}
 //去除左端空格
String.prototype.LTrim = function() 
{ 
return this.replace(/(^\s*)/g, ""); 
} 
 //去除右端空格
String.prototype.RTrim = function() 
{ 
return this.replace(/(\s*$)/g, ""); 
} 

//设置textbox的边框颜色
//try{
//    var obj=document.body.getElementsByTagName("input");
//    for(i=0;i<obj.length;i++)
//    { 
//        if(obj[i].outerHTML.indexOf("type")<1)
//        {    
//           obj[i].style.border="solid 1 groove";
//        } 
//    }
//}
//catch(e){}

 //数值格式化函数，Dight要 格式化的  数字，How要保留的小数位数。  
  function  ForDight(Dight,How)  
  {  
      Dight  = parseFloat(Dight); 
      return Dight.toFixed(How);
  } 
    //判断输入的是否是数字
  function IsNumreic(TextObj) 
{  
    var i,j,strTemp,str; 
    str=TextObj.value; 
    strTemp="0123456789."; 
    if ( str.length== 0) 
    return false; 
    for (i=0;i<str.length;i++) 
    {  j=strTemp.indexOf(str.charAt(i)); 
        if (j==-1) 
        {//说明有字符不是数字
            TextObj.value=str.replace(str.charAt(i),"");
            return false; 
        } 
    } 
    //说明是数字 
    return true; 
}
//////////////////////////////////////rxf
  function showInfo(obj,layerid)   
  {   
  var objleft=obj.getBoundingClientRect().left
  var objtop=obj.getBoundingClientRect().top
  document.getElementById(layerid).style.left=objleft-250
  document.getElementById(layerid).style.top=objtop-10
  document.getElementById(layerid).style.display="block"
  }
  function hiddeninfo(layerid)
  {
  document.getElementById(layerid).style.display="none"
  }   
//////////////////////////////////////