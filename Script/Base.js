//=============================================
//  �汾: 1.0.5.0
//  ����: ����ΰ
//=============================================
var $=function(id){return document.getElementById(id);}
window.onerror=function(){return false;}

document.write("<style type=\"text/css\">.TextEdit{border:none}</style>");

//�༭�ı���
   var EditTextBox =function(obj)
   {    
       obj.contentEditable="true";  
       obj.className="";
       obj.style.width=99;
       obj.style.textAlign="left";
       if($("Submit1"))
       {       
            if($("Submit1").innerText==" �� �� ")
               $("Submit1").innerText=" �� �� ";
            else 
              $("Submit1").innerText=" �� �� ";
              
      }
       obj.onblur=function()
        {
          if($("Submit1")) $("Submit1").innerText=" �� �� ";
            obj.contentEditable="false"; 
            obj.style.textAlign="center"; 
            obj.className="TextEdit";
        }
   }
   //ȥ�����˿ո�
  String.prototype.Trim = function() 
{ 
    return this.replace(/(^\s*)|(\s*$)/g, ""); 
}
 //ȥ����˿ո�
String.prototype.LTrim = function() 
{ 
return this.replace(/(^\s*)/g, ""); 
} 
 //ȥ���Ҷ˿ո�
String.prototype.RTrim = function() 
{ 
return this.replace(/(\s*$)/g, ""); 
} 

//����textbox�ı߿���ɫ
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

 //��ֵ��ʽ��������DightҪ ��ʽ����  ���֣�HowҪ������С��λ����  
  function  ForDight(Dight,How)  
  {  
      Dight  = parseFloat(Dight); 
      return Dight.toFixed(How);
  } 
    //�ж�������Ƿ�������
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
        {//˵�����ַ���������
            TextObj.value=str.replace(str.charAt(i),"");
            return false; 
        } 
    } 
    //˵�������� 
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