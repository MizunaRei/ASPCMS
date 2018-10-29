<%
dim arrPurview(10),PurviewIndex,strThisFile
strThisFile=lcase(request.ServerVariables("SCRIPT_NAME"))
strThisFile=mid(strThisFile,instrrev(strThisFile,"/")+1,30)
if ShowPopMenu="Yes" and strThisFile<>"admin_index_left.asp" and strThisFile<>"admin_articlecontent.asp" then
%>
<head><STYLE>
.menutable {
	BORDER-RIGHT: #39867B 1px solid; BORDER-TOP: #39867B 1px solid; FONT-SIZE: 12px; Z-INDEX: 100; BORDER-LEFT: #39867B 6px solid; BORDER-BOTTOM: #39867B 1px solid; POSITION: absolute; BACKGROUND-COLOR: #ffffff
}
.menutrin {
	CURSOR: hand; COLOR: #0000ff; BACKGROUND-COLOR: #E1F4EE
}
.menutrout {
	CURSOR: hand; COLOR: #000000
}
.menutd0 {
	WIDTH: 28px; HEIGHT: 22px; TEXT-ALIGN: center; 改变这个修改菜单高度---: 
}
.menutd1 {
	WIDTH: 22px; FONT-FAMILY: Webdings; TEXT-ALIGN: right
}
.linktd1 {
	WIDTH: 22px
}
.menutd2 {
	WIDTH: 4px
}
.menuhr {
	BORDER-RIGHT: #39867B 1px inset; BORDER-TOP: #39867B 1px inset; BORDER-LEFT: #39867B 1px inset; BORDER-BOTTOM: #39867B 1px inset
}
</STYLE>
<!--<BGSOUND id=theBS src="" loop=0></head>-->
<body><script>
<!----

/*-----------------------------------------------------------
鼠标右键菜单 1.0 Designed By Stroll  e-mail: csy-163@163.com



--------------------------------------------------------------*/

//---------------  有关数据 -----------------//

var IconList = new Array();   // icon图片 集合， 下标从 1 开始

	IconList[1] = new Image();
	IconList[1].src = "images/popmenu/001.gif";
	IconList[2] = new Image();
	IconList[2].src = "images/popmenu/002.gif";
	IconList[3] = new Image();
	IconList[3].src = "images/popmenu/003.gif";	
	IconList[4] = new Image();
	IconList[4].src = "images/popmenu/004.gif";	
	IconList[5] = new Image();
	IconList[5].src = "images/popmenu/005.gif";	
	IconList[6] = new Image();
	IconList[6].src = "images/popmenu/0061.gif";	
	IconList[7] = new Image();
	IconList[7].src = "images/popmenu/0062.gif";	
	IconList[8] = new Image();
	IconList[8].src = "images/popmenu/0063.gif";	
	IconList[9] = new Image();
	IconList[9].src = "images/popmenu/007.gif";	
	IconList[10] = new Image();
	IconList[10].src = "images/popmenu/008.gif";	
	IconList[11] = new Image();
	IconList[11].src = "images/popmenu/009.gif";	
	IconList[12] = new Image();
	IconList[12].src = "images/popmenu/010.gif";	
	IconList[13] = new Image();
	IconList[13].src = "images/popmenu/011.gif";	
	IconList[14] = new Image();
	IconList[14].src = "images/popmenu/012.gif";
	IconList[15] = new Image();
	IconList[15].src = "images/popmenu/013.gif";
	IconList[16] = new Image();
	IconList[16].src = "images/popmenu/014.gif";	

//----------------  检测变量 菜单的显示隐藏就靠它了！！！  ------------------//	

var JustMenuID = "";
var SubMenuList = new Array();
var NowSubMenu = "";	
var mouseCanSound = false;	  //---------------------------  声音开关 ------  声音开关 ------------------//
var menuSpeed     =  50;   //---------- 菜单显示速度 ------------//
var alphaStep     =  50;   //---------- Alpaha 变化 度 -----------//
var index         =   -1;
//---------------  菜单内容 -----------------//
var one = new MouseMenu("one");
<%if AdminPurview=1 or AdminPurview_Article<=3 then%>
one.addMenu("文章管理",6);
index+=1
	one[index].addLink("添加文章（简洁模式）","","Admin_ArticleAdd1.asp")
	one[index].addLink("添加文章（高级模式）","","Admin_ArticleAdd2.asp")
	one[index].addLink("我发表的文章","","Admin_ArticleManage.asp?ManageType=MyArticle")
	one[index].addLink("文章管理","","Admin_ArticleManage.asp")
	one[index].addLink("文章审核","","Admin_ArticleCheck.asp")
<%if AdminPurview=1 or AdminPurview_Article=1 then%>
	one[index].addLink("专题文章管理","","Admin_ArticleManageSpecial.asp")
	one[index].addLink("文章回收站管理","","Admin_ArticleRecyclebin.asp")
	one[index].addLink("文章评论管理","","Admin_ArticleComment.asp")
	one[index].addLink("文章栏目添加","","Admin_Class_Article.asp?Action=Add")
	one[index].addLink("文章栏目管理","","Admin_Class_Article.asp")
	one[index].addLink("文章专题添加","","Admin_Special.asp?Action=Add")
	one[index].addLink("文章专题管理","","Admin_Special.asp")
<%
	end if
end if
if AdminPurview=1 or AdminPurview_Soft<=3 then
%>
one.addMenu("软件管理",7);
index+=1
	one[index].addLink("添加软件","","Admin_SoftAdd.asp")
	one[index].addLink("我添加的软件","","Admin_SoftManage.asp?ManageType=MySoft")
	one[index].addLink("软件管理","","Admin_SoftManage.asp")
	one[index].addLink("软件审核","","Admin_SoftCheck.asp")
<%if AdminPurview=1 or AdminPurview_Soft=1 then%>
	one[index].addLink("软件回收站管理","","Admin_SoftRecyclebin.asp")
	one[index].addLink("软件评论管理","","Admin_SoftComment.asp")
	one[index].addLink("下载栏目添加","","Admin_Class_Soft.asp?Action=Add")
	one[index].addLink("下载栏目管理","","Admin_Class_Soft.asp")
<%
	end if
end if
if AdminPurview=1 or AdminPurview_Photo<=3 then
%>
one.addMenu("图片管理",8);
index+=1
	one[index].addLink("添加图片","","Admin_PhotoAdd.asp")
	one[index].addLink("我添加的图片","","Admin_PhotoManage.asp?ManageType=MyPhoto")
	one[index].addLink("图片管理","","Admin_PhotoManage.asp")
	one[index].addLink("图片审核","","Admin_PhotoCheck.asp")
<%if AdminPurview=1 or AdminPurview_Photo=1 then%>
	one[index].addLink("图片回收站管理","","Admin_PhotoRecyclebin.asp")
	one[index].addLink("图片评论管理","","Admin_PhotoComment.asp")
	one[index].addLink("图片栏目添加","","Admin_Class_Photo.asp?Action=Add")
	one[index].addLink("图片栏目管理","","Admin_Class_Photo.asp")
<%
	end if
end if

PurviewPassed=False
arrPurview(0)=CheckPurview(AdminPurview_Others,"Channel")
arrPurview(1)=CheckPurview(AdminPurview_Others,"AD")
arrPurview(2)=CheckPurview(AdminPurview_Others,"FriendSite")
arrPurview(3)=CheckPurview(AdminPurview_Others,"Announce")
arrPurview(4)=CheckPurview(AdminPurview_Others,"Vote")
arrPurview(5)=CheckPurview(AdminPurview_Others,"Count")
for PurviewIndex=0 to 5
	if arrPurview(PurviewIndex)=True then
		PurviewPassed=True
		exit for
	end if
next
if AdminPurview=1 or PurviewPassed=True then
%>
one.addMenu("常规设置",9);
index+=1
<%if AdminPurview=1 then%>
	one[index].addLink("网站信息配置","","Admin_SiteConfig.asp")
<%
end if
if AdminPurview=1 or arrPurview(0)=True then
%>			  
	one[index].addLink("网站频道添加","","Admin_Channel.asp?Action=Add")
	one[index].addLink("网站频道管理","","Admin_Channel.asp")
<%
end if
if AdminPurview=1 or arrPurview(1)=True then
%>			  
	one[index].addLink("网站广告添加","","Admin_Advertisement.asp?Action=Add")
	one[index].addLink("网站广告管理","","Admin_Advertisement.asp")
<%
end if
if AdminPurview=1 or arrPurview(2)=True then
%>			  
	one[index].addLink("友情链接添加","","Admin_FriendSite.asp?Action=Add")
	one[index].addLink("友情链接管理","","Admin_FriendSite.asp")
<%
end if
if AdminPurview=1 or arrPurview(3)=True then
%>			  
	one[index].addLink("发布网站公告","","Admin_Announce.asp?Action=Add")
	one[index].addLink("网站公告管理","","Admin_Announce.asp")
<%
end if
if AdminPurview=1 or arrPurview(4)=True then
%>			  
	one[index].addLink("发布网站调查","","Admin_Vote.asp?Action=Add")
	one[index].addLink("网站调查管理","","Admin_Vote.asp")
<%
end if
if AdminPurview=1 or arrPurview(5)=True then
%>			  
	one[index].addLink("网站统计系统","","Count/main.asp")
	one[index].addLink("网站统计管理","","Count/index.asp")
<%end if%>
<%
end if
PurviewPassed=False
arrPurview(0)=CheckPurview(AdminPurview_Guest,"Reply")
arrPurview(1)=CheckPurview(AdminPurview_Guest,"Modify")
arrPurview(2)=CheckPurview(AdminPurview_Guest,"Del")
arrPurview(3)=CheckPurview(AdminPurview_Guest,"Check")
for PurviewIndex=0 to 3
	if arrPurview(PurviewIndex)=True then
		PurviewPassed=True
		exit for
	end if
next
if AdminPurview=1 or PurviewPassed=True then
%>
one.addMenu("留言板管理",10)
index+=1
<%if AdminPurview=1 or arrPurview(3)=True then%>
	one[index].addLink("网站留言审核","","Admin_Guest.asp")
<%end if%>
	one[index].addLink("网站留言管理","","Admin_Guest.asp")
<%
end if

PurviewPassed=False
arrPurview(0)=CheckPurview(AdminPurview_Others,"User")
arrPurview(1)=CheckPurview(AdminPurview_Others,"MailList")
for PurviewIndex=0 to 1
	if arrPurview(PurviewIndex)=True then
		PurviewPassed=True
		exit for
	end if
next
if AdminPurview=1 or PurviewPassed=True then
%>
one.addMenu("用户管理",11)
index+=1
<%if AdminPurview=1 or arrPurview(0)=True then%>
	one[index].addLink("注册用户管理","","Admin_User.asp")
	one[index].addLink("更新用户数据","","Admin_User.asp?Action=Update")
<%
end if
if AdminPurview=1 or arrPurview(2)=True then
%>
	one[index].addLink("邮件列表","","Admin_Maillist.asp")
	one[index].addLink("列表导出","","Admin_Maillist.asp?Action=Export")
<%
end if
if AdminPurview=1 then
%>
	one[index].addLink("管理员添加","","Admin_Admin.asp?Action=Add")
	one[index].addLink("管理员管理","","Admin_Admin.asp")
<%end if%>
<%
end if
PurviewPassed=False
arrPurview(0)=CheckPurview(AdminPurview_Others,"Skin")
arrPurview(1)=CheckPurview(AdminPurview_Others,"Layout")
for PurviewIndex=0 to 1
	if arrPurview(PurviewIndex)=True then
		PurviewPassed=True
		exit for
	end if
next
if AdminPurview=1 or PurviewPassed=True then
%>
one.addMenu("模板管理",12)	
index+=1
<%if AdminPurview=1 or arrPurview(0)=True then%>
	one[index].addLink("配色模板添加","","Admin_Skin.asp?Action=Add")
	one[index].addLink("配色模板管理 ","","Admin_Skin.asp")
	one[index].addLink("配色模板导出","","Admin_Skin.asp?Action=Export")
	one[index].addLink("配色模板导入","","Admin_Skin.asp?Action=Import")
<%
end if
if AdminPurview=1 or arrPurview(1)=True then
%>
	one[index].addLink("版面设计添加","","Admin_Layout.asp?Action=Add")
	one[index].addLink("版面设计管理","","Admin_Layout.asp")
	one[index].addLink("版面设计导出","","Admin_Temp.asp")
	one[index].addLink("版面设计导入 ","","Admin_Temp.asp")
<%end if%>
<%
end if
if AdminPurview=1 or CheckPurview(AdminPurview_Others,"JS")=True then
%>
one.addMenu("JS代码管理",13)
index+=1
	one[index].addLink("普通文章的JS代码","","Admin_MakeJS.asp?Action=JS_Common")
	one[index].addLink("首页图文的JS代码","","Admin_MakeJS.asp?Action=JS_Pic")
<%
end if
if AdminPurview=1 or CheckPurview(AdminPurview_Others,"Database")=True then
%>
one.addMenu("数据库管理",14)
index+=1
	one[index].addLink("压缩数据库","","Admin_Database.asp?Action=Compact")
	one[index].addLink("备份数据库","","Admin_Database.asp?Action=Backup")
	one[index].addLink("恢复数据库","","Admin_Database.asp?Action=Restore")
	one[index].addLink("系统初始化","","Admin_Database.asp?Action=Init")
	one[index].addLink("系统空间占用","","Admin_Database.asp?Action=SpaceSize")
<%end if
if AdminPurview=1 or CheckPurview(AdminPurview_Others,"UpFile")=True then
%>
one.addMenu("上传文件管理",15)
index+=1
	one[index].addLink("文章频道的上传文件","","Admin_UploadFile.asp?UploadDir=UploadFiles")
	one[index].addLink("图片频道的缩略图","","Admin_UploadFile.asp?UploadDir=UploadThumbs")
	one[index].addLink("图片频道的上传图片","","Admin_UploadFile.asp?UploadDir=UploadPhotos")
	one[index].addLink("下载频道的软件图片","","Admin_UploadFile.asp?UploadDir=UploadSoftPic")
	one[index].addLink("下载频道的上传软件","","Admin_UploadFile.asp?UploadDir=UploadSoft")
	one[index].addLink("清除无用文件","","Admin_UploadFile.asp?Action=Clear")
<%end if%>
one.addHR();
one.addLink("刷新",5,"javascript:location.reload()","main")	
one.addLink("前进",4,"javascript:history.go(1)","main")	
one.addLink("后退",3,"javascript:history.back(1)","main")	
one.addHR();
one.addLink("修改密码",2,"Admin_AdminModifyPwd.asp","main")	
one.addLink("管理首页",1,"Admin_Index_Main.asp","main")	





//------------- 构建 主菜单 对象 -------------//

function MouseMenu(objName)
{
	this.id  	 = "Menu_"+objName;
	this.obj 	 = objName;
	this.length  = 0;
	
	
	this.addMenu = addMenu;
	this.addLink = addLink;
	this.addHR   = addHR;	
	
	JustMenuID = this.id;
 	document.body.insertAdjacentHTML('beforeEnd','<table id="'+this.id+'" border="0" cellspacing="0" cellpadding="0" style="top: 0; left: 0; visibility: hidden; filter:Alpha(Opacity=0);" class="menutable" onmousedown=event.cancelBubble=true; onmouseup=event.cancelBubble=true></table>');
}

//----------- 构建 子菜单 对象 -------------//

function SubMenu(objName,objID)
{
	this.obj = objName;
	this.id  = objID;

	this.addMenu = addMenu;
	this.addLink = addLink;
	this.addHR   = addHR;

	this.length  = 0;
}


//-------------- 生成 菜单 makeMenu 方法 -----------//
function makeMenu(subID,oldID,word,icon,url,target,thetitle)
{
	var thelink = '';
	

	if(icon&&icon!="")
	{
		icon = '<img border="0" src="'+IconList[icon].src+'">';
	}
	else
	{
		icon = '';
	}
	
	if(!thetitle||thetitle=="")
	{
		thetitle = '';
	}
	
	
	if(url&&url!="")
	{
		thelink += '<a href="'+url+'" ';
		
		if(target&&target!="")
		{
			thelink += '  ';
			thelink += 'target="'+target+'" '
		}
		
		thelink += '></a>';
	}
	
	var Oobj = document.getElementById(oldID);

	/*--------------------------------------------- 菜单html样式
	
	  <tr class="menutrout" id="trMenu_one_0" title="I am title">
      <td class="menutd0"><img src="icon/sub.gif" border="0" width="16" height="16"></td>
      <td><a href="javascript:alert('I am menu');" target="_self"></a><nobr>菜单一</nobr></td>
      <td class="menutd1">4</td>
      <td class="menutd2">&nbsp;</td>
    </tr>

	
	--------------------------------------------------*/
	
	Oobj.insertRow();
	

	with(Oobj.rows(Oobj.rows.length-1))
	{
		id 			= "tr"+subID;
		className	= "menutrout";
		
		title       = thetitle;

	}
	
	eventObj = "tr"+subID;
	
	eval(eventObj+'.attachEvent("onmouseover",MtrOver('+eventObj+'))');	
	eval(eventObj+'.attachEvent("onclick",MtrClick('+eventObj+'))');	
		
	var trObj = eval(eventObj);

	for(i=0;i<4;i++)
	{
		trObj.insertCell();
	}

	with(Oobj.rows(Oobj.rows.length-1))
	{
		cells(0).className = "menutd0";
		cells(0).innerHTML = icon;

		cells(1).innerHTML = thelink+'<nobr class=indentWord>'+word+'</nobr>';
		cells(1).calssName = "indentWord"
		
		cells(2).className = "menutd1";
		cells(2).innerHTML = "4";
		
		cells(3).className = "menutd2";
		cells(3).innerText = " ";
		
	}	
	
	
	
	document.body.insertAdjacentHTML('beforeEnd','<table id="'+subID+'" border="0" cellspacing="0" cellpadding="0" style="top: 0; left: 0; visibility: hidden; filter:Alpha(Opacity=0);" class="menutable" onmousedown=event.cancelBubble=true; onmouseup=event.cancelBubble=true></table>');
	
	
		
}


//---------------- 生成连接 makeLink 方法 ------------//
function makeLink(subID,oldID,word,icon,url,target,thetitle)
{
	
	
	var thelink = '';
	
	if(icon&&icon!="")
	{
		icon = '<img border="0" src="'+IconList[icon].src+'">';
	}
	else
	{
		icon = '';
	}
	
	if(!thetitle||thetitle=="")
	{
		thetitle = '';
	}
	
	
	if(url&&url!="")
	{
		thelink += '<a href="'+url+'" ';
		
		if(target&&target!="")
		{
			thelink += '  ';
			thelink += 'target="'+target+'" '
		}
		
		thelink += '></a>';
	}
	
	var Oobj = document.getElementById(oldID);
	
	
	/*--------------------------------------------- 连接html样式
	
	  <tr class="menutrout" id="trMenu_one_0" title="I am title">
      <td class="menutd0"><img src="icon/sub.gif" border="0" width="16" height="16"></td>
      <td><a href="javascript:alert('I am link');" target="_self"></a><nobr>连接一</nobr></td>
      <td class="linktd1"></td>
      <td class="menutd2">&nbsp;</td>
    </tr>

	
	--------------------------------------------------*/	
	
	Oobj.insertRow();
	

	with(Oobj.rows(Oobj.rows.length-1))
	{
		id 			= "tr"+subID;
		className	= "menutrout";		
		title       = thetitle;

	}
	
	eventObj = "tr"+subID;
	
	eval(eventObj+'.attachEvent("onmouseover",LtrOver('+eventObj+'))');	
	eval(eventObj+'.attachEvent("onmouseout",LtrOut('+eventObj+'))');		
	eval(eventObj+'.attachEvent("onclick",MtrClick('+eventObj+'))');	
		
	var trObj = eval(eventObj);

	for(i=0;i<4;i++)
	{
		trObj.insertCell();
	}

	with(Oobj.rows(Oobj.rows.length-1))
	{
		cells(0).className = "menutd0";
		cells(0).innerHTML = icon;

		cells(1).innerHTML = thelink+'<nobr class=indentWord>'+word+'</nobr>';

		cells(2).className = "linktd1";
		cells(2).innerText = " ";
		
		cells(3).className = "menutd2";
		cells(3).innerText = " ";
		
	}	

}


//-------------- 菜单对象 addMenu 方法 ------------//
function addMenu(word,icon,url,target,title)
{
	var subID    = this.id + "_" + this.length;
	var subObj  = this.obj+"["+this.length+"]";
	
	var oldID   = this.id;
	
	eval(subObj+"= new SubMenu('"+subObj+"','"+subID+"')");
	
	 makeMenu(subID,oldID,word,icon,url,target,title);
	 
	 this.length++;
	
}


//------------- 菜单对象 addLink 方法 -------------//
function addLink(word,icon,url,target,title)
{
	var subID    = this.id + "_" + this.length;
	var oldID  = this.id;
	
	 makeLink(subID,oldID,word,icon,url,target,title);
	 
	 this.length++;	
}

//------------ 菜单对象 addHR 方法 -----------------//
function addHR()
{
	var oldID = this.id;

	var Oobj = document.getElementById(oldID);
	
	Oobj.insertRow();
	
	/*------------------------------------------
	
	 <tr>
      <td colspan="4">
		<hr class="menuhr" size="1" width="95%">
       </td>
    </tr>
	
	--------------------------------------------*/	

	
	Oobj.rows(Oobj.rows.length-1).insertCell();

	with(Oobj.rows(Oobj.rows.length-1))
	{
		cells(0).colSpan= 4;
		cells(0).insertAdjacentHTML('beforeEnd','<hr class="menuhr" size="1" width="95%">');		
	}	
	
}






//--------- MtrOver(obj)-------------------//
function MtrOver(obj)
{
	return sub_over;
	
	function sub_over()
	{
	
		var sonid = obj.id.substring(2,obj.id.length);
		
		var topobj = obj.parentElement.parentElement; 
		
		NowSubMenu = topobj.id;
		
		if(obj.className=="menutrout")
		{
			mouseWave();
		}		
		
		HideMenu(1);		
		
		SubMenuList[returnIndex(NowSubMenu)] = NowSubMenu;

		ShowTheMenu(sonid,MPreturn(sonid))		
		
		SubMenuList[returnIndex(obj.id)] = sonid;
		
		if(topobj.oldTR)
		{ 
			eval(topobj.oldTR+'.className = "menutrout"'); 
		} 

		obj.className = "menutrin"; 

		topobj.oldTR = obj.id; 
		

	}
}

//--------- LtrOver(obj)-------------------//
function LtrOver(obj)
{
	return sub_over;
	
	function sub_over()
	{
		var topobj = obj.parentElement.parentElement; 

		NowSubMenu = topobj.id;
		
		HideMenu(1);
		
		SubMenuList[returnIndex(NowSubMenu)] = NowSubMenu;
				
		if(topobj.oldTR)
		{ 
			eval(topobj.oldTR+'.className = "menutrout"'); 
		} 

		obj.className = "menutrin"; 

		topobj.oldTR = obj.id; 

	}
}

//--------- LtrOut(obj)-------------------//
function LtrOut(obj)
{
	return sub_out;
	
	function sub_out()
	{
		var topobj = obj.parentElement.parentElement; 
		
		obj.className = "menutrout"; 
		
		topobj.oldTR = false; 
	}
}

//----------MtrClick(obj)-----------------//

function MtrClick(obj)
{
	return sub_click;
	
	function sub_click()
	{
		if(obj.cells(1).all.tags("A").length>0)
		{
			obj.cells(1).all.tags("A")(0).click();
		}	

	}
}


//---------- returnIndex(str)--------------//

function returnIndex(str)
{
	return (str.split("_").length-3)
}


//---------ShowTheMenu(obj,num)-----------------//

function ShowTheMenu(obj,num)
{
	var topobj = eval(obj.substring(0,obj.length-2));
	
	var trobj  = eval("tr"+obj);
	
	var obj = eval(obj);
	
	if(num==0)
	{
		with(obj.style)
		{
			pixelLeft = topobj.style.pixelLeft +topobj.offsetWidth;
			pixelTop  = topobj.style.pixelTop + trobj.offsetTop;
		}
	}
	if(num==1)
	{
		with(obj.style)
		{
			pixelLeft = topobj.style.pixelLeft + topobj.offsetWidth;
			pixelTop  = topobj.style.pixelTop  + trobj.offsetTop + trobj.offsetHeight - obj.offsetHeight;
		}
	}
	if(num==2)
	{
		with(obj.style)
		{
			pixelLeft = topobj.style.pixelLeft -  obj.offsetWidth;
			pixelTop  = topobj.style.pixelTop + trobj.offsetTop;
		}	
	}
	if(num==3)
	{
		with(obj.style)
		{
			pixelLeft = topobj.style.pixelLeft -  obj.offsetWidth;
			pixelTop  = topobj.style.pixelTop  + trobj.offsetTop + trobj.offsetHeight - obj.offsetHeight;
		}	
	}
	
	obj.style.visibility  = "visible"; 	
	
	if(obj.alphaing)
	{
		clearInterval(obj.alphaing);
	}
	
	obj.alphaing = setInterval("menu_alpha_up("+obj.id+","+alphaStep+")",menuSpeed);	
}

//----------HideMenu(num)-------------------//

/*----------------------
var SubMenuList = new Array();

var NowSubMenu = "";	

---------------------*/

function HideMenu(num)
{
	var thenowMenu = "";
	
	var obj = null;
	
	if(num==1)
	{
		thenowMenu = NowSubMenu
	}
	
	
	
	for(i=SubMenuList.length-1;i>=0;i--)
	{
		if(SubMenuList[i]&&SubMenuList[i]!=thenowMenu)
		{
			
			obj = eval(SubMenuList[i]);
			
			if(obj.alphaing)
			{
				clearInterval(obj.alphaing);
			}	

			obj.alphaing = setInterval("menu_alpha_down("+obj.id+","+alphaStep+")",menuSpeed);
			
			obj.style.visibility = "hidden";		
			
			eval("tr"+SubMenuList[i]).className = "menutrout";
						
			SubMenuList[i] = null;	
		}
		else
		{
			if(SubMenuList[i]==thenowMenu)
			{
				return;
			}
		}
	}
	
	NowSubMenu = "";
}




//-----------MainMenuPosition return()------------//

function MMPreturn()
{
	var obj = eval(JustMenuID);
	
	var x = event.clientX;
	var y = event.clientY;
	
	var judgerX = x + obj.offsetWidth;
	var judgerY = y + obj.offsetHeight;

	var px = 0;
	var py = 0;
	
	if(judgerX>document.body.clientWidth)
	{
		px = 2;
	}
	if(judgerY>document.body.clientHeight)
	{
		py = 1;
	}
		
	return (px+py);
}

//-----------MenuPosition return(obj)--------------//

function MPreturn(obj)
{
	var topobj = eval(obj.substring(0,obj.length-2));
	
	var trobj  = eval("tr"+obj);
	
	var x = topobj.style.pixelLeft + topobj.offsetWidth;
	var y = topobj.style.pixelTop  + trobj.offsetTop;

	obj = eval(obj);
	
	var judgerY =  obj.offsetHeight + y;
	var judgerX =  obj.offsetWidth  + x;
	
	var py = 0;
	var px = 0;
	
	if(judgerY>=document.body.clientHeight)
	{
		py = 1;
	}
	
	if(judgerX>= document.body.clientWidth)
	{
		px = 2;
	} 
			
	return (px+py);
}

//-----------mouseWave()-------------//

function mouseWave()
{
	if(mouseCanSound)
	{
		theBS.src= "images/popmenu/sound.wav";
	}	
}

//----------- menu_alpha_down -------//

function menu_alpha_down(obj,num)
{
		var obj = eval(obj);
		
		if(obj.filters.Alpha.Opacity > 0 )
		{
			obj.filters.Alpha.Opacity += -num;
		}	
		else
		{	
			clearInterval(obj.alphaing);
			obj.filters.Alpha.Opacity = 0;
			obj.alphaing = false;			
			obj.style.visibility = "hidden";
		}	
}


//------------ menu_alpha_up --------//

function menu_alpha_up(obj,num)
{
		var obj = eval(obj);
		
		if(obj.filters.Alpha.Opacity<100)
			obj.filters.Alpha.Opacity += num;
		else
		{	
			clearInterval(obj.alphaing);
			obj.filters.Alpha.Opacity = 100;
			obj.alphaing = false;
		}	
}


//----------- IE ContextMenu -----------------//

function document.oncontextmenu()
{
  var RangeType = document.selection.type;
	if(RangeType != "Text" && event.srcElement.type!='text')
	{
	return false;
	}
	else
	{
	return true;
	}
}


//----------- IE Mouseup ----------------//

function document.onmouseup()
{
  var RangeType = document.selection.type;
	if(event.button==2 && RangeType != "Text" && event.srcElement.type!='text')
	{
	
		HideMenu(0);
		

		var obj = eval(JustMenuID)
		
		
			obj.style.visibility = "hidden";
			
			
			if(obj.alphaing)
			{
				clearInterval(obj.alphaing);
			}
			
			obj.filters.Alpha.Opacity = 0;
			
			var judger = MMPreturn()
			
			if(judger==0)
			{
				with(obj.style)
				{
					pixelLeft = event.clientX + document.body.scrollLeft;
					pixelTop  = event.clientY + document.body.scrollTop;
				}
			}
			if(judger==1)
			{
				with(obj.style)
				{
					pixelLeft = event.clientX + document.body.scrollLeft;
					pixelTop  = event.clientY - obj.offsetHeight + document.body.scrollTop;
					if (event.clientY - obj.offsetHeight<=document.body.scrollTop){ pixelTop=document.body.scrollTop}
				}
			}
			if(judger==2)
			{
				with(obj.style)
				{
					pixelLeft = event.clientX - obj.offsetWidth + document.body.scrollLeft;
					pixelTop  = event.clientY + document.body.scrollTop;
				}
			}
			if(judger==3)
			{
				with(obj.style)
				{
					pixelLeft = event.clientX - obj.offsetWidth + document.body.scrollLeft;
					pixelTop  = event.clientY - obj.offsetHeight + document.body.scrollTop;
					if (event.clientY - obj.offsetHeight<=document.body.scrollTop){ pixelTop=document.body.scrollTop}
				}
			}
			
			mouseWave();
						
			obj.style.visibility = "visible";
			
			obj.alphaing = setInterval("menu_alpha_up("+obj.id+","+alphaStep+")",menuSpeed);

		
		
	}
}

//---------- IE MouseDown --------------//

function document.onmousedown()
{
	if(event.button==1)
	{
		HideMenu();
		
		var obj = eval(JustMenuID)
		
		if(obj.alphaing)
		{
			clearInterval(obj.alphaing);
		}
		
		obj.alphaing = setInterval("menu_alpha_down("+obj.id+","+alphaStep+")",menuSpeed);
		
	}
}
//----->
</script></body>
<%end if%>