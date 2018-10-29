<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%@language=vbscript codepage=936 %>
<%
'option explicit
response.buffer=true	
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="conn.asp"-->
<!--#include file="Conn_User.asp"-->
<!--#include file="config.asp"-->
<!--#include file="ubbcode.asp"-->
<!--#include file="function.asp"-->
<!--#include file="admin_code_article.asp"-->
<!--#include file="../sms_function.asp"-->
<%
dim strChannel,sqlChannel,rsChannel,ChannelUrl,ChannelName
dim ArticleID,ArticleTitle
dim FileName,strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim ClassID,SpecialID,keyword,strField,SpecialName
dim rs,sql,sqlArticle,rsArticle,sqlSearch,rsSearch,rsPic,sqlSpecial,rsSpecial,sqlUser,rsUser
dim tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,SkinID,LayoutID,LayoutFileName,ChildID,tID,tChild
dim tSpecial
dim strPic,AnnounceCount
dim PageTitle,strPath,strPageTitle
dim strClassTree
UserLogined=CheckUserLogined()

BeginTime=Timer
ArticleID=trim(request("ArticleID"))
ClassID=trim(request("ClassID"))
SpecialID=trim(request("SpecialID"))
strField=trim(request("Field"))
keyword=trim(request("keyword"))
'UserLevel=request.Cookies("asp163")("UserLevel")

if ArticleId="" then
	ArticleID=0
else
	ArticleID=Clng(ArticleID)
end if
if ClassID<>"" then
	ClassID=CLng(ClassID)
else
	ClassID=0
end if
if SpecialID="" then
	SpecialID=0
else
	SpecialID=CLng(SpecialID)
end if
if UserLevel="" then
	UserLevel=5000
else
	UserLevel=Cint(UserLevel)
end if
strPath= "&nbsp;您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>"
strPageTitle= SiteTitle
if ShowSiteChannel="Yes" then
	strChannel= "<font color=white>|&nbsp;</font>"
	sqlChannel="select * from Channel order by OrderID"
	set rsChannel=server.CreateObject("adodb.recordset")
	rsChannel.open sqlChannel,conn,1,1
	do while not rsChannel.eof
		if rsChannel("ChannelID")=ChannelID then
			ChannelUrl=rsChannel("LinkUrl")
			ChannelName=rsChannel("ChannelName")
			strChannel=strChannel & "<a href='" & ChannelUrl & "'><font color=white>" & ChannelName & "</a>&nbsp;|&nbsp;</font>"
		else
			strChannel=strChannel & "<a href='" & rsChannel("LinkUrl") & "'><font color=white>" & rsChannel("ChannelName") & "</a>&nbsp;|&nbsp;</font>"
		end if
		rsChannel.movenext
	loop
	rsChannel.close
	set rsChannel=nothing
end if

if ArticleID>0 then
	sql="select * from article where ArticleID=" & ArticleID & ""
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你要找的文章不存在，或者已经被管理员删除！</li>"
	else	
		if rs("Passed")=False then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>你要找的文章尚未被管理员审核！</li>"
		else
			ClassID=rs("ClassID")
			SpecialID=rs("SpecialID")
			SkinID=rs("SkinID")
			LayoutID=rs("LayoutID")
			ArticleTitle=rs("Title")
	
			rs("Hits")=rs("Hits")+1
			rs.update
			if rs("hits")>=HitsOfHot then
				rs("Hot")=True
				rs.update
			end if
		end if
	end if
end if
if FoundErr=True then
	Call WriteErrMsg()
	response.end
end if
if ClassID>0 then
	sql="select C.ClassName,C.RootID,C.ParentID,C.Depth,C.ParentPath,C.Child,C.SkinID,L.LayoutID,L.LayoutFileName,C.BrowsePurview From ArticleClass C"
	sql=sql & " inner join Layout L on C.LayoutID=L.LayoutID where C.ClassID=" & ClassID
	set tClass=conn.execute(sql)
	if tClass.bof and tClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
	else
		if tClass(9)<UserLevel then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>对不起，你没有浏览本栏目的权限！</li>"
			ErrMsg=ErrMsg & "你不是" & CheckLevel(tClass(9)) & "！"
		else
			ClassName=tClass(0)
			RootID=tClass(1)
			ParentID=tClass(2)
			Depth=tClass(3)
			ParentPath=tClass(4)
			Child=tClass(5)
			if ArticleID<=0 then
				SkinID=tClass(6)
				LayoutID=tClass(7)
			end if
			LayoutFileName=tClass(8)

			strPath=strPath & "&nbsp;&gt;&gt;&nbsp;"
			strPageTitle=strPageTitle & " >> "
			if ParentID>0 then
				dim sqlPath,rsPath
				sqlPath="select ArticleClass.ClassID,ArticleClass.ClassName,Layout.LayoutFileName,Layout.LayoutID From ArticleClass"
				sqlPath= sqlPath & " inner join Layout on ArticleClass.LayoutID=Layout.LayoutID where ArticleClass.ClassID in (" & ParentPath & ") order by ArticleClass.Depth"
				set rsPath=server.createobject("adodb.recordset")
				rsPath.open sqlPath,conn,1,1
				do while not rsPath.eof
					strPath=strPath & "<a href='" & rsPath(2) & "?ClassID=" & rsPath(0) & "&LayoutID=" & rsPath(3) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
					strPageTitle=strPageTitle & rsPath(1) & " >> "
					rsPath.movenext
				loop
				rsPath.close
				set rsPath=nothing
			end if
			strPath=strPath & "<a href='" & LayoutFileName & "?ClassID=" & ClassID & "'>" & ClassName & "</a>"
			strPageTitle=strPageTitle & ClassName
		end if	
	end if
end if
if FoundErr=True then
	Call WriteErrMsg()
	response.end
end if
if SpecialID>0 then
	sql="select S.SpecialID,S.SpecialName,S.SkinID,S.LayoutID,L.LayoutFileName,S.BrowsePurview,S.SpecialAbbreviation from Special S inner join Layout L on L.LayoutID=S.LayoutID where S.SpecialID=" & SpecialID
	set tSpecial=conn.execute(sql)
	if tSpecial.bof and tSpecial.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
	else
		if tSpecial(5)<UserLevel then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>对不起，你没有浏览本课程的权限！</li>"
			ErrMsg=ErrMsg & "你不是" & CheckLevel(tSpecial(5)) & "！"
		else
			SpecialName=tSpecial("SpecialAbbreviation")
			if ArticleID<=0 then
				SkinID=tSpecial(2)
				LayoutID=tSpecial(3)
			end if
			LayoutFilename=tSpecial(4)
			
			strPath=strPath & "&nbsp;&gt;&gt;&nbsp;<font color=blue>[课程]</font><a href='" & tSpecial(4) & "?SpecialID=" & tSpecial(0) & "'>" & SpecialName & "</a>"
			strPageTitle=strPageTitle & " >> [课程]" & SpecialName
		end if
	end if
end if
if FoundErr=True then
	Call WriteErrMsg()
	response.end
end if

if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if


'=================================================
'过程名：ShowRootClass
'作  用：显示一级栏目（无特殊效果）
'参  数：无
'=================================================
sub ShowRootClass()
	dim sqlRoot,rsRoot
	sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,C.LinkUrl From ArticleClass C"
	sqlRoot= sqlRoot & " inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=0 and C.ShowOnTop=True order by C.RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1
	if rsRoot.bof and rsRoot.eof then 
		response.Write("还没有任何栏目，请首先添加栏目。")
	else
		if ClassID>0 then
			response.write "|<a href='" & ChannelUrl & "'>&nbsp;"&ChannelName&"首页&nbsp;</a>|"
		else
			response.write "|<a href='" & ChannelUrl & "'><font color=red>&nbsp;"&ChannelName&"首页&nbsp;</font></a>|"
		end if
		do while not rsRoot.eof
			if rsRoot(4)<>"" then
				response.write "<a href='" & rsRoot(4) & "'>&nbsp;" & rsRoot(1) & "&nbsp;</a>|"
			else
				if rsRoot(2)=RootID then
					response.Write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'><font color=red>&nbsp;" & rsRoot(1) & "</font>&nbsp;</a>|"
				else
					response.Write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>&nbsp;" & rsRoot(1) & "&nbsp;</a>|"
				end if
			end if
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
	if ShowMyStyle="Yes" then
		response.write "<a href='#' onMouseOver='ShowMenu(menu_skin,100)'>&nbsp;自选风格&nbsp;</a>|"
	end if
end sub

dim pNum,pNum2
pNum=1
pNum2=0
'=================================================
'过程名：ShowRootClass_Menu
'作  用：显示一级栏目（下拉菜单效果）
'参  数：无
'=================================================
sub ShowRootClass_Menu()
	response.write "<script type='text/javascript' language='JavaScript1.2'>" & vbcrlf & "<!--" & vbcrlf
	response.write "stm_bm(['uueoehr',400,'','images/blank.gif',0,'','',0,0,0,0,0,1,0,0]);" & vbcrlf
	response.write "stm_bp('p0',[0,4,0,0,2,2,0,0,100,'',-2,'',-2,90,0,0,'#000000','transparent','',3,0,0,'#000000']);" & vbcrlf
	response.write "stm_ai('p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbcrlf
	response.write "stm_aix('p0i1','p0i0',[0,'"&ChannelName&"首页','','',-1,-1,0,'" & ChannelUrl & "','_self','" & ChannelUrl & "','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbcrlf
	response.write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbcrlf

	dim sqlRoot,rsRoot,j
	sqlRoot="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl,C.Child,C.Readme From ArticleClass C"
	sqlRoot= sqlRoot & " inner join Layout L on C.LayoutID=L.LayoutID where C.Depth=0 and C.ShowOnTop=True order by C.RootID"
	Set rsRoot= Server.CreateObject("ADODB.Recordset")
	rsRoot.open sqlRoot,conn,1,1
	if not(rsRoot.bof and rsRoot.eof) then 
		j=3
		do while not rsRoot.eof
			if rsRoot(5)<>"" then
				response.write "stm_aix('p0i"&j&"','p0i0',[0,'" & rsRoot(1) & "','','',-1,-1,0,'" & rsRoot(5) & "','_blank','" & rsRoot(5) & "','" & rsRoot(7) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbcrlf
			else
				response.write "stm_aix('p0i"&j&"','p0i0',[0,'" & rsRoot(1) & "','','',-1,-1,0,'" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "','_self','" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "','" & rsRoot(7) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体','9pt 宋体']);" & vbcrlf
			end if
			if rsRoot(6)>0 then
				call GetClassMenu(rsRoot(0),0)
			end if
			j=j+1
			response.write "stm_aix('p0i2','p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',1,'','',3,3,0,0,'#fffff7','#000000','#000000','#000000','9pt 宋体','9pt 宋体',0,0]);" & vbcrlf
			j=j+1
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
	response.write "stm_em();" & vbcrlf
	response.write "//-->" & vbcrlf & "</script>" & vbcrlf
end sub

sub GetClassMenu(ID,ShowType)
	dim sqlClass,rsClass,k
	
	if pNum=1 then
		response.write "stm_bp('p" & pNum & "',[1,4,0,0,2,3,6,7,100,'progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)',-2,'',-2,67,2,3,'#999999','#ffffff','',3,1,1,'#aca899']);" & vbcrlf
	else
		if ShowType=0 then
			response.write "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,4,0,0,2,3,6]);" & vbcrlf
		else
			response.write "stm_bpx('p" & pNum & "','p" & pNum2 & "',[1,2,-2,-3,2,3,0]);" & vbcrlf
		end if
	end if
	
	k=0
	sqlClass="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl,C.Child,C.Readme From ArticleClass C"
	sqlClass= sqlClass & " inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ID & " order by C.OrderID asc"
	Set rsClass= Server.CreateObject("ADODB.Recordset")
	rsClass.open sqlClass,conn,1,1
	do while not rsClass.eof
		if rsClass(5)<>"" then
			if rsClass(6)>0 then
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'" & rsClass(1) & "','','',-1,-1,0,'" & rsClass(5) & "','_blank','" & rsClass(5) & "','" & rsClass(7) & "','','',6,0,0,'images/arrow_r.gif','images/arrow_w.gif',7,7,0,0,1,'#ffffff',0,'#cccccc',0,'','',3,3,0,0,'#fffff7','#000000','#000000','#ffffff','9pt 宋体']);" & vbcrlf
				pNum=pNum+1
				pNum2=pNum2+1
				call GetClassMenu(rsClass(0),1)
			else
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'" & rsClass(1) & "','','',-1,-1,0,'" & rsClass(5) & "','_blank','" & rsClass(5) & "','" & rsClass(7) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',0,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体']);" & vbcrlf
			end if
		else
			if rsClass(6)>0 then
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'" & rsClass(1) & "','','',-1,-1,0,'" & rsClass(3) & "?ClassID=" & rsClass(0) & "','_self','" & rsClass(3) & "?ClassID=" & rsClass(0) & "','" & rsClass(7) & "','','',6,0,0,'images/arrow_r.gif','images/arrow_w.gif',7,7,0,0,1,'#ffffff',0,'#cccccc',0,'','',3,3,0,0,'#fffff7','#000000','#000000','#ffffff','9pt 宋体']);" & vbcrlf
				pNum=pNum+1
				pNum2=pNum2+1
				call GetClassMenu(rsClass(0),1)
			else
				response.write "stm_aix('p"&pNum&"i"&k&"','p"&pNum2&"i0',[0,'" & rsClass(1) & "','','',-1,-1,0,'" & rsClass(3) & "?ClassID=" & rsClass(0) & "','_self','" & rsClass(3) & "?ClassID=" & rsClass(0) & "','" & rsClass(7) & "','','',0,0,0,'','',0,0,0,0,1,'#f1f2ee',1,'#cccccc',0,'','',3,3,0,0,'#fffff7','#ff0000','#000000','#cc0000','9pt 宋体']);" & vbcrlf
			end if
		end if
		k=k+1
		rsClass.movenext
	loop
	rsClass.close
	set rsClass=nothing
	response.write "stm_ep();" & vbcrlf
end sub

'=================================================
'过程名：ShowJumpClass
'作  用：显示“跳转栏目到…”下拉列表框
'参  数：无
'=================================================
sub ShowJumpClass()
	response.write "<div z-index:1><select name='ClassID' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">"
    response.write "<option value='' selected>跳转栏目至…</option>"
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	sqlClass="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl From ArticleClass C"
	sqlClass= sqlClass & " inner join Layout L on C.LayoutID=L.LayoutID order by C.RootID,C.OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
	if rsClass.bof and rsClass.bof then
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass(2)
			if rsClass(4)>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
			if rsClass(5)="" then
				strTemp="<option value='" & rsClass(3) & "?ClassID=" & rsClass(0) & "'>"
			else
				strTemp="<option value='" & rsClass(5) & "'>"
			end if
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass(4)>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass(1)
			if rsClass(5)<>"" then
				strTemp=strTemp & "(外)"
			end if
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
	response.write "</select></div>"
end sub

'=================================================
'过程名：ShowClass_Tree
'作  用：显示所有栏目（树形目录效果）
'参  数：无
'=================================================
sub ShowClass_Tree()
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim rsClass,sqlClass,tmpDepth,i
	sqlClass="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl,C.Child From ArticleClass C"
	sqlClass= sqlClass & " inner join Layout L on C.LayoutID=L.LayoutID order by C.RootID,C.OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
	if rsClass.bof and rsClass.bof then
		strClassTree="没有任何栏目"
	else
		strClassTree=""
		do while not rsClass.eof
			tmpDepth=rsClass(2)
			if rsClass(4)>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
			if tmpDepth>0 then
				for i=1 to tmpDepth
					if i=tmpDepth then
						if rsClass(4)>0 then
							strClassTree=strClassTree & "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
						else
							strClassTree=strClassTree & "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
						end if
					else
						if arrShowLine(i)=True then
							strClassTree=strClassTree & "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
						else
							strClassTree=strClassTree & "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
						end if
					end if
				next
			end if
			if rsClass(6)>0 then 
				strClassTree=strClassTree & "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
			else 
				strClassTree=strClassTree & "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
			end if 
			if rsClass(5)="" then
				strClassTree=strClassTree & "<a href='" & rsClass(3) & "?ClassID=" & rsClass(0) & "'>"
			else
				strClassTree=strClassTree & "<a href='" & rsClass(5) & "' target='_blank'>"
			end if
			if rsClass(2)=0 then 
				strClassTree=strClassTree & "<b>"  & rsClass(1) & "</b>"
			else
				strClassTree=strClassTree & rsClass(1)
			end if 
			'if rsClass(5)<>"" then
			'	strClassTree=strClassTree & "(外)"
			'end if
			strClassTree=strClassTree & "</a>"
			if rsClass(6)>0 then 
				strClassTree=strClassTree & "（" & rsClass(6) & "）" 
			end if 
			strClassTree=strClassTree & "<br>"
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
	response.write strClassTree
end sub

'=================================================
'过程名：ShowChildClass
'作  用：显示当前栏目的下一级子栏目
'参  数：ShowType--------显示方式，1为竖向列表，2为横向列表
'=================================================
sub ShowChildClass(ShowType)
	dim sqlChild,rsChild,i
	sqlChild="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl,C.Child From ArticleClass C"
	sqlChild= sqlChild & " inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ClassID & " order by C.OrderID"
	Set rsChild= Server.CreateObject("ADODB.Recordset")
	rsChild.open sqlChild,conn,1,1
	if rsChild.bof and rsChild.eof then
		response.write "没有任何子栏目"
	else
		if ShowType=1 then
			do while not rsChild.eof
				if rsChild(5)<>"" then
					response.write "<li><a href='" & rsChild(5) & "'>" & rsChild(1) & "</a></li>"
				else
					response.Write "<li><a href='" & rsChild(3) & "?ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a></li>"
				end if
				if rsChild(6)>0 then
					response.write "(" & rsChild(6) & ")"
				end if
				response.write "<br>"		
				rsChild.movenext
			loop
		else
			i=0
			do while not rsChild.eof
				if rsChild(5)<>"" then
					response.write "&nbsp;&nbsp;<a href='" & rsChild(5) & "'>" & rsChild(1) & "</a>"
				else
					response.Write "&nbsp;&nbsp;<a href='" & rsChild(3) & "?ClassID=" & rsChild(0) & "'>" & rsChild(1) & "</a>"
				end if
				if rsChild(6)>0 then
					response.write "(" & rsChild(6) & ")"
				end if		
				rsChild.movenext
				i=i+1
				if i mod 5=0 then
					response.write "<br>"
				end if
			loop
		end if
	end if
	rsChild.close
	set rsChild=nothing
end sub

'=================================================
'过程名：ShowClassNavigation
'作  用：显示栏目导航
'参  数：无
'=================================================
sub ShowClassNavigation()
	dim rsNavigation,sqlNavigation,strNavigation,PrevRootID,i
	sqlNavigation="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.RootID,C.LinkUrl,C.Child,C.Readme From ArticleClass C"
	sqlNavigation= sqlNavigation & " inner join Layout L on C.LayoutID=L.LayoutID where C.Depth<=1 order by C.RootID,C.OrderID"
	Set rsNavigation= Server.CreateObject("ADODB.Recordset")
	rsNavigation.open sqlNavigation,conn,1,1
	if rsNavigation.bof and rsNavigation.eof then
		response.write "没有任何栏目"
	else
		strNavigation="<table border='0' cellpadding='0' cellspacing='2'><tr><td valign='top' nowrap>【<a href='" & rsNavigation(3) & "?ClassID=" & rsNavigation(0) & "' title='" & rsNavigation(7) & "'>" & rsNavigation(1) & "</a>】</td><td>"
		PrevRootID=rsNavigation(4)
		rsNavigation.movenext
		i=1
		do while not rsNavigation.eof
			if PrevRootID=rsNavigation(4) then
				if i mod 9=0 then
					strNavigation=strNavigation & "<br>"
				end if
				strNavigation=strNavigation & "<a href='" & rsNavigation(3) & "?ClassID=" & rsNavigation(0) & "' title='" & rsNavigation(7) & "'>" & rsNavigation(1) & "</a>&nbsp;&nbsp;"
				i=i+1
			else
				strNavigation=strNavigation & "</td></tr><tr><td valign='top' nowrap>【<a href='" & rsNavigation(3) & "?ClassID=" & rsNavigation(0) & "' title='" & rsNavigation(7) & "'>" & rsNavigation(1) & "</a>】</td><td>"
				i=1
			end if
			PrevRootID=rsNavigation(4)
			rsNavigation.movenext
		loop
		strNavigation=strNavigation & "</td></tr></table>"
		response.write strNavigation
	end if
	rsNavigation.close
	set rsNavigation=nothing
end sub

'=================================================
'过程名：ShowSpecial
'作  用：以竖向列表方式显示课程名称
'参  数：SpecialNum  ------最多显示多少个课程名称
'=================================================
sub ShowSpecial(SpecialNum)
	dim i
	i=1
	if SpecialNum<=0 or SpecialNum>100 then
		SpecialNum=10
	end if
	sqlSpecial="select S.SpecialID,S.SpecialName,L.LayoutFileName from Special S inner join Layout L on L.LayoutID=S.LayoutID where S.BrowsePurview>=" & UserLevel & " order by S.OrderID"
	Set rsSpecial= Server.CreateObject("ADODB.Recordset")
	rsSpecial.open sqlSpecial,conn,1,1
	totalPut=rsSpecial.recordcount
	if rsSpecial.bof and rsSpecial.eof then 
		response.Write "&nbsp;没有任何课程栏目"
	else
		rsSpecial.movefirst
		do while not rsSpecial.eof
			response.Write(" - <a href='" & rsSpecial(2) & "?SpecialID=" & rsSpecial(0) & "'>" & rsSpecial(1) & "</a></li><br>")
			rsSpecial.movenext
			i=i+1
			if i>SpecialNum then exit do
		loop
	end if
	if not rsSpecial.eof then
		response.write "<p align='right'><a href='Article_SpecialList.asp'>更多课程</a></p>"
	end if
end sub

'=================================================
'过程名：ShowAllSpecial
'作  用：分页显示所有课程
'参  数：无
'=================================================
sub ShowAllSpecial()
	sqlSpecial="select S.SpecialID,S.SpecialName,L.LayoutFileName from Special S inner join Layout L on L.LayoutID=S.LayoutID where S.BrowsePurview>=" & UserLevel & " order by S.OrderID"
	Set rsSpecial= Server.CreateObject("ADODB.Recordset")
	rsSpecial.open sqlSpecial,conn,1,1
	totalPut=rsSpecial.recordcount
	if rsSpecial.bof and rsSpecial.eof then 
		response.Write "&nbsp;没有任何课程栏目"
	else
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage>totalput then
			if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if
		if currentPage=1 then
			call SpecialContent()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsSpecial.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsSpecial.bookmark
            	call SpecialContent()
        	else
	        	currentPage=1
           		call SpecialContent()
	    	end if
		end if
	end if
end sub

sub SpecialContent()
	dim i
	i=1
	do while not rsSpecial.eof
		response.Write("<li><a href='" & rsSpecial(2) & "?SpecialID=" & rsSpecial(0) & "'>" & rsSpecial(1) & "</a></li><br>")
		rsSpecial.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub

'=================================================
'过程名：ShowSiteCount
'作  用：显示站点统计信息
'参  数：无
'=================================================
sub ShowSiteCount()
	dim sqlCount,rsCount
	Set rsCount= Server.CreateObject("ADODB.Recordset")
	sqlCount="select count(ArticleID) from Article where Deleted=False"
	rsCount.open sqlCount,conn,1,1
	response.write "文章总数：" & rsCount(0) & "篇<br>"
	rsCount.close

	sqlCount="select count(ArticleID) from Article where Passed=False and Deleted=False"
	rsCount.open sqlCount,conn,1,1
	response.write "待审文章：" & rsCount(0) & "篇<br>"
	rsCount.close
	
	sqlCount="select count(CommentID) from ArticleComment"
	rsCount.open sqlCount,conn,1,1
	response.write "评论总数：" & rsCount(0) & "条<br>"
	rsCount.close
	
	sqlCount="select count(SpecialID) from Special"
	rsCount.open sqlCount,conn,1,1
	response.write "课程总数：" & rsCount(0) & "个<br>"
	rsCount.close

	sqlCount="select sum(Hits) from article"
	rsCount.open sqlCount,conn,1,1
	response.write "文章阅读：" & rsCount(0) & "人次<br>"
	rsCount.close
	
	set rsCount=nothing
end sub

'=================================================
'过程名：ShowArticle
'作  用：分页显示文章标题等信息
'参  数：TitleLen  ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowArticle(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if

	'sqlArticle的值是否保留？
		'sqlArticle= "select  A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"

	sqlArticle=sqlArticle & "select  A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
	sqlArticle=sqlArticle & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sqlArticle=sqlArticle & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True "
	if SpecialID>0 then
		sqlArticle=sqlArticle & " and A.SpecialID=" & SpecialID
	end if
	if ClassId>0 then
		sqlArticle=sqlArticle &  " and A.ClassID=" & ClassID
	end if
	sqlArticle=sqlArticle & " order by A.OnTop,A.ArticleID desc"

	Set rsArticle= Server.CreateObject("ADODB.Recordset")
	rsArticle.open sqlArticle,conn,1,1
	if rsArticle.bof and  rsArticle.eof then
		totalput=0
		response.Write("<br><li>没有任何符合条件的文章</li>")
	else
		totalput=rsArticle.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage>totalput then
			if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if

		if currentPage=1 then
			call ArticleContent(TitleLen,True,True,True,2,True,True)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsArticle.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsArticle.bookmark
            	call ArticleContent(TitleLen,True,True,True,2,True,True)
        	else
	        	currentPage=1
           		call ArticleContent(TitleLen,True,True,True,2,True,True)
	    	end if
		end if
	end if
	rsArticle.close
	set rsArticle=nothing
end sub



'=================================================
'过程名：ShowArticle_Index
'作  用：分页显示文章标题等信息
'参  数：TitleLen_Index  ----标题最多字符数，一个汉字=两个英文字符
'		ClassID_Index--------文章所属栏目编号
'		SpecialID_Index------文章所属课程编号
' 		MaxPerPage_Index-----显示文章数
'=================================================
dim sqlArticle_Index,rsArticle_Index,totalput_Index
sub ShowArticle_Index(TitleLen_Index,ClassID_Index,SpecialID_Index,MaxPerPage_Index)
	if MaxperPage_Index<0 or MaxPerPage_Index>50 then
		MaxPerPage_Index=10
	end if
	if TitleLen_Index<0 or TitleLen_Index>200 then
		TitleLen_Index=50
	end if

	'sqlArticle是连接字串还是重新赋值？
	sqlArticle_Index= "select  A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
	sqlArticle_Index=sqlArticle_Index & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sqlArticle_Index=sqlArticle_Index & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True "
	if SpecialID_Index>0 then
		sqlArticle_Index=sqlArticle_Index & " and A.SpecialID=" & SpecialID_Index
	end if
	if ClassID_Index>0 then
		sqlArticle_Index=sqlArticle_Index &  " and A.ClassID=" & ClassID_Index
	end if
	sqlArticle_Index=sqlArticle_Index & " order by A.OnTop,A.ArticleID desc"

	Set rsArticle_Index= Server.CreateObject("ADODB.Recordset")
	rsArticle_Index.open sqlArticle_Index,conn,1,1
	if rsArticle_Index.bof and  rsArticle_Index.eof then
		totalput_Index=0
		response.Write("<br><li>没有任何符合条件的文章</li>")
	else
		totalput_Index=rsArticle_Index.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxperPage_Index>totalput_Index then
			if (totalput_Index mod MaxperPage_Index)=0 then
				currentpage= totalput_Index \ MaxperPage_Index
			else
				currentpage= totalput_Index \ MaxperPage_Index + 1
			end if
		end if

		if currentPage=1 then
			
			call ArticleContent_Index(TitleLen_Index,True,True,True,2,True,True,MaxPerPage_Index)
		else
			if (currentPage-1)*MaxperPage_Index<totalput_Index then
         	   	rsArticle_Index.move  (currentPage-1)*MaxperPage_Index
         		dim bookmark
           		bookmark=rsArticle_Index.bookmark
            	call ArticleContent_Index(TitleLen_Index,True,True,True,2,True,True,MaxPerPage_Index)
        	else
	        	currentPage=1
           		call ArticleContent_Index(TitleLen_Index,True,True,True,2,True,True,MaxPerPage_Index)
	    	end if
		end if
	end if
	rsArticle_Index.close
	set rsArticle_Index=nothing
end sub
'end sub ShowArticle_Index

'=================================================
'过程名：ArticleContent_Index
'作  用：显示文章属性、标题、作者、更新日期、点击数等信息
'参  数：intTitleLen  ----标题最多字符数，一个汉字=两个英文字符
'        ShowProperty ----是否显示文章属性（固顶/推荐/普通），True为显示，False为不显示
'        ShowIncludePic ---是否显示“[图文]”字样，True为显示，False为不显示
'        ShowAuthor -------是否显示文章作者，True为显示，False为不显示
'        ShowDateType -----显示更新日期的样式，0为不显示，1为显示年月日，2为只显示月日。
'        ShowHits ---------是否显示文章点击数，True为显示，False为不显示
'        ShowHot ----------是否显示热门文章标志，True为显示，False为不显示
'     		MaxPerPage-----显示文章数
'=================================================
sub ArticleContent_Index(intTitleLen,ShowProperty,ShowIncludePic,ShowAuthor,ShowDateType,ShowHits,ShowHot,MaxPerPage_Index)
   	dim i,strTemp,TitleStr,Author,AuthorName,AuthorEmail
    i=0
	response.write "<table width='100%' border='0'><tr>"
	'response.write "<table width='100%' border='0'><tr align='top'><td align='top' width='40%'>&nbsp;&nbsp;<font color='#009900'>文章标题</font></td><td width='20%'><font color='#009900'>日期</font></td><td width='20%'><font color='#009900'>阅读</font></td><td width='20%'><font color='#009900'>作者</font></td></tr><tr>"
	do while not rsArticle_Index.eof
	strTemp=""
		if rsArticle_Index("IncludePic")=true then strTemp = strTemp & ""
		strTemp = strTemp & "<td>&nbsp;&nbsp;<a href='" & rsArticle_Index("LayoutFileName") & "?ArticleID=" & rsArticle_Index("articleid") & "' target='_blank'>"
		TitleStr=gotTopic(rsArticle_Index("title"),intTitleLen)
		if rsArticle_Index("TitleFontType")=1 then
			TitleStr="<b>" & TitleStr & "</b>"
		elseif rsArticle_Index("TitleFontType")=2 then
			TitleStr="<em>" & TitleStr & "</em>"
		elseif rsArticle_Index("TitleFontType")=3 then
			TitleStr="<b><em>" & TitleStr & "</em></b>"
		end if
		if rsArticle_Index("TitleFontColor")<>"" then
			TitleStr="<font color='" & rsArticle_Index("TitleFontColor") & "'>" & TitleStr & "</font>"
		end if
		strTemp=strTemp & TitleStr & "</a>"
		if rsArticle_Index("Elite")=true then
				strTemp = strTemp & "<img src='images/tuijian.gif' alt='推荐文章'>"
		end if
		strTemp=strTemp & "</td><td align='top'>"
		if CDate(FormatDateTime(rsArticle_Index("UpdateTime"),2))=date() then
			strTemp = strTemp & "<font color=red>"
		else
			strTemp= strTemp & "<font color=#999999>"
		end if
		strTemp= strTemp & FormatDateTime(rsArticle_Index("UpdateTime"),1) & "</font></td><td align='top'>" 
 
		strTemp= strTemp & rsArticle_Index("Hits") & "</td><td align='top'>" 
		strTemp= strTemp & rsArticle_Index("Author") &"</td></tr>"
		rsArticle_Index.movenext
		
		response.write strTemp
		i=i+1
		
		if i>=MaxperPage_Index then exit do
		'response.Write("123456678999")
		loop 
     response.write"</table>"
	 end sub
		
'end sub ArticleContent_Index
'=================================================
'过程名：ArticleContent
'作  用：显示文章属性、标题、作者、更新日期、点击数等信息
'参  数：intTitleLen  ----标题最多字符数，一个汉字=两个英文字符
'        ShowProperty ----是否显示文章属性（固顶/推荐/普通），True为显示，False为不显示
'        ShowIncludePic ---是否显示“[图文]”字样，True为显示，False为不显示
'        ShowAuthor -------是否显示文章作者，True为显示，False为不显示
'        ShowDateType -----显示更新日期的样式，0为不显示，1为显示年月日，2为只显示月日。
'        ShowHits ---------是否显示文章点击数，True为显示，False为不显示
'        ShowHot ----------是否显示热门文章标志，True为显示，False为不显示
'=================================================
sub ArticleContent(intTitleLen,ShowProperty,ShowIncludePic,ShowAuthor,ShowDateType,ShowHits,ShowHot)
   	dim i,strTemp,TitleStr,Author,AuthorName,AuthorEmail
    i=0
		response.write "<table width='100%' border='0'><tr align='top'><td align='top' width='40%'>&nbsp;&nbsp;<font color='#009900'>文章标题</font></td><td width='20%'><font color='#009900'>日期</font></td><td width='20%'><font color='#009900'>阅读</font></td><td width='20%'><font color='#009900'>作者</font></td></tr><tr>"
	do while not rsArticle.eof
	strTemp=""
		if rsArticle("IncludePic")=true then strTemp = strTemp & ""
		strTemp = strTemp & "<td>&nbsp;&nbsp;<a href='" & rsArticle("LayoutFileName") & "?ArticleID=" & rsArticle("articleid") & "' target='_blank'>"
		TitleStr=gotTopic(rsArticle("title"),intTitleLen)
		if rsArticle("TitleFontType")=1 then
			TitleStr="<b>" & TitleStr & "</b>"
		elseif rsArticle("TitleFontType")=2 then
			TitleStr="<em>" & TitleStr & "</em>"
		elseif rsArticle("TitleFontType")=3 then
			TitleStr="<b><em>" & TitleStr & "</em></b>"
		end if
		if rsArticle("TitleFontColor")<>"" then
			TitleStr="<font color='" & rsArticle("TitleFontColor") & "'>" & TitleStr & "</font>"
		end if
		strTemp=strTemp & TitleStr & "</a>"
		if rsArticle("Elite")=true then
				strTemp = strTemp & "<img src='images/tuijian.gif' alt='推荐文章'>"
		end if
		strTemp=strTemp & "</td><td align='top'>"
		if CDate(FormatDateTime(rsArticle("UpdateTime"),2))=date() then
			strTemp = strTemp & "<font color=red>"
		else
			strTemp= strTemp & "<font color=#999999>"
		end if
		strTemp= strTemp & FormatDateTime(rsArticle("UpdateTime"),1) & "</font></td><td align='top'>" 
 
		strTemp= strTemp & rsArticle("Hits") & "</td><td align='top'>" 
		strTemp= strTemp & rsArticle("Author") &"</td></tr>"
		rsArticle.movenext
		
		response.write strTemp
		i=i+1
		if i>=MaxPerPage then exit do	
         loop 
     response.write"</table>"
	 end sub
		
'=================================================
'过程名：ShowUserArticle
'作  用：分页显示用户文章标题等信息
'参  数：TitleLen  ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowUserArticle(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if

	sqlArticle=sqlArticle & "select  A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
	sqlArticle=sqlArticle & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sqlArticle=sqlArticle & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and Editor='" & UserName & "'"
	if SpecialID>0 then
		sqlArticle=sqlArticle & " and A.SpecialID=" & SpecialID
	end if
	if ClassId>0 then
		sqlArticle=sqlArticle &  " and A.ClassID=" & ClassID
	end if
	sqlArticle=sqlArticle & " order by A.OnTop,A.ArticleID desc"

	Set rsArticle= Server.CreateObject("ADODB.Recordset")
	rsArticle.open sqlArticle,conn,1,1
	if rsArticle.bof and  rsArticle.eof then
		totalput=0
		response.Write("<br><li>没有任何符合条件的文章</li>")
	else
		totalput=rsArticle.recordcount
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage>totalput then
			if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if
		if currentPage=1 then
			call ArticleContent(TitleLen,True,True,True,2,True,True)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsArticle.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsArticle.bookmark
            	call ArticleContent(TitleLen,True,True,True,2,True,True)
        	else
	        	currentPage=1
           		call ArticleContent(TitleLen,True,True,True,2,True,True)
	    	end if
		end if
	end if
	rsArticle.close
	set rsArticle=nothing
end sub

'=================================================
'过程名：ShowSearchResult
'作  用：分页显示搜索结果
'参  数：无
'=================================================
sub ShowSearchResult()
	dim arrClassID,trs
	sqlSearch=sqlSearch & "select A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,A.Content,"
	sqlSearch=sqlSearch & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sqlSearch=sqlSearch & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True"
	if ClassID>0 then
		if Child>0 then
			arrClassID=ClassID
			if ParentID>0 then
				set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & ClassID & " or ParentPath like '%" & ParentPath & "," & ClassID & ",%' and Child=0 and LinkUrl='' and BrowsePurview>=" & UserLevel)
			else
				set trs=conn.execute("select ClassID from ArticleClass where RootID=" & RootID & " and Child=0 and LinkUrl='' and BrowsePurview>=" & UserLevel)
			end if
			do while not trs.eof
				arrClassID=arrClassID & "," & trs(0)
				trs.movenext
			loop
			set trs=nothing			
			sqlSearch=sqlSearch & " and A.ClassID in (" & arrClassID & ")"
		else
			sqlSearch=sqlSearch & " and A.ClassID=" & ClassID
		end if
	end if
	if keyword<>"" then
		select case strField
			case "Title"
				sqlSearch=sqlSearch & " and A.Title like '%" & keyword & "%' "
			case "Content"
				sqlSearch=sqlSearch & " and A.Content like '%" & keyword & "%' "
			case "Author"
				sqlSearch=sqlSearch & " and A.Author like '%" & keyword & "%' "
			case "Editor"
				sqlSearch=sqlSearch & " and A.Editor like '%" & keyword & "%' "
			case else
				sqlSearch=sqlSearch & " and A.Title like '%" & keyword & "%' "
		end select
	end if
	sqlSearch=sqlSearch & " order by A.Articleid desc"
	Set rsSearch= Server.CreateObject("ADODB.Recordset")
	rsSearch.open sqlSearch,conn,1,1
 	if rsSearch.eof and rsSearch.bof then 
		totalput=0
		response.write "<p align='center'><br><br>没有或没有找到任何文章</p>" 
   	else 
   		totalput=rsSearch.recordcount		
 		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage>totalput then
			if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if
  		if currentPage=1 then 
       		call SearchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<totalPut then 
       			rsSearch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSearch.bookmark 
       			call SearchResultContent()
      		else 
        		currentPage=1 
       			call SearchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSearch.close 
   	set rsSearch=nothing   
end sub

sub SearchResultContent()
   	dim i,strTemp,content
	i=1
	do while not rsSearch.eof
		strTemp=""
		strTemp=strTemp & cstr(MaxPerPage*(CurrentPage-1)+i) & ".<a href='" & rsSearch("LayoutFileName") & "?ArticleID=" & rsSearch("articleid") & "'>"
		if strField="Title" then
			strTemp=strTemp & "<b>" & replace(rsSearch("title"),""&keyword&"","<font color=red>"&keyword&"</font>") & "</b></font></a>"
		else
			strTemp=strTemp & "<b>" & rsSearch("title") & "</b>"
		end if
		if strField="Author" then
			strTemp=strTemp & "&nbsp;[" & replace(rsSearch("Author"),""&keyword&"","<font color=red>"&keyword&"</font>") & "]"
		else
			strTemp=strTemp & "&nbsp;[" & rsSearch("Author") & "]"
		end if
		strTemp=strTemp & "[" & FormatDateTime(rsSearch("UpdateTime"),1) & "][" & rsSearch("Hits") & "]"
		content=left(replace(replace(nohtml(rsSearch("content")), ">", "&gt;"), "<", "&lt;"),200)
		if strField="Content" then
			strTemp=strTemp & "<div style='padding:10px 20px'>" & replace(content,""&keyword&"","<font color=red>"&keyword&"</font>") & "……</div>"
		else
			strTemp=strTemp & "<div style='padding:10px 20px'>" & content & "……</div>"
		end if
		strTemp=strTemp & "</a>"
		response.write strTemp
		i=i+1
		if i>MaxPerPage then exit do
		rsSearch.movenext
	loop
end sub 

'=================================================
'过程名：ShowNewArticle
'作  用：显示最新文章
'参  数：ArticleNum  ----最多显示多少篇文章
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowNewArticle(ArticleNum,TitleLen)
	dim sqlNew,rsNew
	if ArticleNum>0 and ArticleNum<=100 then
		sqlNew="select top " & ArticleNum
	else
		sqlNew="select top 10 "
	end if
	sqlNew=sqlNew & " A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName from article A inner join Layout L on L.LayoutID=A.LayoutID where A.Deleted=False and A.Passed=True order by A.articleid desc"
	Set rsNew= Server.CreateObject("ADODB.Recordset")
	rsNew.open sqlNew,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsNew.bof and rsNew.eof then 
		response.write "<li>没有文章</li>" 
	else 
		do while not rsNew.eof   
			response.Write "<li><a href='" & rsNew("LayoutFileName") & "?ArticleID=" & rsNew("articleid") &"' title='文章标题：" & rsNew("Title") & vbcrlf & "作    者：" & rsNew("Author") & vbcrlf & "更新时间：" & rsNew("UpdateTime") & vbcrlf & "点击次数：" & rsNew("Hits") & "' target='_blank'>" & gotTopic(rsNew("title"),TitleLen) & "</a>[<font color=red>" & formatdatetime(rsNew("UpdateTime"),2) & "</font>]</li><br>"
        	rsNew.movenext     
		loop
	end if  
	rsNew.close
	set rsNew=nothing
end sub

'=================================================
'过程名：ShowHot
'作  用：显示热门文章
'参  数：ArticleNum  ----最多显示多少篇文章
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowHot(ArticleNum,TitleLen)
	dim sqlHot,rsHot
	if ArticleNum>0 and ArticleNum<=100 then
		sqlHot="select top " & ArticleNum
	else
		sqlHot="select top 10 "
	end if
	sqlHot=sqlHot & " A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName from article A inner join Layout L on L.LayoutID=A.LayoutID where A.Deleted=False and A.Passed=True And A.Hits>=" & HitsOfHot & " order by A.ArticleID desc"
	Set rsHot= Server.CreateObject("ADODB.Recordset")
	rsHot.open sqlHot,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsHot.bof and rsHot.eof then 
		response.write "<li>无热门文章</li>" 
	else 
		do while not rsHot.eof   
			response.Write "&nbsp;&nbsp;&nbsp; - <a href='" & rsHot("LayoutFileName") & "?ArticleID=" & rsHot("articleid") &"' title='文章标题：" & rsHot("Title") & vbcrlf & "作    者：" & rsHot("Author") & vbcrlf & "更新时间：" & rsHot("UpdateTime") & vbcrlf & "点击次数：" & rsHot("Hits") & "' target='_blank'>" & gotTopic(rsHot("title"),TitleLen) & "</a>&nbsp;[<font color=red>" & rsHot("hits") & "</font>]<br>"
        	rsHot.movenext     
		loop
	end if  
	rsHot.close
	set rsHot=nothing
end sub

'=================================================
'过程名：ShowElite
'作  用：显示推荐文章
'参  数：ArticleNum  ----最多显示多少篇文章
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowElite(ArticleNum,TitleLen)
	dim sqlElite,rsElite
	if ArticleNum>0 and ArticleNum<=100 then
		sqlElite="select top " & ArticleNum
	else
		sqlElite="select top 10 "
	end if
	sqlElite=sqlElite & " A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName from article A inner join Layout L on L.LayoutID=A.LayoutID where A.Deleted=False and A.Passed=True And A.Elite=True order by A.articleid desc"
	Set rsElite= Server.CreateObject("ADODB.Recordset")
	rsElite.open sqlElite,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsElite.bof and rsElite.eof then 
		response.write "<li>无推荐文章</li>" 
	else 
		do while not rsElite.eof   
			response.Write " - <a href='" & rsElite("LayoutFileName") & "?ArticleID=" & rsElite("articleid") &"' title='文章标题：" & rsElite("Title") & vbcrlf & "作    者：" & rsElite("Author") & vbcrlf & "更新时间：" & rsElite("UpdateTime") & vbcrlf & "点击次数：" & rsElite("Hits") & "' target='_blank'>" & gotTopic(rsElite("title"),TitleLen) & "</a>[<font color=red>" & rsElite("hits") & "</font>]<br>"
        	rsElite.movenext     
		loop
	end if  
	rsElite.close
	set rsElite=nothing
end sub

'=================================================
'过程名：ShowCorrelative
'作  用：显示相关文章
'参  数：ArticleNum  ----最多显示多少篇文章
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowCorrelative(ArticleNum,TitleLen)
	dim rsCorrelative,sqlCorrelative
	dim strKey,arrKey,i
	if ArticleNum>0 and ArticleNum<=100 then
		sqlCorrelative="select top " & ArticleNum
	else	
		sqlCorrelative="Select Top 5 "
	end if
	strKey=mid(rs("Key"),2,len(rs("Key"))-2)
	if instr(strkey,"|")>1 then
		arrKey=split(strKey,"|")
		strKey="((A.Key like '%|" & arrKey(0) & "|%')"
		for i=1 to ubound(arrKey)
			strKey=strKey & " or (A.Key like '%|" & arrKey(i) & "|%')"
		next
		strKey=strKey & ")"
	else
		strKey="(A.Key like '%|" & strKey & "|%')"
	end if
	sqlCorrelative=sqlCorrelative & " A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName From Article A inner join Layout L on L.LayoutID=A.LayoutID Where A.Deleted=False and A.Passed=True and " & strKey & " and A.ArticleID<>" & ArticleID & " Order by A.ArticleID desc"
	Set rsCorrelative= Server.CreateObject("ADODB.Recordset")
	rsCorrelative.open sqlCorrelative,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsCorrelative.bof and rsCorrelative.Eof then
		response.write "没有相关文章"
	else
	 	do while not rsCorrelative.eof	
			response.write "&nbsp;&nbsp; - <a href='" & rsCorrelative("LayoutFileName") & "?ArticleID=" & rsCorrelative("ArticleID") & "' title='文章标题：" & rsCorrelative("Title") & vbcrlf & "作    者：" & rsCorrelative("Author") & vbcrlf & "更新时间：" & rsCorrelative("UpdateTime") & vbcrlf & "点击次数：" & rsCorrelative("Hits") & "'>" & gotTopic(rsCorrelative("Title"),TitleLen) & "</a>[<font color=red>" & rsCorrelative("hits") & "</font>]<br>"
			rsCorrelative.movenext
		loop
	end if
	rsCorrelative.close
	set rsCorrelative=nothing
end sub


'++++++++++++++++++++++++++++
'显示作者名及其链接

sub ShowAuthorName()

		set tUser=Conn_User.execute("select " & db_User_ID & " from " & db_User_Table & " where " & db_User_Name & "='" & rs("Editor") & "'")
		if tUser.bof and tUser.eof then
			response.write rs("Editor")
		else
			response.write "<strong>" & rs("Editor") & "</strong>"
		end if
		
end sub
'++++++++++++++++++++++++++++++

'=================================================
'过程名：ShowComment
'作  用：显示相关评论
'参  数：CommentNum  ----最多显示多少个评论
'=================================================
sub ShowComment(CommentNum)
	dim rsComment,sqlComment,rsCommentUser
	if CommentNum>0 and CommentNum<=100 then
		sqlComment="select top " & CommentNum
	else
		sqlComment="select top 10 "
	end if
	sqlComment=sqlComment & " * from ArticleComment where ArticleID=" & ArticleID & " order by CommentID desc"
	Set rsComment= Server.CreateObject("ADODB.Recordset")
	rsComment.open sqlComment,conn,1,1
	if rsComment.bof and rsComment.eof then
		response.write "没有任何评论"
	else
		response.write "<a href='Article_CommentShow.asp?ArticleID=" & ArticleID & "'><font color=blue>查看此文的所有评论</font></a>"
	end if
end sub

'=================================================
'过程名：ShowPrevArticle
'作  用：显示上一篇文章
'参  数：TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowPrevArticle(TitleLen)
	dim rsPrev,sqlPrev
	sqlPrev="Select Top 1 A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName From Article A inner join Layout L on A.LayoutID=L.LayoutID Where Deleted=False and Passed=True and ClassID=" & rs("ClassID") & " and ArticleID<" & rs("ArticleID")& " order by ArticleID DESC"
	Set rsPrev= Server.CreateObject("ADODB.Recordset")
	rsPrev.open sqlPrev,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsPrev.Eof then
	  	response.write "没有了"
	else
		response.write "<a href='" & rsPrev("LayoutFileName") & "?ArticleID=" & rsPrev("ArticleID")& "' title='文章标题：" & rsPrev("Title") & vbcrlf & "作    者：" & rsPrev("Author") & vbcrlf & "更新时间：" & rsPrev("UpdateTime") & vbcrlf & "点击次数：" & rsPrev("Hits") &"'>" & gotTopic(rsPrev("Title"),TitleLen) & "</a>"
	end if
	rsPrev.close
	set rsPrev=nothing
end sub

'=================================================
'过程名：ShowNextArticle
'作  用：显示上一篇文章
'参  数：TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'=================================================
sub ShowNextArticle(TitleLen)
	dim rsNext,sqlNext
	sqlNext="Select Top 1 A.ArticleID,A.Title,A.Author,A.UpdateTime,A.Hits,L.LayoutFileName From Article A inner join Layout L on A.LayoutID=L.LayoutID Where Deleted=False and Passed=True and ClassID=" & rs("ClassID") & " and ArticleID>" & rs("ArticleID")& " order by ArticleID ASC"
	Set rsNext= Server.CreateObject("ADODB.Recordset")
	rsNext.open sqlNext,conn,1,1
	if TitleLen<0 or TitleLen>255 then TitleLen=50
	if rsNext.Eof then
	 	response.write "没有了"
	else
	  	response.write "<a href='"& rsNext("LayoutFileName") & "?ArticleID="&rsNext("ArticleID")& "' title='文章标题：" & rsNext("Title") & vbcrlf & "作    者：" & rsNext("Author") & vbcrlf & "更新时间：" & rsNext("UpdateTime") & vbcrlf & "点击次数：" & rsNext("Hits") &"'>" & gotTopic(rsNext("Title"),TitleLen) & "</a>"
	end if
	rsNext.close
	set rsNext=nothing
end sub

'=================================================
'过程名：ShowPicArticle
'作  用：显示图片文章
'参  数：intClassID  ----栏目ID，0为所有栏目，若大于0，则显示指定栏目及其子栏目的图片文章
'        ArticleNum  ----最多显示多少篇文章
'        TitleLen   ----标题最多字符数，一个汉字=两个英文字符
'        ShowType   ----显示方式。1为只有图片+标题，2为图片+标题+内容简介
'        Cols       ----列数。超过此列数就换行。
'        ImgWidth   ----图片宽度
'        ImgHeight  ----图片高度
'        ContentLen ----内容最多字符数
'        Hot        ----是否是热门文章
'        Elite      ----是否是推荐文章
'=================================================
sub ShowPicArticle(intClassID,ArticleNum,TitleLen,ShowType,Cols,ImgWidth,ImgHeight,ContentLen,Hot,Elite)
	dim sqlPic,i,tClass,trs,arrClassID
	if ArticleNum<0 or ArticleNum>=50 then
		ArticleNum=5
	end if
	if ShowType<>1 and ShowType<>2 then
		ShowType=1
	end if
	if Cols<=0 or Cols>=10 then
		Cols=5
	end if
	if ImgWidth<0 or ImgWidth>500 then
		ImgWidth=150
	end if
	if ImgHeight<0 or ImgHeight>500 then
		ImgHeight=150
	end if
	if Hot<>True and Hot<>False then
		Hot=False
	end if
	if Elite<>True and Elite<>False then
		Elite=False
	end if

	sqlPic="select top " & ArticleNum
	sqlPic=sqlPic & " A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
	if ShowType=2 then
		sqlPic=sqlPic & "A.Content,"
	end if
	sqlPic=sqlPic & " A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sqlPic=sqlPic & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and DefaultPicUrl<>''"

	if intClassID>0 then
		set tClass=conn.execute("select ClassID,Child,ParentPath from ArticleClass where ClassID=" & intClassID)
		if not(tClass.bof and tClass.eof) then
			if tClass(1)>0 then
				arrClassID=ClassID
				set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & tClass(0) & " or ParentPath like '%" & tClass(2) & "," & tClass(0) & ",%' and Child=0 and LinkUrl=''")
				do while not trs.eof
					arrClassID=arrClassID & "," & trs(0)
					trs.movenext
				loop
				set trs=nothing	
				sqlPic=sqlPic & " and A.ClassID in (" & arrClassID & ")"
			else
				sqlPic=sqlPic & " and A.ClassID=" & tClass(0)
			end if
			set trs=nothing
		else
			sqlPic=sqlPic & " and A.ClassID=" & tClass(0)
		end if
		set tClass=nothing
	end if
	if Hot=True then
		sqlPic=sqlPic & " and A.Hits>=" & HitsOfHot
	end if
	if Elite=True then
		sqlPic=sqlPic & " and A.Elite=True "
	end if
	sqlPic=sqlPic & " order by A.OnTop,A.ArticleID desc"
	set rsPic=Server.CreateObject("ADODB.Recordset")
	rsPic.open sqlPic,conn,1,1
	strPic= "<table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
	if rsPic.bof and rsPic.eof then
		strPic= strPic & "<td align='center'><img src='images/NoPic.jpg' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>没有任何图片文章</td>"
	else
		i=0
		if ShowType=1 then
			do while not rsPic.eof
				strPic=strPic & "<td align='center'>"
				call GetPicArticleTitle(TitleLen,ImgWidth,ImgHeight)
				strPic=strPic & "</td>"
				rsPic.movenext
				i=i+1
				if ((i mod Cols=0) and (not rsPic.eof)) then strPic=strPic & "</tr><tr valign='top'>"
			loop
		elseif ShowType=2 then
			do while not rsPic.eof
				strPic=strPic & "<td align='center'>"
				call GetPicArticleTitle(TitleLen,ImgWidth,ImgHeight)
				strPic=strPic & "</td><td valign='top' algin='center' class='left'><a href='" & rsPic("LayoutFileName") & "?ArticleID=" & rsPic("ArticleID") & "'>" & left(nohtml(rsPic("Content")),ContentLen) & "……</a></td>"
				rsPic.movenext
				i=i+1
				if ((i mod Cols=0) and (not rsPic.eof)) then strPic=strPic & "</tr><tr valign='top'>"
			loop
		end if
		
	end if
	strPic=strPic &  "</tr></table>"
	response.write strPic
	rsPic.close
end sub

'=================================================
'过程名：GetPicArticleTitle
'作  用：显示图片文章的标题
'参  数：intTitleLen  ----标题最多字符数，一个汉字=两个英文字符
'        intImgWidth   ----图片宽度
'        intImgHeight  ----图片高度
'=================================================
sub GetPicArticleTitle(intTitleLen,intImgWidth,intImgHeight)
	dim FileType,TitleStr
	FileType=right(lcase(rsPic("DefaultPicUrl")),3)
	TitleStr=gotTopic(rsPic("Title"),intTitleLen)
	strPic=strPic & "<a href='" & rsPic("LayoutFileName") & "?ArticleID=" & rsPic("ArticleID") & "' title='文章标题：" & rsPic("Title") & vbcrlf & "作    者：" & rsPic("Author") & vbcrlf & "更新时间：" & rsPic("UpdateTime") & vbcrlf & "点击次数：" & rsPic("Hits") & "' target='_blank'>"
	if FileType="swf" then
		strPic=strPic & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='" & intImgWidth & "' height='" & intImgHeight & "'><param name='movie' value='" & rsPic("DefaultPicUrl") & "'><param name='quality' value='high'><embed src='" & rsPic("DefaultPicUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='" & intImgWidth & "' height='" & intImgHeight & "'></embed></object>"
	elseif fileType="jpg" or fileType="bmp" or fileType="png" or fileType="gif" then
		strPic=strPic & "<img src='" & rsPic("DefaultPicUrl") & "' width='" & intImgWidth & "' height='" & intImgHeight & "' border='0'>"
	else
		strPic=strPic & "<img src='images/NoPic2.jpg' width='" & intImgWidth & "' height='" & intImgHeight & "' border='0'>"
	end if
	if rsPic("TitleFontType")=1 then
		TitleStr="<b>" & TitleStr & "</b>"
	elseif rsPic("TitleFontType")=2 then
		TitleStr="<em>" & TitleStr & "</em>"
	elseif rsPic("TitleFontType")=3 then
		TitleStr="<b><em>" & TitleStr & "</em></b>"
	end if
	if rsPic("TitleFontColor")<>"" then
		TitleStr="<font color='" & rsPic("TitleFontColor") & "'>" & TitleStr & "</font>"
	end if
	strPic=strPic & "<br>" & TitleStr & "</a>"
end sub



'=================================================
'过程名：ShowArticleContent
'作  用：显示文章具体的内容，可以分页显示
'参  数：无
'=================================================
sub ShowArticleContent()
	if rs("ReadLevel")<=999 then
		if UserLogined<>True then 
			FoundErr=True
			ErrMsg=ErrMsg & "<br>&nbsp;&nbsp;&nbsp;&nbsp;你还没注册？或者没有登录？这篇文章要求至少是本站的注册用户才能阅读！<br><br>"
			ErrMsg=ErrMsg & "&nbsp;&nbsp;&nbsp;&nbsp;如果你还没注册，请赶紧<a href='User_Reg.asp'><font color=red>点此注册</font></a>吧！<br><br>"
			ErrMsg=ErrMsg & "&nbsp;&nbsp;&nbsp;&nbsp;如果你已经注册但还没登录，请赶紧<a href='User_Login.asp'><font color=red>点此登录</font></a>吧！<br><br>"
		else
			if UserLevel>rs("ReadLevel") then
				FoundErr=True
				ErrMsg=ErrMsg & "<p align='center'><br><br><font color=red><b>对不起，你的权限不够，不能阅读此文章!</b></font></p>"
			else
				if ChargeType=1 and rs("ReadPoint")>0 then
					if Request.Cookies("asp163")("Pay_Article" & ArticleID)<>"yes" then
						if UserPoint<rs("ReadPoint") then
							FoundErr=True
							ErrMsg=ErrMsg &"<p align='center'><br><br>对不起，阅读本文需要消耗 <b><font color=red>" & rs("ReadPoint") & "</font></b> 点！"
							ErrMsg=ErrMsg &"而你目前只有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。点数不足，无法阅读本文。请与我们联系进行充值。</p>"
						else
							if lcase(trim(request("Pay")))="yes" then
								Conn_User.execute "update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & rs("ReadPoint") & " where " & db_User_Name & "='" & UserName & "'"
								response.Cookies("asp163")("Pay_Article" & ArticleID)="yes"
							else
								FoundErr=True
								ErrMsg=ErrMsg &"<p align='center'><br><br>阅读本文需要消耗 <b><font color=red>" & rs("ReadPoint") & "</font></b> 点！"
								ErrMsg=ErrMsg &"你目前尚有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。阅读本文后，你将剩下 <b><font color=green>" & UserPoint-rs("ReadPoint") & "</font></b> 点"
								ErrMsg=ErrMsg &"<br><br>你确实愿意花费 <b><font color=red>" & rs("ReadPoint") & "</font></b> 点来阅读本文吗？"
								ErrMsg=ErrMsg &"<br><br><a href='"& strFileName & "?Pay=yes&ArticleID=" & ArticleID & "'>我愿意</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='index.asp'>我不愿意</a></p>"
							end if
						end if
					end if
				elseif ChargeType=2 then
					if ValidDays<=0 then
						FoundErr=True
						ErrMsg=ErrMsg & "<p align='center'><br><br><font color=red>对不起，本文为收费内容，而您的有效期已经过期，所以无法阅读本文。请与我们联系进行充值。</font></p>"
					end if
				end if
			end if
		end if
	end if

	if FoundErr=True then
		ErrMsg="<p align=left><b>内容预览：</b><br><br>" & left(nohtml(rs("Content")),300) & "……</p>" & ErrMsg
		response.write ErrMsg
		exit sub
	end if
	
	dim PaginationType
	PaginationType=rs("PaginationType")
	select case PaginationType
		case 0    '不分页显示
			response.write rs("Content")
		case 1    '自动分页显示
			call AutoPagination()
		case 2    '手动分页显示
			call ManualPagination()
	end select
end sub

'=================================================
'过程名：ManualPagination
'作  用：采用手动分页方式显示文章具体的内容
'参  数：无
'=================================================
sub ManualPagination()
	dim ArticleID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i
	dim arrContent
	ArticleID=rs("ArticleID")
	strContent=rs("Content")
	ContentLen=len(strContent)
	CurrentPage=trim(request("ArticlePage"))
	if Instr(strContent,"[NextPage]")<=0 then
		response.write strContent
		response.write "</p><p align='center'><font color='red'><b>[1]</b></font></p>"
	else
		arrContent=split(strContent,"[NextPage]")

		pages=Ubound(arrContent)+1
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		response.write arrContent(CurrentPage-1)

		response.write "</p><p align='center'><b>"
		if CurrentPage>1 then
			response.write "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage-1 & "'>上一页</a>&nbsp;&nbsp;"
		end if
		for i=1 to pages
			if i=CurrentPage then
				response.write "<font color='red'>[" & cstr(i) & "]</font>&nbsp;"
			else
				response.write "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & i & "'>[" & i & "]</a>&nbsp;"
			end if
		next
		if CurrentPage<pages then
			response.write "&nbsp;<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage+1 & "'>下一页</a>"
		end if
		response.write "</b></p>"
		
	end if

end sub
'=================================================
'过程名：AutoPagination
'作  用：采用自动分页方式显示文章具体的内容
'参  数：无
'=================================================
sub AutoPagination()
	dim ArticleID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleID=rs("ArticleID")
	strContent=rs("Content")
	ContentLen=len(strContent)
	CurrentPage=trim(request("ArticlePage"))
	if ContentLen<=rs("MaxCharPerPage") then
		response.write strContent
		response.write "</p><p align='center'><font color='red'><b>[1]</b></font></p>"
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\rs("MaxCharPerPage")
		if rs("MaxCharPerPage")*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=ContentLen          '最大误差范围
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=rs("MaxCharPerPage")*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=rs("MaxCharPerPage")*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		response.write mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		response.write "</p><p align='center'><b>"
		if CurrentPage>1 then
			response.write "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage-1 & "'>上一页</a>&nbsp;&nbsp;"
		end if
		for i=1 to pages
			if i=CurrentPage then
				response.write "<font color='red'>[" & cstr(i) & "]</font>&nbsp;"
			else
				response.write "<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & i & "'>[" & i & "]</a>&nbsp;"
			end if
		next
		if CurrentPage<pages then
			response.write "&nbsp;<a href='" & strFileName & "?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage+1 & "'>下一页</a>"
		end if
		response.write "</b></p>"
	end if

end sub

%>
