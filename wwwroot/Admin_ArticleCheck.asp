<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=3
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim strFileName,FileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim i,j
dim keyword,strField
dim sql,rsArticleList
dim ClassID,Passed
dim PurviewChecked
dim strAdmin,arrAdmin
dim tClass,ClassName,RootID,ParentID,Depth,ParentPath,Child,ChildID,tID,tChild,ClassMaster
dim ManageType
ManageType=trim(request("ManageType"))
PurviewChecked=false
FileName="Admin_ArticleCheck.asp"
ClassID=Trim(request("ClassID"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=ReplaceBadChar(keyword)
end if
strField=trim(request("Field"))
Passed=trim(request("Passed"))
if Passed="" then
	if Session("Passed")="" then
		Passed="False"
	else
		Passed=Session("Passed")
	end if
end if
session("Passed")=Passed
if ClassID="" then
	ClassID=0
	if strField="" and  (AdminPurview=2 or AdminPurview=3 )and AdminPurview_Article=3 then
		set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassChecker,ClassID From ArticleClass where ClassMaster like '%" & AdminName & "%'")
		do while not tClass.eof
			if CheckClassMaster(tClass(6),AdminName)=true then
				ClassName=tClass(0)
				RootID=tClass(1)
				ParentID=tClass(2)
				Depth=tClass(3)
				ParentPath=tClass(4)
				Child=tClass(5)
				ClassMaster=tClass(6)
				ClassID=tClass(7)
				PurviewChecked=True
				exit do
			end if
			tClass.movenext
		loop
	end if
else
	ClassID=CLng(ClassID)
end if
if ClassID>0 and PurviewChecked=False then
	set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,Child,ClassChecker From ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
		Call WriteErrMsg()
		response.end
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		Child=tClass(5)
		ClassMaster=tClass(6)
		PurviewChecked=CheckClassMaster(tClass(6),AdminName)
		if PurviewChecked=False and ParentID>0 then
			set tClass=conn.execute("select ClassChecker from ArticleClass where ClassID in (" & ParentPath & ")")
			do while not tClass.eof
				PurviewChecked=CheckClassMaster(tClass(0),AdminName)
				if PurviewChecked=True then exit do
				tClass.movenext
			loop
		end if
	end if
end if
strFileName=FileName & "?ClassID=" & ClassID & "&strField=" & strField & "&keyword=" & keyword
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
%>
<html>
<head>
<title>文章管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2"  align="center"><strong>文 章 审 核</strong></td>
  </tr>
  <form name="Passed" method="Post" action="<%=strFileName%>">
    <tr class="tdbg">
      <td width="70" height="30" ><strong>文章选项：</strong></td>
      <td ><input name="Passed" type="radio" value="False" onClick="submit();" <%if Passed="False" then response.write " checked"%>>
        未审核的文章&nbsp;&nbsp;&nbsp;&nbsp;
        <input name="Passed" type="radio" value="True" onClick="submit();" <%if Passed="True" then response.write " checked"%>>
        已审核的文章 </td>
    </tr>
  </form>
</table>
<br/>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title">
    <td height="22"><%call Admin_ShowRootClass()%></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="22"><%call Admin_ShowPath("文章审核")%></td>
    <td width="200" height="22" align="right"><select name='JumpClass' id="JumpClass" onChange="if(this.options[this.selectedIndex].value!=''){location='<%=FileName & "?ClassID="%>'+this.options[this.selectedIndex].value;}">
        <option value='' selected>跳转栏目至…</option>
        <%call Admin_ShowClass_Option(1,0)%>
      </select>
    </td>
  </tr>
</table>
<%



sql="select A.ArticleID,A.ClassID,C.ClassName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.SpecialID," '补上课程编号列
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint from Article A"
'判定权限
if AdminPurview=1 then
	sql=sql & " inner join ArticleClass C on A.ClassID=C.ClassID where A.Deleted=False and A.nopass=False and Passed=" & Passed
else if AdminPurview=2 then
	sql=sql & " inner join ArticleClass C on A.ClassID=C.ClassID where A.Deleted=False and A.nopass=False and Passed=" & Passed & " and ( A.TeacherName='" & session("AdminTeacherName") & "' or A.TeacherName='admin' )"
else if AdminPurview=3 then
		sql=sql & " inner join ArticleClass C on A.ClassID=C.ClassID where A.Deleted=False and A.nopass=False and Passed=" & Passed & " and A.TeacherName='" & session("AdminTeacherName") & "'"

	'sql=sql & " inner join ArticleClass C on A.ClassID=C.ClassID where A.Deleted=False and A.nopass=False and Passed=" & Passed & " and A.TeacherName='" & session("AdminTeacherName") & "' and  A.SpecialID like '%" & session("AdminPurview_SpecialID") & "%' "
end if
end if
end if
'结束判定
if ClassID>0 then
	if Child>0 then
		ChildID=""
		set tClass=conn.execute("select ClassID from ArticleClass where ParentID=" & ClassID & " or ParentPath like '" & ParentPath & "," & ClassID & ",%'")
		do while not tClass.eof
			if ChildID="" then
				ChildID=tClass(0)
			else
				ChildID=ChildID & "," & tClass(0)
			end if
			tClass.movenext
		loop
		sql=sql & " and A.ClassID in (" & ChildID & ")"
	else
		sql=sql & " and A.ClassID=" & ClassID
	end if
end if

if strField<>"" then
	if keyword<>"" then
		select case strField
		case "Title"
			sql=sql & " and A.Title like '%" & keyword & "%' "
		case "Content"
			sql=sql & " and A.Content like '%" & keyword & "%' "
		case "Author"
			sql=sql & " and A.Author like '%" & keyword & "%' "
		case "Editor"
			sql=sql & " and A.Editor like '%" & keyword & "%' "
		case else
			sql=sql & " and A.Title like '%" & keyword & "%' "
		end select
	end if
end if
sql=sql & " order by A.ArticleID desc"

Set rsArticleList= Server.CreateObject("ADODB.Recordset")
'response.write sql
'response.end

rsArticleList.open sql,conn,1,1
'response.Write(sql)
if rsArticleList.eof and rsArticleList.bof then
	totalPut=0
	if Child=0 then
		if Passed="True" then
			response.write "<p align='center'><br>没有任何已审核的文章！<br></p>"
		else
			response.write "<p align='center'><br>没有任何待审核的文章！<br></p>"
		end if
	else
		if Passed="True" then
			response.write "<p align='center'><br>此栏目的下一级子栏目中没有任何已审核的文章！<br></p>"
		else
			response.write "<p align='center'><br>此栏目的下一级子栏目中没有任何待审核的文章！<br></p>"
		end if
	end if
else
   	totalPut=rsArticleList.recordcount
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
       	showContent
       	showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	rsArticleList.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=rsArticleList.bookmark
            showContent
            showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
       	else
	        currentPage=1
           	showContent
          	showpage strFileName,totalput,MaxPerPage,true,true,"篇文章"
	    end if
	end if
end if
rsArticleList.close
set rsArticleList=nothing  


sub showContent
   	dim ArticleNum
    ArticleNum=0
%>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
    <form name="myform" method="Post" action="Admin_ArticleProperty.asp">
      <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22">
            <td height="22" width="30" align="center"><strong>选中</strong></td>
            <td width="25" align="center"  height="22"><strong>ID</strong></td>
            <td align="center" ><strong>文章标题</strong></td>
            <td width="60" align="center" ><strong>录入</strong></td>
            <td width="40" align="center" ><strong>点击数</strong></td>
            <td width="60" align="center" ><strong>文章属性</strong></td>
            <td width="40" align="center" ><strong>已审核</strong></td>
            <td width="120" align="center" ><strong>操作</strong></td>
          </tr>
           <% '开始判断权限。前面的SQL查询语句要补上SpecialID列.从数据库获得课程名.
				  
		 'dim rsArticleList_Special , sql_Special
		
		
		 %>
		  <%do while not rsArticleList.eof%>
          <!-- 两课网站代码，根据课程判断权限。Admin_ChkPurview.asp判断管理员权限。三段代码开始,定义变量在上一行-->
          <%
		  '判断是否属于某课程
'		if   ( rsArticleList("SpecialID") = "" ) then
'			sql_Special="select SpecialID ,SpecialName from Special where SpecialID=0"
'		
'		else
'			sql_Special="select SpecialID ,SpecialName from Special where SpecialID=" &  rsArticleList("SpecialID")
'		
'		end if
		'结束判断是否属于某课程
		 
'		 Set rsArticleList_Special= Server.CreateObject("ADODB.Recordset")
'			rsArticleList_Special.open sql_Special,conn,1,3 
'		  if  rsArticleList("SpecialID")=0 or AdminPurview=1  or ( not     rsArticleList_Special.eof  )  then 
'		  	 
'				if rsArticleList("SpecialID")=0 or AdminPurview=1 or (instr( AdminPurview_Special , rsArticleList_Special("SpecialName"))>0)   then	
'			 		if 1 < 3  then
			  %><!--结束判断-->
          <!-- 'instr()用法是后者在前者之中，举例：(instr(rsAdminPurview_Special("AdminPurview_Special"),rsSpecial("SpecialName")) >0)-->
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td width="30" align="center"><input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rsArticleList("articleID"))%>'></td>
            <td width="25" align="center"><%=rsArticleList("articleid")%></td>
            <td><%
			if rsArticleList("ClassID")<>ClassID then
				response.write "<a href='" & FileName & "?ClassID=" & rsArticleList("ClassID") & "'>[" & rsArticleList("ClassName") & "]</a>&nbsp;"
			end if
			if rsArticleList("IncludePic")=true then
				response.write "<font color=blue>[图文]</font>"
			end if
			response.write "<a href='Admin_ArticleShow.asp?ArticleID=" & rsArticleList("articleid") & "'"
			response.write " title='标    题：" & rsArticleList("Title") & vbcrlf & "作    者：" & rsArticleList("Author") & vbcrlf & "转 贴 自：" & rsArticleList("CopyFrom") & vbcrlf & "更新时间：" & rsArticleList("UpdateTime") & vbcrlf
			response.write "点 击 数：" & rsArticleList("Hits") & vbcrlf & "关 键 字：" & mid(rsArticleList("Key"),2,len(rsArticleList("Key"))-2) & vbcrlf & "推荐等级："
			if rsArticleList("Stars")=0 then
				response.write "无"
			else
				response.write string(rsArticleList("Stars"),"★")
			end if			
			response.write vbcrlf & "分页方式："
			if rsArticleList("PaginationType")=0 then
				response.write "不分页"
			elseif rsArticleList("PaginationType")=1 then
				response.write "自动分页"
			elseif rsArticleList("PaginationType")=2 then
				response.write "手动分页"
			end if
			response.write vbcrlf & "阅读等级："	
			if rsArticleList("ReadLevel")=9999 then
				response.write "游客"
			elseif  rsArticleList("ReadLevel")=999 then
				response.write "注册用户"
			elseif  rsArticleList("ReadLevel")=99 then
				response.write "收费用户"
			elseif  rsArticleList("ReadLevel")=9 then
				response.write "VIP用户"
			elseif  rsArticleList("ReadLevel")=5 then
				response.write "管理员"
			end if
			response.write vbcrlf & "阅读点数：" & rsArticleList("ReadPoint")
			response.write "'>" & rsArticleList("title") & "</a>"
			%></td>
            <td width="60" align="center"><%
			response.write "<a href='" & FileName & "?field=Editor&keyword=" & rsArticleList("Editor") & "' title='点击将查看此用户录入的所有文章'>" & rsArticleList("Editor") & "</a>"
			%>
            </td>
            <td width="40" align="center"><%= rsArticleList("Hits") %></td>
            <td width="60" align="center"><%
			if rsArticleList("OnTop")=true then
				response.Write "<font color=blue>顶</font> "
			else
				response.write "&nbsp;&nbsp;&nbsp;"
			end if
			if rsArticleList("Hits")>=HitsOfHot then
				response.write "<font color=red>热</a> "
			else
				response.write "&nbsp;&nbsp;&nbsp;"
			end if
			if rsArticleList("Elite")=true then
				response.write "<font color=green>荐</a>"
			else
				response.write "&nbsp;&nbsp;"
			end if
			%>
            </td>
            <td width="40" align="center"><%
			if rsArticleList("Passed")=true then
				response.write "是"
			else
				response.write "否"
			end if%></td>
            <td width="120" align="center"><%
	  if AdminPurview=1 or AdminPurview_Article<=2 or PurviewChecked=true then
	  	If rsArticleList("Passed")=False Then
        	response.write "<a href='Admin_ArticleProperty.asp?Action=SetPassed&ArticleID=" & rsArticleList("ArticleID") & "'>通过审核</a>"
		'退稿插件添加代码开始
		response.write "&nbsp;<a href='Admin_Article_nopass.asp?ArticleID=" & rsArticleList("ArticleID") & "'>退稿</a>"
		response.write "&nbsp;<a href='Admin_ArticleProperty.asp?Action=deleted&ArticleID=" & rsArticleList("ArticleID") & "'>删除</a>"
        	'退稿插件添加代码结束
	Else
	        response.write "<a href='Admin_ArticleProperty.asp?Action=CancelPassed&ArticleID=" & rsArticleList("ArticleID") & "'>取消通过</a>"
      	End If
	end if %></td>
          </tr>
          <%
		ArticleNum=ArticleNum+1
	   	if ArticleNum>=MaxPerPage then exit do
	    '两课网站代码判断权限。第四段代码
		
'		end if
'		end if
'		end if
'		
'		rsArticleList_Special.close
' 		set rsArticleList_Special=nothing
		 '结束两课网站代码判断权限
		rsArticleList.movenext
	loop
	'两课网站代码写出权限提示
   if  ArticleNum=0 then
    response.Write("<font color=blue><div align=center><strong>在您的权限范围内没有符合条件的文章。</strong></div></font>")
   end if
 
   '结束两课网站代码写出权限提示.五段代码结束
%>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有文章 </td>
            <td><input name="submit" type='submit' value="<%if Passed="True" then response.write "取消"%>审核选定的文章" <%if AdminPurview=2 and PurviewChecked=False then response.write "disabled"%>>
              <input name="Action" type="hidden" id="Action" value="<% if Passed="False" then
			  	response.write "SetPassed"
			  else
			    response.write "CancelPassed"
			  end if%>">
              &nbsp;&nbsp;
              <input name="submit" type=<% if Passed="True" then response.write "hidden" else response.write "submit" %> value="删除选定的文章" onClick="document.myform.Action.value='deleted'" <%if AdminPurview=2 and PurviewChecked=False or Passed="True" then response.write "disabled"%>>
            </td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<%
end sub 

if ClassID>0 and Child>0 then
%>
<br>
<table width="100%" height="5" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class='border'>
  <tr height="20" class='tdbg'>
    <td width='150' align="right">【
      <%response.write "<a href='" & strFileName & "'>" & ClassName & "</a>"%>
      】子栏目导航：</td>
    <td><%call Admin_ShowChild()%></td>
  </tr>
</table>
<%
end if
%>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
  <tr class="tdbg">
    <td width="80" align="right"><strong>文章搜索：</strong></td>
    <td><%call Admin_ShowSearchForm(FileName,2)%>
    </td>
  </tr>
</table>
</body>
</html>
<%
call CloseConn()
%>
