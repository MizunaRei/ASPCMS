<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=2
Const PurviewLevel_Article=1
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
'Response.Write("Line 16")
'Response.End()
dim Action,FoundErr,ErrMsg
dim rs,sql
dim SkinCount,LayoutCount

'开课表的添加。定义变量。接收变量
'dim  StudentClass , StudentClassName , StudentClassYear , StudentClassNumber
'dim College  'dim TeacherName 此变量已定义,未知在何处定义

'dim  TermNumber
Action=trim(request("Action"))



'接收变量
'College=Trim(Request("College"))
'StudentClass=Trim(Request("StudentClassName")) & Trim(Request("StudentClassYear"))  &  Trim(Request("StudentClassNumber"))
'StudentClassName=Trim(Request("StudentClassName"))
'StudentClassYear=Trim(Request("StudentClassYear"))
'StudentClassNumber=Trim(Request("StudentClassNumber"))

'TeacherName=Trim(Request("TeacherName"))

'TermNumber=Trim(Request("TermNumber"))
'Response.Write("Line 38")
'Response.End()
%>
<html>
<head>
<title>开课时间班级管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin_Style.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="3" align="center"><strong>开课时间班级管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td><a href="Admin_Course.asp">开课时间班级 | 管理</a></td><td align=right>鼠标停在文字上将弹出提示</td>
  </tr>
</table>
<%
if Action="Add" then
	call AddSpecial()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelSpecial()
elseif Action="Clear" then
	call ClearSpecial()
elseif Action="UpOrder" then 
	call UpOrder() 
elseif Action="DownOrder" then 
	call DownOrder() 
elseif Action="Unite" then
	call ShowUniteForm()
elseif Action="UniteSpecial" then
	call UniteSpecial()
elseif Action="Order" then
	call ShowOrder()
	
	
	'思想政治理论课新增功能，开课表
elseif Action="CourseDel" then
	call CourseDel()
elseif Action="CourseAdd" then
	call CourseAdd()
	
	
	'新增功能就这些了
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
if 	AdminPurview=3 then
 call Admin3Course()
 elseif AdminPurview=2 then
 call Admin2Course()
 elseif AdminPurview=1 then
 call Admin1Course()
 end if
%>
<br/>
<%
	
end sub

%>
</body>
</html>
<%
sub Admin3Course()
FoundErr=True
		
		ErrMsg=ErrMsg & "<br><li>学生管理员没有权限管理开课安排！</li><!--<li> <a Href=Admin_Course.asp>返回</a> </li>-->"
end sub

'超级管理员的网页主体代码
sub Admin1Course()
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <form name='Admin_Course' action='Admin_Course.asp' method='post' id='Admin_Course'>
    <tr class="tdbg" align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"  >
      <!--集中在某一行设定列宽-->
      <td  align="center"  ><%  call TeacherTrueNameComboBox()	   %></td>
      <td align="center" ><% call SpecialAbbreviationComboBox()  %></td>
      <td><!--课程说明不用填，就在课程设置 里有--></td>
      <td><!--开课时间下拉框-->
        <% call TermYear() %>
        &nbsp;
        <% call TermOrder() %></td>
      <td><!--开课班级-->
        <% call StudentClassName() %>
        &nbsp;
        <% call StudentClassYear() %>
        &nbsp;
        <% call StudentClassNumber() %>
        &nbsp;
        <% call College() %></td>
      <td align="center" ><input type='hidden' name='Action' value='CourseAdd' />
        <input name="Submit" type="submit" id="Submit" value="新增此行" /></td>
    </tr>
  </form>
  <!--先列出同一老师教同一课程的多个班-->
  <%
    dim rsAdminList 
  	dim rsSpecialList
	'思想政治理论课教学平台新增功能开课时间班级管理的变量声明
	dim sqlCourseList
	dim rsCourseList
	'
	'Response.End()
	set rsAdminList=conn.execute("select  ID , TrueName ,TeacherName from Admin where Purview=1 or Purview=2")
	'Response.Write("Line 108")
	
	'每位教师都要做一次循环查一次开课表
	do while not rsAdminList.eof
	%>
  <tr class="title">
    <!--每个教师的课以独立的一个表呈现-->
    <!-- 此行列出教师所开课程的列表的表头-->
    <td height="22"  align="center" width="100"><strong>教师姓名</strong></td>
    <td height="22"  align="center" width="100"><strong>课程名称</strong></td>
    <td    align="center" width="70"><strong>课程说明</strong></td>
    <td width="100" align="center"><strong>开课时间</strong></td>
    <td   width="200" align="center"><strong>上课班级</strong></td>
    <td width="100" align="center"><strong>操作</strong></td>
    <!--<td width="110" align="center"><strong>上课班级所属学院</strong></td>-->
    <!--<td width="100" height="22" align="center"><strong> 常规操作</strong></td>-->
  </tr>
  <%
	set rsSpecialList=conn.execute("select SpecialID,SpecialName , ReadMe , SpecialAbbreviation  from Special ")
	if rsSpecialList.bof and rsSpecialList.eof then
		'response.write ("")
		FoundErr=True
		
		ErrMsg=ErrMsg & "<br/><li>课程列表为空，请先添加课程！</li>"
		rsSpecialList=nothing
		exit sub
		
	else
		do while not rsSpecialList.eof '列表为空，错误提示缺少对象
			if   (   InStr( AdminPurview_SpecialID  ,  rsSpecialList("SpecialID")   )   >0 ) then
				Set rsCourseList=Server.CreateObject("Adodb.RecordSet")
				'sql="select * from CourseList , Special , Admin where CourseList.SpecialID=Special.SpecialID and CourseList.TeacherName='" & session("AdminTrueName") &  "'"
				'sqlAdmin="select AdminPurview_SpecialID from Admin"
				sqlCourseList="select * from CourseList where  CourseList.SpecialID=" &      rsSpecialList("SpecialID") & " and TeacherName='" & rsAdminList("TeacherName") & "'"
				rsCourseList.Open sqlCourseList,conn,1,1
				 
			
  %>
  <!--<hr/>-->
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td align="center" rowspan="<%=(rsCourseList.RecordCount+2)%>"><a href="Admin_UserList.asp?Action=CountByTeacher&TeacherName=<%=rsAdminList("TeacherName")%>"  title="点击管理这位老师的学生"><%=rsAdminList("TeacherName")%></a></td>
    <td align="center" rowspan="<%=(rsCourseList.RecordCount+2)%>"><!--显示课程名称,加第一行是为了表对齐，加第二行是为了新增开课表一行-->
      <a href="Admin_UserList.asp?Action=CountBySpecial&SpecialID=<%=rsSpecialList("SpecialID")%>" title="点击管理此课程的学生"><%=rsSpecialList("SpecialAbbreviation")%></a></td>
    <td align="center" rowspan="<%=(rsCourseList.RecordCount+2)%>"><!--显示课程简介，现在未使用-->
      <%=dvhtmlencode(rsSpecialList("ReadMe"))%></td>
    <!---->
  </tr>
  <!--<td colspan="2"><table>-->
  <% 'if rsCourseList.bof and rsCourseList.eof then
					'		response.write ("请先添加开课时间和班级列表")
					'else
         					'response.Write(" <table>  ") %>
  <% 			if  (rsCourseList.bof and rsCourseList.EOF) then
						%>
  <tr class="tdbg"  align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"  >
    <td align="center">未开授此课程</td>
    <td></td>
    <td></td>
  </tr>
  <%
					else
						do while not rsCourseList.EOF %>
  <!--一个课程的多个班-->
  <!--学生的班级的表示方式尚未定，因为不会做下拉列表框-->
  <!--学生的学院的表示方式尚未定，因为不会做下拉列表框-->
  <% 'Response.Write("<tr class=tdbg > <td align=center>") %>
  <tr class="tdbg" align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td align="center"  ><!--显示开课时间，时间用四位数年份和星期数表示-->
      <%=rsCourseList("TermYear")%><%=rsCourseList("TermOrder")%></td>
    <td align='center' ><% 
	  Response.Write(rsCourseList("StudentClassName"))
	  Response.Write("&nbsp;&nbsp;")
	  	  Response.Write("&nbsp;")

	  
	  Response.Write(rsCourseList("StudentClassYear"))
	  	  Response.Write("&nbsp;&nbsp;")

	  Response.Write("&nbsp;")

	  Response.Write(rsCourseList("StudentClassNumber"))
	  	  Response.Write("&nbsp;&nbsp;")

	  Response.Write("&nbsp;")

	  %>
      &nbsp;&nbsp;<%
	  Response.Write( "<a href='Admin_UserList.asp?Action=CountByCollege&College='" & rsCourseList("College") & "' title='点击查看此学院所有学生' >" & rsCourseList("College") & "</a>"    )
	  %>
      <% 'Response.Write(" </td> </tr>") %></td>
    <td align="center" ><a  href="Admin_Course.asp?Action=CourseDel&CourseListID=<%=rsCourseList("CourseListID")%>">删除此行</a></td>
  </tr>
  <%
						rsCourseList.MoveNext
  	 					loop
						'end if
						rsCourseList.Close
						set rsCourseList=Nothing
 		 				%>
  <!--下一个课程-->
  <!--添加开课表-->
  <% 'Response.Write("</table>") %>
  <!--</table></td>-->
  <!-- </tr>-->
  <!--<tr>-->
  <!--</tr>-->
  <% 				end if
   				 end if
		%>
  <tr  ></tr>
  <!--<hr/>-->
  <!--用于增加空白分隔<tr></tr> <tr></tr> <tr></tr> <tr></tr> <tr></tr>-->
  <%
		rsSpecialList.movenext
		loop
	end if
	set rsSpecialList=nothing
  rsAdminList.movenext
  loop
  set rsAdminList=nothing
  
  '添加开课表一项
  %>
</table>
<%
end sub
'结束超级管理员的网页主体
'。权限为2的教师管理员网页主体

sub Admin2Course()
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <form name='Admin_Course' action='Admin_Course.asp' method='post' id='Admin_Course'>
    <tr class="tdbg" align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"  >
      <td align="center" ><% call SpecialAbbreviationComboBox()  %></td>
      <td><!--课程说明不用填，就在课程设置 里有--></td>
      <td><!--开课时间下拉框-->
        <% call TermYear() %>
        &nbsp;
        <% call TermOrder() %></td>
      <td><!--开课班级-->
        <% call StudentClassName() %>
        &nbsp;
        <% call StudentClassYear() %>
        &nbsp;
        <% call StudentClassNumber() %>
        &nbsp;
        <% call College() %></td>
      <td align="center"><input type='hidden' name='Action' value='CourseAdd'>
        <input name="Submit" type="submit" id="Submit" value="新增此行"></td>
    </tr>
  </form>
  <!--先列出同一老师教同一课程的多个班-->
  <%
  	dim RowCounter
	dim rsSpecialList
	'思想政治理论课教学平台新增功能开课时间班级管理的变量声明
	dim sqlCourseList
	dim rsCourseList
	'Response.Write("Line 108")
	'Response.End()
	RowCounter=1
	set rsSpecialList=conn.execute("select SpecialID,SpecialName , ReadMe ,SpecialAbbreviation  from Special ")
	'if rsSpecialList.bof and rsSpecialList.eof then
	'	response.write ("请先添加课程")
	'else
		do while not rsSpecialList.eof
			if   (   InStr( AdminPurview_SpecialID  ,  rsSpecialList("SpecialID")   )   >0 ) then
				Set rsCourseList=Server.CreateObject("Adodb.RecordSet")
				'sql="select * from CourseList , Special , Admin where CourseList.SpecialID=Special.SpecialID and CourseList.TeacherName='" & session("AdminTrueName") &  "'"
				'sqlAdmin="select AdminPurview_SpecialID from Admin"
				sqlCourseList="select * from CourseList where  CourseList.SpecialID=" &      rsSpecialList("SpecialID") & " and TeacherName='" & session("AdminTeacherName") & "'"
				rsCourseList.Open sqlCourseList,conn,1,1
				 
			
  %>
  <!--<hr/>-->
  <tr class="title">
    <!-- 此行列出教师所开课程的列表的表头-->
    <td height="22"  align="center" width="100"><strong>课程名称</strong></td>
    <td  align="center"><strong>课程说明</strong></td>
    <td width="100" align="center"><strong>开课时间</strong></td>
    <td   width="200" align="center"><strong>上课班级</strong></td>
    <td width="100" align="center"><strong>操作</strong></td>
    <!--<td width="110" align="center"><strong>上课班级所属学院</strong></td>-->
    <!--<td width="100" height="22" align="center"><strong> 常规操作</strong></td>-->
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td align="center" rowspan="<%=(rsCourseList.RecordCount+2)%>"><!--显示课程名称,加第一行是为了表对齐，加第二行是为了新增开课表一行-->
      <a href="Admin_ArticleManageSpecial.asp?SpecialID=<%=rsSpecialList("SpecialID")%>" title="点击进入管理此课程的文章"><%=rsSpecialList("SpecialAbbreviation")%></a></td>
    <td align="center" rowspan="<%=(rsCourseList.RecordCount+2)%>"><!--显示课程简介，现在未使用-->
      <%=dvhtmlencode(rsSpecialList("ReadMe"))%></td>
    <!---->
  </tr>
  <!--<td colspan="2"><table>-->
  <% 'if rsCourseList.bof and rsCourseList.eof then
					'		response.write ("请先添加开课时间和班级列表")
					'else
         					'response.Write(" <table>  ") %>
  <% 			if  (rsCourseList.bof and rsCourseList.EOF) then
						%>
  <tr class="tdbg"  align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"  >
    <td align="center">未开授此课程</td>
    <td></td>
    <td></td>
  </tr>
  <%
					else
						do while not rsCourseList.EOF %>
  <!--一个课程的多个班-->
  <!--学生的班级的表示方式尚未定，因为不会做下拉列表框-->
  <!--学生的学院的表示方式尚未定，因为不会做下拉列表框-->
  <% 'Response.Write("<tr class=tdbg > <td align=center>") %>
  <tr class="tdbg" align="center" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td align="center"  ><!--显示开课时间，时间用四位数年份和星期数表示-->
      <%=rsCourseList("TermYear")%><%=rsCourseList("TermOrder")%></td>
    <td align='center' ><% 
	  Response.Write(rsCourseList("StudentClassName"))
	  Response.Write(rsCourseList("StudentClassYear"))
	  Response.Write(rsCourseList("StudentClassNumber"))
	  %>
      &nbsp;&nbsp;<%=rsCourseList("College")%>
      <% 'Response.Write(" </td> </tr>") %></td>
    <td align="center" ><a  href="Admin_Course.asp?Action=CourseDel&CourseListID=<%=rsCourseList("CourseListID")%>">删除此行</a></td>
  </tr>
  <%
						rsCourseList.MoveNext
  	 					loop
						'end if
						rsCourseList.Close
						set rsCourseList=Nothing
 		 				%>
  <!--下一个课程-->
  <!--添加开课表-->
  <% 'Response.Write("</table>") %>
  <!--</table></td>-->
  <!-- </tr>-->
  <!--<tr>-->
  <!--</tr>-->
  <% 				end if
   				 end if
		%>
  <tr  ></tr>
  <!--<hr/>-->
  <!--用于增加空白分隔<tr></tr> <tr></tr> <tr></tr> <tr></tr> <tr></tr>-->
  <%
		rsSpecialList.movenext
		loop
	'end if
	set rsSpecialList=nothing
  
  
  '添加开课表一项
  %>
</table>
<%
end sub
'结束教师管理员主体网页代码

'开始很难做得美观的新增开课表，这版式排不齐啊
sub CourseAdd()
dim rsCourseAdd , sqlCourseAdd , SpecialID , TermNumber , StudentGradeClass , TermYear , TermOrder ,StudentClassYear , StudentClassName , StudentClassNumber
if Trim(Request("SpecialID"))="" or Trim(Request("TermYear"))="" or Trim(Request("TermOrder"))="" or Trim(Request("College"))="" or Trim(Request("StudentClassName"))="" or Trim(Request("StudentClassYear"))="" or Trim(Request("StudentClassNumber"))=""   then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选齐全部项目！</li> <!--<a Href=Admin_Course.asp>返回</a>-->   "
		'Response.Write(ErrMsg)
'		Call Main()
		'rsCourseAdd.close
		'set rsCourseAdd=nothing
		'response.redirect "Admin_Course.asp"		
		exit sub
		end if
SpecialID=Clng(Trim(Request("SpecialID")))
StudentClassName=Trim(Request("StudentClassName")) 
StudentClassYear=Trim(Request("StudentClassYear"))
StudentClassNumber=Trim(Request("StudentClassNumber"))
TermYear=Clng(Trim(Request("TermYear")))
TermOrder=Trim(Request("TermOrder"))
StudentGradeClass=Trim(Request("StudentClassName")) & Trim(Request("StudentClassYear")) & Trim(Request("StudentClassNumber"))
sqlCourseAdd="Select * from CourseList where SpecialID=" & SpecialID & "and TermYear=" & TermYear & " and TermOrder='" & TermOrder & "' and StudentClassName='" & StudentClassName & "' and StudentClassYear='" & StudentClassYear & "' and StudentClassNumber='" & StudentClassNumber & "'"
	Set rsCourseAdd=Server.CreateObject("Adodb.RecordSet")
	rsCourseAdd.Open sqlCourseAdd,conn,1,3
	if not (rsCourseAdd.bof and rsCourseAdd.EOF) then
		FoundErr=True
		
		ErrMsg=ErrMsg & "<br><li>此班级在此学期已有此课！</li><!--<li> <a Href=Admin_Course.asp>返回</a> </li>-->"
		'Response.Write(ErrMsg)
'		Call Main()
		rsCourseAdd.close
		set rsCourseAdd=nothing
		'response.redirect "Admin_Course.asp"		
		exit sub
	else
		rsCourseAdd.addnew
		'共使用了八个字段
 	rsCourseAdd("SpecialID")=SpecialID
	rsCourseAdd("TermYear")=TermYear
	rsCourseAdd("TermOrder")=TermOrder
	if AdminPurview=2 then
		rsCourseAdd("TeacherName")=session("AdminTeacherName")
	elseif AdminPurview=1 then
		rsCourseAdd("TeacherName")=Trim(Request("TeacherName"))
	end if
	rsCourseAdd("College")=Trim(Request("College"))
	rsCourseAdd("StudentClassName")=StudentClassName
	rsCourseAdd("StudentClassYear")=StudentClassYear
	rsCourseAdd("StudentClassNumber")=StudentClassNumber
		
	rsCourseAdd.update
    rsCourseAdd.Close
	set rsCourseAdd=Nothing
	end if


response.redirect "Admin_Course.asp"
end sub
'结束新增开课表


'在开课表中删除一行
sub CourseDel()
	'dim rsCourseDel , sqlCourseDel
	'Set rsCourseDel=Server.CreateObject("Adodb.RecordSet")
'				sqlCourseDel="select CourseListID from CourseList where                CourseList.CourseListID=" & Trim(Request("CourseListID"))
'				rsCourseList.Open sqlCourseList,conn,1,1
'	
'				rsCourseDel.Close
'				set rsCourseDel=Nothing

	dim CourseListID
	CourseListID=Trim(Request("CourseListID"))
	if CourseListID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的开课编号CourseListID！</li>"
		exit sub
	else
		CourseListID=Clng(CourseListID)
	end if
	conn.Execute("delete from CourseList where CourseListID=" & CourseListID)
'	conn.execute("update Article set SpecialID=0 where SpecialID=" & SpecialID)
	call CloseConn()      
	response.redirect "Admin_Course.asp"
end sub
'结束删除开课表一行的代码
sub SaveAdd()
	dim SpecialName,ReadMe,SkinID,LayoutID,BrowsePurview,AddPurview,rs,mrs,MaxOrderID
	SpecialName=trim(request.Form("SpecialName"))
	ReadMe=trim(request("ReadMe"))
	SkinID=Trim(request("SkinID"))
	LayoutID=trim(request("LayoutID"))
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if SpecialName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程名称不能为空！</li>"
	end if
	
	''两课网站新增代码
	if SpecialName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程简称不能为空！</li>"
	end if

	'结束两课网站代码
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定栏目配色模板</li>"
	else
		SkinID=CLng(SkinID)
	end if
	if LayoutID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定版面设计模板</li>"
	else
		LayoutID=CLng(LayoutID)
	end if
	if FoundErr=True then
		exit sub
	end if
	set mrs=conn.execute("select max(OrderID) from Special")
	MaxOrderID=mrs(0)
	if isnull(MaxOrderID) then MaxOrderID=0
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select * From Special Where SpecialName='" & SpecialName & "'",conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程名称已经存在！</li>"
		rs.close
	    set rs=Nothing
    	exit sub
	end if
    
	'检查是否已有此简称
	
	set mrs=conn.execute("select max(OrderID) from Special")
	MaxOrderID=mrs(0)
	if isnull(MaxOrderID) then MaxOrderID=0
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select * From Special Where SpecialAbbreviation='" & Trim(Request("SpecialAbbreviation")) & "'",conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程名称已经存在！</li>"
		rs.close
	    set rs=Nothing
    	exit sub
	end if
	'结束检查
	
	
	
	
	rs.addnew
	rs("OrderID")=MaxOrderID+1
    rs("SpecialName")=SpecialName
	rs("ReadMe")=ReadMe
	rs("SkinID")=SkinID
	rs("LayoutID")=LayoutID
	rs("BrowsePurview")=BrowsePurview
	rs("AddPurview")=AddPurview
	
	'两课网站新增代码
	rs("SpecialAbbreviation")=Trim(Request("SpecialAbbreviation"))
	'结束代码
	rs.update
    rs.Close
    set rs=Nothing
	Response.Redirect "Admin_Special.asp"  
end sub

sub SaveModify()
	dim SpecialID,SpecialName,ReadMe,SkinID,LayoutID,BrowsePurview,AddPurview
	SpecialID=trim(request("SpecialID"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的课程ID！</li>"
		exit sub
	else
		SpecialID=Clng(SpecialID)
	end if
	SpecialName=trim(request.Form("SpecialName"))
	Readme=trim(request("Readme"))
	SkinID=Trim(request("SkinID"))
	LayoutID=trim(request("LayoutID"))
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if SpecialName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程名称不能为空！</li>"
	end if
	
		''两课网站新增代码
	if SpecialName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>课程简称不能为空！</li>"
	end if

	'结束两课网站代码

	
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定栏目配色模板</li>"
	else
		SkinID=CLng(SkinID)
	end if
	if LayoutID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定版面设计模板</li>"
	else
		LayoutID=CLng(LayoutID)
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * From Special Where SpecialID=" & SpecialID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,3
	if rs.bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的课程，可能已经被删除！</li>"
		rs.close
		set rs=nothing
	else
	    rs("SpecialName")=SpecialName
		rs("ReadMe")=ReadMe
		rs("SkinID")=SkinID
		rs("LayoutID")=LayoutID
		rs("BrowsePurview")=BrowsePurview
		rs("AddPurview")=AddPurview
			'两课网站新增代码
		rs("SpecialAbbreviation")=Trim(Request("SpecialAbbreviation"))
	'结束代码

		
		rs.update
		rs.close
		set rs=nothing
		call CloseConn()
		Response.Redirect "Admin_Special.asp"  
	end if
end sub

sub DelSpecial()
	dim SpecialID
	SpecialID=trim(request("SpecialID"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的课程ID！</li>"
		exit sub
	else
		SpecialID=Clng(SpecialID)
	end if
	conn.Execute("delete from Special where SpecialID=" & SpecialID)
	conn.execute("update Article set SpecialID=0 where SpecialID=" & SpecialID)
	call CloseConn()      
	response.redirect "Admin_Special.asp"

end sub

sub ClearSpecial()
	dim SpecialID
	SpecialID=trim(request("SpecialID"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的课程ID！</li>"
		exit sub
	else
		SpecialID=Clng(SpecialID)
	end if
	conn.execute("update Article set SpecialID=0 where SpecialID=" & SpecialID)
	call CloseConn()      
	response.redirect "Admin_Special.asp"
end sub

sub UpOrder()
	dim SpecialID,sqlOrder,rsOrder,MoveNum,cOrderID,tOrderID,i,rs
	SpecialID=trim(request("SpecialID"))
	cOrderID=Trim(request("cOrderID"))
	MoveNum=trim(request("MoveNum"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		SpecialID=CLng(SpecialID)
	end if
	if cOrderID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cOrderID=Cint(cOrderID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim mrs,MaxOrderID
	set mrs=conn.execute("select max(OrderID) From Special")
	MaxOrderID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
	
	'然后将位于当前栏目以上的栏目的OrderID依次加一，范围为要提升的数字
	sqlOrder="select * From Special where OrderID<" & cOrderID & " order by OrderID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tOrderID=rsOrder("OrderID")       '得到要提升位置的OrderID，包括子栏目
		conn.execute("update Special set OrderID=OrderID+1 where OrderID=" & tOrderID)
		i=i+1
		if i>MoveNum then
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
	call CloseConn()      
	response.redirect "Admin_Special.asp"
end sub

sub DownOrder()
	dim SpecialID,sqlOrder,rsOrder,MoveNum,cOrderID,tOrderID,i,rs,PrevID,NextID
	SpecialID=trim(request("SpecialID"))
	cOrderID=Trim(request("cOrderID"))
	MoveNum=trim(request("MoveNum"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		SpecialID=CLng(SpecialID)
	end if
	if cOrderID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cOrderID=Cint(cOrderID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim mrs,MaxOrderID
	set mrs=conn.execute("select max(OrderID) From Special")
	MaxOrderID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update Special set OrderID=" & MaxOrderID & " where SpecialID=" & SpecialID)
	
	'然后将位于当前栏目以下的栏目的OrderID依次减一，范围为要下降的数字
	sqlOrder="select * From Special where OrderID>" & cOrderID & " order by OrderID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tOrderID=rsOrder("OrderID")       '得到要提升位置的OrderID，包括子栏目
		conn.execute("update Special set OrderID=OrderID-1 where OrderID=" & tOrderID)
		i=i+1
		if i>MoveNum then
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update Special set OrderID=" & tOrderID & " where SpecialID=" & SpecialID)
	call CloseConn()      
	response.redirect "Admin_Special.asp"
end sub

sub ShowSpecial()
	dim rsSpecial
	set rsSpecial=conn.execute("select SpecialID,SpecialName from Special")
	if rsSpecial.bof and rsSpecial.eof then
		response.write "<option value=''>请先添加课程</option>"
	else
		do while not rsSpecial.eof
			response.write "<option value='" & rsSpecial(0) & "'>" & rsSpecial(1) & "</option>"
			rsSpecial.movenext
		loop
	end if
	set rsSpecial=nothing
end sub

sub UniteSpecial()
	dim SpecialID,TargetSpecialID,SuccessMsg
	SpecialID=trim(request("SpecialID"))
	TargetSpecialID=trim(request("TargetSpecialID"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要合并的课程！</li>"
	else
		SpecialID=CLng(SpecialID)
	end if
	if TargetSpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标课程！</li>"
	else
		TargetSpecialID=CLng(TargetSpecialID)
	end if
	if SpecialID=TargetSpecialID then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请不要在相同课程内进行操作</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	if FoundErr=True then
		exit sub
	end if
	
	'更新文章所属课程
	conn.execute("update Article set SpecialID="&TargetSpecialID&" where SpecialID="&SpecialID)
	'删除被合并课程及其下属课程
	conn.execute("delete from Special where SpecialID="&SpecialID)
		
	SuccessMsg="课程合并成功！已经将被合并课程的所有数据转入目标课程中。"
	call WriteSuccessMsg(SuccessMsg)
end sub

%>
