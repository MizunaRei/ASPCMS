<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
Const PurviewLevel=0    '操作权限
Const CheckChannelID=0    '所属频道，0为不检测所属频道
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
dim ChannelID,rs,sql
dim FoundErr,ErrMsg
dim AdminPurview_Channel

ChannelID=trim(request("ChannelID"))
if ChannelID="" then
	ChannelID=0
else
	ChannelID=Clng(ChannelID)
end if
%>
<html>
<head>
<title>查看管理权限</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>查 看 管 理 权 限</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_ShowPurview.asp">所有管理权限</a> | <a href="Admin_ShowPurview.asp?ChannelID=2">文章频道权限</a> 
    </td>
  </tr>
</table>
<%
response.write "<br>您现在的位置：查看管理权限&nbsp;&gt;&gt;&nbsp;<font color=red>"
select case ChannelID
	case 0
		response.write "所有管理权限"
	case 2
		response.write "文章频道权限"
	case 3
		response.write "软件频道权限"
	case 4
		response.write "图片频道权限"
	case else
		response.write "错误的参数"
end select
response.write "</font><br>"
if ChannelID=0 then
	call ShowAllPurview()
else
	call ShowChannelPurview()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()
%>
</body>
</html>

<%
sub ShowAllPurview()
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td colspan="2"><strong>文章频道</strong> 
      <%
		if AdminPurview_Article=1 then response.write "（频道管理员）"
		if AdminPurview_Article=2 then response.write "（栏目总编）"
		if AdminPurview_Article=3 then response.write "（栏目管理员）"
		if AdminPurview_Article=4 then response.write "（无权限）"
		%>
    </td>
    <td height="22" colspan="2"><strong>下载频道</strong> 
      <%
		if AdminPurview_Soft=1 then response.write "（频道管理员）"
		if AdminPurview_Soft=2 then response.write "（栏目总编）"
		if AdminPurview_Soft=3 then response.write "（栏目管理员）"
		if AdminPurview_Soft=4 then response.write "（无权限）"
		%>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以添加、管理栏目和专题</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以添加、管理栏目</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理文章评论及文章回收站</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理软件评论</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理专题文章</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article<=2 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理软件回收站</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">各栏目文章录入、审核、管理权限</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Article<=2 then
			response.write "<font color=blue>全部权限</font>"
		elseif AdminPurview_Article=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=2'><font color=blue>部分权限</font></a>"
		else
			response.write "<font color=red>无权限</font>"
		end if
	  %>
    </td>
    <td width="30%">各栏目软件录入、审核、管理权限</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Soft<=2 then
			response.write "<font color=blue>全部权限</font>"
		elseif AdminPurview_Soft=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=3'><font color=blue>部分权限</font></a>"
		else
			response.write "<font color=red>无权限</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="title"> 
    <td colspan="2"><strong>图片频道</strong> 
      <%
		if AdminPurview_Photo=1 then response.write "（频道管理员）"
		if AdminPurview_Photo=2 then response.write "（栏目总编）"
		if AdminPurview_Photo=3 then response.write "（栏目管理员）"
		if AdminPurview_Photo=4 then response.write "（无权限）"
		%>
    </td>
    <td height="22" colspan="2"><strong>留言频道</strong></td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以添加、管理栏目</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">回复留言 </td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Reply")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理软件评论</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">修改留言 </td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Modify")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理软件回收站</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">删除留言</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Del")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">各栏目图片录入、审核、管理权限</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Photo<=2 then
			response.write "<font color=blue>全部权限</font>"
		elseif AdminPurview_Photo=3 then
			response.write "<a href='Admin_ShowPurview.asp?ChannelID=4'><font color=blue>部分权限</font></a>"
		else
			response.write "<font color=red>无权限</font>"
		end if
	  %>
    </td>
    <td width="30%">审核留言</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Guest,"Check")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="title"> 
    <td colspan="4" height="22"><strong>网站管理权限</strong><strong> </strong></td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以修改自己的登录密码</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"ModifyPwd")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以修改网站信息配置 <br>
    </td>
    <td align="center" width="20%"> 
      <%if AdminPurview=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以进行频道管理</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Channel")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理网站广告</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"AD")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理网站公告</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Announce")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理友情链接</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"FriendSite")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理网站调查</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Vote")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理网站统计</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Counter")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理注册用户</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"User")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理邮件列表</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"MailList")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理配色模板</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Skin")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理版面设计模板</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"Layout")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理JS代码</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"JS")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">可以管理上传文件</td>
    <td align="center" width="20%"> 
      <%if CheckPurview(AdminPurview_Others,"UpFile")=True then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理数据库</td>
    <td align="center" width="20%"> 
      <%if AdminPurview=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
    <td width="30%">&nbsp;</td>
    <td align="center" width="20%">&nbsp; </td>
  </tr>
</table>
<%
end sub


Sub ShowChannelPurview()
	dim AdminChannel_Name
	select case ChannelID
	case 2
		AdminPurview_Channel=AdminPurview_Article
		AdminChannel_Name="文章"
	case 3
		AdminPurview_Channel=AdminPurview_Soft
		AdminChannel_Name="软件"
	case 4
		AdminPurview_Channel=AdminPurview_Photo
		AdminChannel_Name="图片"
	end select

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td colspan="2" height="22"><strong><%=AdminChannel_Name%>频道</strong> 
      <%
		if AdminPurview_Channel=1 then response.write "（频道管理员）"
		if AdminPurview_Channel=2 then response.write "（栏目总编）"
		if AdminPurview_Channel=3 then response.write "（栏目管理员）"
		if AdminPurview_Channel=4 then response.write "（无权限）"
		%>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以添加、管理栏目</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理<%=AdminChannel_Name%>评论</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理<%=AdminChannel_Name%>回收站</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel=1 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <%if  ChannelID=2 then%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">可以管理专题文章</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel<=2 then
			response.write "<font color=blue>√</font>"
		else
			response.write "<font color=red>×</font>"
		end if
	  %>
    </td>
  </tr>
  <%end if%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30%">各栏目<%=AdminChannel_Name%>录入、审核、管理权限</td>
    <td align="center" width="20%"> 
      <%if AdminPurview_Channel<=2 then
			response.write "<font color=blue>全部权限</font>"
		elseif AdminPurview_Channel=3 then
			response.write "<font color=blue>部分权限</font>"
		else
			response.write "<font color=red>无权限</font>"
		end if
	  %>
    </td>
  </tr>
</table>
<br>
<% 
	if AdminPurview_Channel=3 then
		dim arrShowLine(10)
		for i=0 to ubound(arrShowLine)
			arrShowLine(i)=False
		next
		dim sqlClass,rsClass,i,iDepth
		select case ChannelID
		case 2
			sqlClass="select * From ArticleClass order by RootID,OrderID"
		case 3
			sqlClass="select * From SoftClass order by RootID,OrderID"
		case 4
			sqlClass="select * From PhotoClass order by RootID,OrderID"
		end select
		set rsClass=server.CreateObject("adodb.recordset")
		rsClass.open sqlClass,conn,1,1
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="2" class="border">
	  <tr align="center" class="title">
		<td height="22"><strong>栏目名称</strong></td>
		<td width="100" height="22"><strong>录入</strong></td>
		<td width="100"><strong>审核</strong></td>
		<td width="100" height="22"><strong>管理</strong></td>
	  </tr>
		<% 
		do while not rsClass.eof 
		%>
	  <tr class="tdbg">
		<td><% 
		iDepth=rsClass("Depth")
		if rsClass("NextID")>0 then
			arrShowLine(iDepth)=True
		else
			arrShowLine(iDepth)=False
		end if
		if iDepth>0 then
			for i=1 to iDepth 
				if i=iDepth then 
					if rsClass("NextID")>0 then 
						response.write "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>" 
					else 
						response.write "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>" 
					end if 
				else 
					if arrShowLine(i)=True then
						response.write "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>" 
					else
						response.write "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>" 
					end if
				end if 
			next 
		  end if 
		  if rsClass("Child")>0 then 
			response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
		  else 
			response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
		  end if 
		  if rsClass("Depth")=0 then 
			response.write "<b>" 
		  end if 
		  response.write rsClass("ClassName")
		  %>
		</td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassInputer"),AdminName)=True then
				response.write "<font color=blue>√</font>"
			else
				response.write "<font color=red>×</font>"
			end if
		end if
		%>
		</td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassChecker"),AdminName)=True then
				response.write "<font color=blue>√</font>"
			else
				response.write "<font color=red>×</font>"
			end if
		end if
		%></td>
		<td align="center"><%
		if AdminPurview_Channel=3 then
			if CheckClassMaster(rsClass("ClassMaster"),AdminName)=True then
				response.write "<font color=blue>√</font>"
			else
				response.write "<font color=red>×</font>"
			end if
		end if
		%></td>
	  </tr>
		<% 
		rsClass.movenext 
		loop 
		rsClass.close
		set rsClass=nothing
		%>
	</table>
	<%
	end if
end sub
%>
