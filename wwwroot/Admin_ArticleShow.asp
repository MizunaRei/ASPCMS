<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<SCRIPT language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除选中的文章吗？本操作将把选中的文章移到回收站中。必要时您可从回收站中恢复！"))
     return true;
   else
     return false;
}
</SCRIPT>
<%
dim ArticleID,sql,rs,FoundErr,ErrMsg,PurviewChecked,PurviewChecked2
dim ClassID,tClass,ClassName,RootID,ParentID,ParentPath,Depth,ClassMaster,ClassChecker
ArticleID=trim(request("ArticleID"))
FoundErr=False
PurviewChecked=False
PurviewChecked2=False

call main()
if FoundErr=True then
	WriteErrMsg()
end if
call CloseConn()

sub main()
if ArticleId="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足</li>"
	exit sub
else
	ArticleID=Clng(ArticleID)
end if
sql="select * from article where Deleted=False and ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到文章</li>"
else
	ClassID=rs("ClassID")
	set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster,ClassChecker From ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		founderr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
	else
		ClassName=tClass(0)
		RootID=tClass(1)
		ParentID=tClass(2)
		Depth=tClass(3)
		ParentPath=tClass(4)
		ClassMaster=tClass(5)
		ClassChecker=tClass(6)
	end if
	set tClass=nothing
end if
if FoundErr=True then
	rs.close
	set rs=nothing
	exit sub
end if
if AdminPurview=1 or AdminPurview_Article<=2 then
	PurviewChecked=True
else
	PurviewChecked=CheckClassMaster(ClassMaster,AdminName)
	if PurviewChecked=False and ParentID>0 then
		set tClass=conn.execute("select ClassMaster from ArticleClass where ClassID in (" & ParentPath & ")")
		do while not tClass.eof
			PurviewChecked=CheckClassMaster(tClass(0),AdminName)
			if PurviewChecked=True then exit do
			tClass.movenext
		loop
	end if
	PurviewChecked2=CheckClassMaster(ClassChecker,AdminName)
	if PurviewChecked2=False and ParentID>0 then
		set tClass=conn.execute("select ClassMaster from ArticleClass where ClassID in (" & ParentPath & ")")
		do while not tClass.eof
			PurviewChecked2=CheckClassMaster(tClass(0),AdminName)
			if PurviewChecked2=True then exit do
			tClass.movenext
		loop
	end if
end if
%>
<html>
<head>
<title><%=rs("Title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="border">
  <tr class="title"> 
    <td height="22">
<%
response.write "您现在的位置：&nbsp;<a href='Admin_ArticleManage.asp'>文章管理</a>&nbsp;&gt;&gt;&nbsp;"
if ParentID>0 then
	dim sqlPath,rsPath
	sqlPath="select ClassID,ClassName From ArticleClass where ClassID in (" & ParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		response.Write "<a href='Admin_ArticleManage.asp?ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
end if
response.write "<a href='Admin_ArticleManage.asp?ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
response.write "<a href='Admin_ArticleShow.asp?ArticleID=" & rs("ArticleID") & "'>" & rs("Title") & "</a>"
%>	
	</td>
    <td width="100" height="22" align="right">
<%
if rs("OnTop")=true then
	response.Write("<font color=blue>顶</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Hits")>=HitsOfHot then
	response.write("<font color=red>热</font>&nbsp;")
else
	response.write("&nbsp;&nbsp;&nbsp;")
end if
if rs("Elite")=true then
	response.write("<font color=green>荐</font>")
else
	response.write("&nbsp;&nbsp;")
end if
%>
    </td>
  </tr>
  <tr align="center" class="tdbg"> 
    <td height="40" colspan="2" valign="bottom"><font size="5"><%=rs("Title")%></font></td>
  </tr>
  <tr align="center" class="tdbg">
    <td colspan="2">
        <%
		dim Author,CopyFrom
		Author=rs("Author")
		CopyFrom=rs("CopyFrom")
		response.write "作者："
		if instr(Author,"|")>0 then
			response.write "<a href='mailto:" & right(Author,len(Author)-instr(Author,"|")) & "'>" & left(Author,instr(Author,"|")-1) & "</a>"
		else
			response.write Author
		end if
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;转贴自："
		if instr(CopyFrom,"|")>0 then
			response.write "<a href='" & right(CopyFrom,len(CopyFrom)-instr(CopyFrom,"|")) & "'>" & left(CopyFrom,instr(CopyFrom,"|")-1) & "</a>"
		else
			response.write CopyFrom
		end if
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;点击数：" & rs("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;文章录入：<a href='Admin_ArticleManage.asp?Field=Editor&Keyword=" & rs("Editor") & "'>" & rs("Editor") & "</a>"
		%>
    </td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><p><%=rs("Content")%></p></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="tdbg"><td colspan="2">
      <li>上一篇文章： 
        <%
	  dim rsPrev
	  sql="Select Top 1 ArticleID,Title From Article Where Deleted=False and ArticleID<" & rs("ArticleID") & " order by ArticleID desc"
	  Set rsPrev= Server.CreateObject("ADODB.Recordset")
	  rsPrev.open sql,conn,1,1
	  if rsPrev.Eof then
	  	response.write "没有了"
	  else
	  	response.write "<a href='Admin_ArticleShow.asp?ArticleID="&rsPrev("ArticleID")& "'>"&rsPrev("Title") & "</a>"
	  end if
	  rsPrev.close
	  set rsPrev=nothing
	  %>
      </li>
      <br> <li> 下一篇文章： 
        <%
	  dim rsNext
	  sql="Select Top 1 ArticleID,Title From Article Where Deleted=False and ArticleID>" & rs("ArticleID") & " order by ArticleID asc"
	  Set rsNext= Server.CreateObject("ADODB.Recordset")
	  rsNext.open sql,conn,1,1
	  if rsNext.Eof then
	  	response.write "没有了"
	  else
	  	response.write "<a href='Admin_ArticleShow.asp?ArticleID="&rsNext("ArticleID")& "'>"&rsNext("Title") & "</a>"
	  end if
	  rsNext.close
	  set rsNext=nothing
	  %>
      </li></td>
  </tr>
  <tr align="right" class="tdbg"> 
    <td height="21" colspan="2">
<%
response.write "<strong>可用操作：</strong>"
if (rs("Editor")=AdminName and rs("Passed")=False) or PurviewChecked=True then
	response.write "<a href='Admin_ArticleModify.asp?ArticleID=" & rs("ArticleID") & "'>修改</a>&nbsp;&nbsp;"
    response.write "<a href='Admin_ArticleDel.asp?Action=Del&ArticleID=" & rs("ArticleID") & "' onclick='return ConfirmDel();'>删除</a>&nbsp;&nbsp;" 
end if
if AdminPurview=1 or AdminPurview_Article<=2 then
	response.write "<a href='Admin_ArticleMove.asp?ArticleID=" & rs("ArticleID") & "'>移动</a>&nbsp;&nbsp;"
end if
if PurviewChecked2=True then
	if rs("Passed")=false then
		response.write "<a href='Admin_ArticleProperty.asp?Action=SetPassed&ArticleID=" & rs("ArticleID") & "'>通过审核</a>&nbsp;&nbsp;"
	Else
  		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelPassed&ArticleID=" & rs("ArticleID") & "'>取消审核</a>&nbsp;&nbsp;"
  	end if
end if
if PurviewChecked=True then
  	if rs("OnTop")=false then
   		response.write "<a href='Admin_ArticleProperty.asp?Action=SetOnTop&ArticleID=" & rs("ArticleID") & "'>固顶</a>&nbsp;&nbsp;"
   	else
		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelOnTop&ArticleID=" & rs("ArticleID") & "'>解固</a>&nbsp;&nbsp;" 
   	end if
  	if rs("Elite")=false then
   		response.write "<a href='Admin_ArticleProperty.asp?Action=SetElite&ArticleID=" & rs("ArticleID") & "'>设为推荐</a>"
   	else
		response.write "<a href='Admin_ArticleProperty.asp?Action=CancelElite&ArticleID=" & rs("ArticleID") & "'>取消推荐</a>"
    end if
end if
%></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"><tr>
    <td>
      <strong>相关评论：</strong><br>
<%
dim rsComment
sql="select * from ArticleComment where ArticleID=" & ArticleID
Set rsComment= Server.CreateObject("ADODB.Recordset")
rsComment.open sql,conn,1,1
if rsComment.eof then
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;暂时没有任何人对本文章发表评论"
else
%>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="border" style="word-break:break-all">
        <tr align="center" class="title"> 
          <td width="30" height="22"><strong>ID</strong></td>
          <td height="22"><strong>内容</strong></td>
          <td width="60" height="22"><strong>评论人</strong></td>
          <td width="120" height="22"><strong>评论人IP</strong></td>
          <td width="120" height="22"><strong>评论时间</strong></td>
          <td width="100" height="22"><strong>操作</strong></td>
        </tr>
<%
	do while not rsComment.eof
%>
        <tr class="tdbg"> 
          <td width="30" align="center"><%= rsComment("CommentID") %></td>
          <td><% response.write "<a href=# title='" & rsComment("Content") & "'>" & left(rsComment("Content"),25) & "</a>" %></td>
          <td width="60" align="center"><%= rsComment("UserName") %></td>
          <td width="120" align="center"><%=rsComment("IP")%></td>
          <td width="120" align="center"><%= rsComment("WriteTime") %></td>
          <td width="100" align="center">
		  <%
		  if AdminPurview=1 or AdminPurview_Article=1 then
			  if rsComment("ReplyName")<>"" then
				  response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
			  else
				  response.write "<a href='Admin_ArticleComment.asp?Action=Reply&CommentID=" & rsComment("Commentid") & "'>回复</a>&nbsp;&nbsp;"
			  end if
			  response.write "<a href='Admin_ArticleComment.asp?Action=Modify&CommentID=" & rsComment("Commentid") & "'>修改</a>&nbsp;&nbsp;"
			  response.write "<a href='Admin_ArticleComment.asp?Action=Del&CommentID=" & rsComment("CommentID") & "'>删除</a>"
		  end if%>
		  </td>
        </tr>
        <%if rsComment("ReplyName")<>"" then%>
		<tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
          <td align="center">&nbsp;</td>
          <td colspan="4"><%response.write "管理员【" & rsComment("ReplyName") & "】于 " & rsComment("ReplyTime") & " 回复：<br><div style='padding:0px 20px'>" & rsComment("ReplyContent") & "</div>"%></td>
          <td align="center"><a href="Admin_ArticleComment.asp?Action=Reply&CommentID=<%=rsComment("CommentID")%>">修改回复内容</a></td>
		</tr>
        <%
		end if
		rsComment.movenext
	loop
%>
      </table>
<%
end if
rsComment.close
set rsComment=nothing
%>
	</td>
  </tr>
</table>
</body>
</html>
<%
end sub
%>