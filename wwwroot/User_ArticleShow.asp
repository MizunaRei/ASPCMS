<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if

dim ArticleID,sql,rs,FoundErr,ErrMsg,PurviewChecked
dim ClassID,tClass,ClassName,RootID,ParentID,ParentPath,Depth,ClassMaster
ArticleID=trim(request("ArticleID"))
FoundErr=False
PurviewChecked=False

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
sql="select * from article where Deleted=False and ArticleID=" & ArticleID
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>找不到文章</li>"
else
	ClassID=rs("ClassID")
	set tClass=conn.execute("select ClassName,RootID,ParentID,Depth,ParentPath,ClassMaster From ArticleClass where ClassID=" & ClassID)
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
	end if
	set tClass=nothing
end if
if FoundErr=True then
	rs.close
	set rs=nothing
	exit sub
end if

if rs("Editor")=UserName and rs("Passed")=False then
	PurviewChecked=True
else
	PurviewChecked=False
end if
%>
<html>
<head>
<title><%=rs("Title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>

<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" class="border" style="word-break:break-all">
  <tr class="title"> 
    <td>
<%
response.write "您现在的位置：&nbsp;<a href='User_ArticleManage.asp'>文章管理</a>&nbsp;&gt;&gt;&nbsp;"
if ParentID>0 then
	dim sqlPath,rsPath
	sqlPath="select ClassID,ClassName From ArticleClass where ClassID in (" & ParentPath & ") order by Depth"
	set rsPath=server.createobject("adodb.recordset")
	rsPath.open sqlPath,conn,1,1
	do while not rsPath.eof
		response.Write "<a href='User_ArticleManage.asp?ClassID=" & rsPath(0) & "'>" & rsPath(1) & "</a>&nbsp;&gt;&gt;&nbsp;"
		rsPath.movenext
	loop
	rsPath.close
	set rsPath=nothing
end if
response.write "<a href='User_ArticleManage.asp?ClassID=" & ClassID & "'>" & ClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
response.write "<a href='User_ArticleShow.asp?ArticleID=" & rs("ArticleID") & "'>" & rs("Title") & "</a>"
%>	
	</td>
    <td width="100" align="right">
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
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;点击数：" & rs("Hits") & "&nbsp;&nbsp;&nbsp;&nbsp;文章录入：" & rs("Editor") & ""
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
  <tr class="tdbg"> 
    <td colspan="2"> <li>上一篇文章： 
        <%
	  dim rsPrev
	  sql="Select Top 1 ArticleID,Title From Article Where Editor='" & UserName & "' and ArticleID<" & rs("ArticleID") & " order by ArticleID desc"
	  Set rsPrev= Server.CreateObject("ADODB.Recordset")
	  rsPrev.open sql,conn,1,1
	  if rsPrev.Eof then
	  	response.write "没有了"
	  else
	  	response.write "<a href='User_ArticleShow.asp?ArticleID="&rsPrev("ArticleID")& "'>"&rsPrev("Title") & "</a>"
	  end if
	  rsPrev.close
	  set rsPrev=nothing
	  %>
      </li>
	  <br> <li>下一篇文章： 
        <%
	  dim rsNext
	  sql="Select Top 1 ArticleID,Title From Article Where Editor='" & UserName & "' and ArticleID>" & rs("ArticleID") & " order by ArticleID"
	  Set rsNext= Server.CreateObject("ADODB.Recordset")
	  rsNext.open sql,conn,1,1
	  if rsNext.Eof then
	  	response.write "没有了"
	  else
	  	response.write "<a href='User_ArticleShow.asp?ArticleID="&rsNext("ArticleID")& "'>"&rsNext("Title") & "</a>"
	  end if
	  rsNext.close
	  set rsNext=nothing
	  %>
      </li></td>
  </tr>
  <% if PurviewChecked=True then%>
  <tr align="right" class="tdbg"> 
    <td colspan="2"><strong>可用操作：</strong> <a href="User_ArticleModify.asp?ArticleID=<%=rs("ArticleID")%>">修改</a> 
      <a href="User_ArticleDel.asp?Action=Del&ArticleID=<%=rs("ArticleID")%>">删除</a> 
    </td>
  </tr>
  <% end if%>
</table>
</body>
</html>
<%
	rs.close
	set rs=nothing
end sub
%>