<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
Const CheckChannelID=2    '所属频道，0为不检测所属频道
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
dim strFileName,ComeUrl
strFileName="Admin_Comment.asp"
ComeUrl=trim(request("ComeUrl"))
if ComeUrl="" then
	ComeUrl=Request.ServerVariables("HTTP_REFERER")
end if
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages,i,j
dim Title
dim Action,FoundErr,Errmsg
Action=Trim(request("Action"))
Title=Trim(request("Title"))
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
dim sql,rs
%>
<html>
<head>
<title>文章评论管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
   if(confirm("确定要删除选中的评论吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><b>文 章 评 论 管 理</b></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30">
      <table width="100%"><tr>
	  <form name="searchsoft" method="get" action="Admin_Comment.asp">
            <td> 
              <input name="Title" type="text" class=smallInput id="Title" size="28">
<input name="Query" type="submit" id="Query" value="查 询">
              &nbsp;&nbsp;请输入评论的标题。如果为空，则查找所有评论。</td>
          </form></tr></table>
	</td>
  </tr>
</table>
<%
if Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelComment()
elseif Action="Del2" then
	call DelComment2()
elseif Action="Reply" then
	call Reply()
elseif Action="SaveReply" then
	call SaveReply()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select A.ArticleID, A.ClassID, A.Title as ArticleTitle, A.IncludePic, C.CommentID,C.UserName,C.IP,C.Title as CommentTitle, C.Content,C.WriteTime,C.ReplyName,C.ReplyContent,C.ReplyTime from Comment C Left Join Article A On C.ArticleID=A.ArticleID "
	if Title<>"" then
		sql=sql & " where C.Title like '%" & Title & "%' "
	end if
	sql=sql & " order by A.ArticleID desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<form name="del" method="Post" action="Admin_Comment.asp" onsubmit="return ConfirmDel();">
  <tr>
    <td align="center">
<table border="0" cellpadding="2" width="100%" cellspacing="0">
  <tr>
            <td>您现在的位置：<a href="Admin_Comment.asp">评论管理</a>&nbsp;&gt;&gt;&nbsp; 
              <%
if request.querystring="" then
	response.write "所有评论"
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "标题中含有“<font color=blue>" & Title & "</font>”的评论"
		else
			response.Write("所有评论")
		end if
	end if
end if
%>
	</td>
	<td width="150" align="right">
<%
  	if rs.eof and rs.bof then
		response.write "共找到 0 篇评论</td></tr></table>"
	else
    	totalPut=rs.recordcount
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
		response.Write "共找到 " & totalPut & " 篇评论"
%>
	</td>
  </tr>
</table>
<%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"篇评论"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"篇评论"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"篇评论"
	    	end if
		end if
	end if
%>
      </form>
	</td>
  </tr>
</table>
<%  
   	rs.close
   	set rs=nothing  
end sub

sub showContent
   	dim i
    i=0
	dim PrevID
	PrevID=rs("ArticleID")
	do while not rs.eof
		if rs("ArticleID")<>PrevID then response.write "</table></td></tr></table><br>"
		if i=0 or rs("ArticleID")<>PrevID then
%>
        <table class="border" width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr class="title"> 
            <td width="80%" height="22">| 
              <%
			if rs("IncludePic")=true then
			  	response.write "<font color=blue>[图文]</font>"
			end if
			response.write "<a href='Admin_ArticleShow.asp?ArticleID=" & rs("ArticleID") & "'>" & rs("ArticleTitle") & "</a>"
			%></td>
            <td width="20%" align="right"><a href="Admin_Comment.asp?ArticleID=<%=rs("ArticleID")%>&Action=Del2">删除此文章下的所有评论</a></td>
          </tr>
          <tr> 
            <td colspan="2" align="right"> <table border="0" cellspacing="1" width="96%" cellpadding="0" style="word-break:break-all">
        <%end if%>
        <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
          <td width="30" align="center"><input name='CommentID' type='checkbox' onclick="unselectall()" id="CommentID" value='<%=cstr(rs("CommentID"))%>'></td>
          <td><a href="#" Title="<%=left(rs("Content"),200)%>"><%=rs("CommentTitle")%></a></td>
          <td width="80" align="center"><%=rs("UserName") %></td>
          <td width="120" align="center"><%=rs("IP")%></td>
          <td width="120" align="center"><%= rs("WriteTime") %></td>
          <td width="100" align="center"><%
		  if rs("ReplyName")<>"" then
		      response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
		  else
			  response.write "<a href='Admin_Comment.asp?Action=Reply&CommentID=" & rs("Commentid") & "'>回复</a>&nbsp;&nbsp;"
		  end if
		  response.write "<a href='Admin_Comment.asp?Action=Modify&CommentID=" & rs("Commentid") & "'>修改</a>&nbsp;&nbsp;" 
          response.write "<a href='Admin_Comment.asp?Action=Del&CommentID=" & rs("CommentID") & "' onclick='return ConfirmDel();'>删除</a>"
		  %></td>
        </tr>
        <%if rs("ReplyName")<>"" then%>
		<tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
          <td align="center">&nbsp;</td>
          <td colspan="4"><%response.write "管理员【" & rs("ReplyName") & "】于 " & rs("ReplyTime") & " 回复：<br><div style='padding:0px 20px'>" & rs("ReplyContent") & "</div>"%></td>
          <td align="center"><a href="Admin_Comment.asp?Action=Reply&CommentID=<%=rs("CommentID")%>">修改回复内容</a></td>
		</tr>
        <%
		end if
		i=i+1
	    if i>=MaxPerPage then exit do
	    PrevID=rs("ArticleID")
		rs.movenext
	loop
%>
      </table></td>
          </tr>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有评论 </td>
    <td><input name="submit" type='submit' value='删除选定的评论'>
              <input name="Action" type="hidden" id="Action" value="Del"></td>
  </tr>
</table>
<%
end sub 

sub Modify()
	dim CommentID
	CommentID=trim(Request("CommentID"))
	if CommentID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	else
		CommentID=Clng(CommentID)
	end if
	sql="Select * from Comment where CommentID=" & CommentID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
	if rs.Bof or rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的评论！</li>"
	else
	
%>
<form method="post" action="Admin_Comment.asp" name="form1">
    <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
      <tr align="center" class="title"> 
        
      <td height="22" colspan="2"> <strong>修 改 评 论</strong></td>
      </tr>
      <tr> 
        <td width="200" align="right" class="tdbg">评论人用户名：</td>
        <td class="tdbg"><input name="UserName" type="text" id="UserName" value="<%=rs("UserName")%>"></td>
      </tr>
      <tr> 
        <td width="200" align="right" class="tdbg">评论人IP：</td>
        <td class="tdbg"> <input name="IP" type="text" id="IP" value="<%=rs("IP")%>"></td>
      </tr>
      <tr> 
        <td width="200" align="right" class="tdbg">评论时间：</td>
        <td class="tdbg"> <input name="WriteTime" type="text" id="WriteTime" value="<%=rs("WriteTime")%>"></td>
      </tr>
      <tr> 
        <td width="200" align="right" class="tdbg">评论标题：</td>
        <td class="tdbg"><input name="Title" type="text" id="Title" value="<%=rs("Title")%>"></td>
      </tr>
      <tr> 
        <td width="200" align="right" class="tdbg">评论内容：</td>
        <td class="tdbg"> <textarea name="Content" cols="50" rows="10" id="Content"><%=rs("Content")%></textarea></td>
      </tr>
      <tr align="center"> 
        <td height="30" colspan="2" class="tdbg"><input name="ComeUrl" type="hidden" id="ComeUrl" value="<%=ComeUrl%>">
        <input name="Action" type="hidden" id="Action" value="SaveModify"> 
          <input name="CommentID" type="hidden" id="CommentID" value="<%=rs("CommentID")%>"> 
          <input  type="submit" name="Submit" value=" 保存修改结果 ">
        </td>
      </tr>
    </table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub Reply()
	dim CommentID
	CommentID=trim(Request("CommentID"))
	if CommentID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	else
		CommentID=Clng(CommentID)
	end if
	sql="select A.ArticleID, A.ClassID, A.Title as ArticleTitle, A.IncludePic, C.CommentID,C.UserName,C.IP,C.Title as CommentTitle, C.Content,C.WriteTime,C.ReplyContent from Comment C Left Join Article A On C.ArticleID=A.ArticleID where C.CommentID=" & CommentID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
	if rs.Bof or rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的评论！</li>"
	else
	
%>
<form method="post" action="Admin_Comment.asp" name="form1">
    
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr align="center" class="title"> 
      <td height="22" colspan="2"> <strong>回 复 评 论</strong></td>
    </tr>
    <tr> 
      <td width="200" align="right" class="tdbg">评论文章标题：</td>
      <td class="tdbg"><%=rs("ArticleTitle")%></td>
    </tr>
    <tr> 
      <td width="200" align="right" class="tdbg">评论人用户名：</td>
      <td class="tdbg"><%=rs("UserName")%></td>
    </tr>
    <tr> 
      <td width="200" align="right" class="tdbg">评论标题：</td>
      <td class="tdbg"><%=rs("CommentTitle")%></td>
    </tr>
    <tr> 
      <td width="200" align="right" class="tdbg">评论内容：</td>
      <td class="tdbg"><%=rs("Content")%></td>
    </tr>
    <tr>
      <td align="right" class="tdbg">回复内容：</td>
      <td class="tdbg"><textarea name="ReplyContent" cols="50" rows="6" id="ReplyContent"><%=rs("ReplyContent")%></textarea></td>
    </tr>
    <tr align="center"> 
      <td height="30" colspan="2" class="tdbg"><input name="ComeUrl" type="hidden" id="ComeUrl" value="<%=ComeUrl%>">
	  <input name="Action" type="hidden" id="Action" value="SaveReply"> 
        <input name="CommentID" type="hidden" id="CommentID" value="<%=rs("CommentID")%>"> 
        <input  type="submit" name="Submit" value=" 回 复 "> </td>
    </tr>
  </table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub
%>
</body>
</html>
<%
sub DelComment()
	dim CommentID,i
	CommentID=trim(Request("CommentID"))
	if CommentID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	end if

	if instr(CommentID,",")>0 then
		dim idarr
		idArr=split(CommentID)
		for i = 0 to ubound(idArr)
		    conn.execute "delete from Comment where Commentid=" & Clng(idArr(i))
		next
	else
		conn.execute "delete from Comment where Commentid=" & Clng(CommentID)
	end if
	call CloseConn()
	response.redirect ComeUrl
End sub

sub DelComment2()
    dim ArticleID
	ArticleID=trim(request("ArticleID"))
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	else
		ArticleID=Clng(ArticleID)
	end if
	
    conn.execute "delete from Comment where ArticleID=" & ArticleID
	call CloseConn()
	response.redirect ComeUrl
End sub

sub SaveModify()
	dim CommentID,UserName,IP,Title,Content,WriteTime
	CommentID=trim(Request("CommentID"))
	UserName=trim(request.form("UserName"))
	IP=trim(request.form("IP"))
	Title=trim(Request.form("Title"))
	Content=trim(Request.form("Content"))
	WriteTime=trim(request.form("WriteTime"))
	if CommentID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	else
		CommentID=Clng(CommentID)
	end if
	if UserName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入评论人用户名</li>"
	end if
	if IP="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入评论人IP</li>"
	end if
	if WriteTime="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入发表评论时间</li>"
	end if
	if Title="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入评论标题</li>"
	end if
	if Content="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入评论内容</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	sql="Select * from Comment where CommentID=" & CommentID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof or rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的评论！</li>"
	else
		rs("UserName")=UserName
		rs("IP")=IP
		rs("WriteTime")=WriteTime
		rs("Title")=dvhtmlencode(Title)
     	rs("Content")=dvhtmlencode(Content)
     	rs.update
	end if
	rs.Close
   	set rs=Nothing
	call CloseConn()
	response.redirect ComeUrl
end sub

sub SaveReply()
	dim CommentID,ReplyName,ReplyContent,ReplyTime
	CommentID=trim(Request("CommentID"))
	ReplyContent=trim(request("ReplyContent"))
	if CommentID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定评论ID</li>"
		Exit sub
	else
		CommentID=Clng(CommentID)
	end if
	if ReplyContent="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请输入回复内容</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from Comment where CommentID=" & CommentID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof or rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的评论！</li>"
	else
		rs("ReplyName")=AdminName
		rs("ReplyTime")=now()
     	rs("ReplyContent")=dvhtmlencode(ReplyContent)
     	rs.update
	end if
	rs.Close
   	set rs=Nothing
	call CloseConn()
	response.redirect ComeUrl
end sub

%>