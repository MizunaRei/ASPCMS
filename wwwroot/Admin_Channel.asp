<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
Const CheckChannelID=0
Const PurviewLevel_Others="Channel"
'response.write "此功能被WEBBOY暂时禁止了！"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
dim Action,FoundErr,ErrMsg
dim rs, sql,iCount,i,j
Action=trim(request("Action"))
%>
<html>
<head>
<title>频道管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>频 道 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td><a href="Admin_Channel.asp">频道管理首页</a> | <a href="Admin_Channel.asp?Action=Add">添加新频道</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddChannel()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelChannel()
elseif Action="UpOrder" then 
	call UpOrder() 
elseif Action="DownOrder" then 
	call DownOrder() 
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from Channel order by OrderID"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
	j=1
%>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" align="center"><strong> 频道名称</strong></td>
    <td width="120" align="center"><strong>频道说明</strong></td>
    <td width="150" align="center"><strong>链接地址</strong></td>
    <td width="80" align="center"><strong>频道类型</strong></td>
    <td width="80" height="22" align="center"><strong> 常规操作</strong></td>
    <td width="240" colspan="2" align="center"><strong>排序操作</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
    <td align="center"><%=rs("ChannelName")%></td>
    <td width="120"><%
	if rs("Readme")<>"" then
		response.write rs("Readme")
	else
		response.write "&nbsp;"
	end if
	%></td>
    <td width="150"><%=rs("LinkUrl")%></td>
    <td width="80" align="center"><%
	if rs("ChannelType")=1 then
		response.write "用户添加"
	else
		response.write "系统提供"
	end if
	%></td>
    <td width="80" align="center"><%
	response.write "<a href='Admin_Channel.asp?Action=Modify&ChannelID=" & rs("ChannelID") & "'>修改</a>&nbsp;&nbsp;"
	if rs("ChannelType")=1 then
		response.write "<a href='Admin_Channel.asp?Action=Del&ChannelID=" & rs("ChannelID") & "' onClick=""return confirm('确定要删除此频道吗？');"">删除</a>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	%></td>
    <form action='Admin_Channel.asp?Action=UpOrder' method='post'>
      <td width='120' align="center"> <% 
	if j>1 then 
		response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ChannelID value="&rs("ChannelID")&">"
		response.write "<input type=hidden name=cOrderID value="&rs("OrderID")&">&nbsp;<input type=submit name=Submit value=修改>" 
	else 
		response.write "&nbsp;" 
	end if 
%> </td>
    </form>
    <form action='Admin_Channel.asp?Action=DownOrder' method='post'>
      <td width='120' align="center"> <%
	if iCount>j then 
		response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ChannelID value="&rs("ChannelID")&">"
		response.write "<input type=hidden name=cOrderID value="&rs("OrderID")&">&nbsp;<input type=submit name=Submit value=修改>" 
	else 
		response.write "&nbsp;" 
	end if 
%> </td>
    </form>
  </tr>
  <%
     	j=j+1	
		rs.MoveNext
   	loop
  %>
</table>  
<%
	rs.Close
	set rs=Nothing
end sub

sub AddChannel()
%>
<form method="post" action="Admin_Channel.asp" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>添 加 新 频 道</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong> 频道名称：</strong></td>
      <td class="tdbg"><input name="ChannelName" type="text" id="ChannelName" size="49" maxlength="30"> 
        &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong>频道说明</strong>：</td>
      <td class="tdbg"><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong>链接地址：</strong><br>
        请输入相对于网站根目录的绝对路径，如根目录为“/”，本系统目录为“/article”。</td>
      <td class="tdbg"><input name="LinkUrl" type="text" id="LinkUrl" size="49" maxlength="200"> </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input  type="submit" name="Submit" value=" 添 加 "> &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Channel.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
end sub

sub Modify()
	dim ChannelID
	ChannelID=trim(request("ChannelID"))
	if ChannelID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改频道ID</li>"
		exit sub
	else
		ChannelID=Clng(ChannelID)
	end if
	set rs=conn.execute("select * from Channel where ChannelID=" & ChannelID)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的频道！</li>"
	else
%>
<form method="post" action="Admin_Channel.asp" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>修 改 频 道 信 息</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong> 频道名称：</strong></td>
      <td class="tdbg"><input name="ChannelName" type="text" id="ChannelName" value="<%=rs("ChannelName")%>" size="49" maxlength="30"> 
        &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong>频道说明</strong>：</td>
      <td class="tdbg"><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"><%=rs("ReadMe")%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="300" class="tdbg"><strong>链接地址：</strong><br>请输入相对于网站根目录的绝对路径，如根目录为“/”，本系统目录为“/article”。</td>
      <td class="tdbg"><input name="LinkUrl" type="text" id="LinkUrl" value="<%=rs("LinkUrl")%>" size="49" maxlength="200"> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="ChannelID" type="hidden" id="ChannelID" value="<%=rs("ChannelID")%>">
        <input name="Action" type="hidden" id="Action" value="SaveModify">
        <input name="Submit"  type="submit" id="Submit" value=" 修 改 "> &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Channel.asp'" style="cursor:hand;"></td>
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
sub SaveAdd()
	dim ChannelName,ReadMe,LinkUrl,mrs,MaxOrderID
	ChannelName=trim(request.Form("ChannelName"))
	ReadMe=trim(request("ReadMe"))
	LinkUrl=trim(request("LinkUrl"))
	if ChannelName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>频道名称不能为空！</li>"
	end if
	if LinkUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接地址不能为空！</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	set mrs=conn.execute("select max(OrderID) from Channel")
	MaxOrderID=mrs(0)
	if isnull(MaxOrderID) then MaxOrderID=0
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select * From Channel Where ChannelName='" & ChannelName & "'",conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>频道名称已经存在！</li>"
		rs.close
	    set rs=Nothing
    	exit sub
	end if
    rs.addnew
	rs("OrderID")=MaxOrderID+1
    rs("ChannelName")=ChannelName
	rs("ReadMe")=ReadMe
	rs("LinkUrl")=LinkUrl
	rs.update
    rs.Close
    set rs=Nothing
	Response.Redirect "Admin_Channel.asp"  
end sub

sub SaveModify()
	dim ChannelID,ChannelName,ReadMe,LinkUrl
	ChannelID=trim(request("ChannelID"))
	ChannelName=trim(request.Form("ChannelName"))
	ReadMe=trim(request("ReadMe"))
	LinkUrl=trim(request("LinkUrl"))
	if ChannelID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的频道ID</li>"
	else
		ChannelID=Clng(ChannelID)
	end if
	if ChannelName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>频道名称不能为空！</li>"
	end if
	if LinkUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接地址不能为空！</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if

	sql="Select * From Channel Where ChannelID=" & ChannelID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,3
	if rs.bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的频道！</li>"
		rs.close
	    set rs=Nothing
    else
		rs("ChannelName")=ChannelName
		rs("ReadMe")=ReadMe
		rs("LinkUrl")=LinkUrl
		rs.update
		rs.Close
		set rs=Nothing
		Response.Redirect "Admin_Channel.asp"  
	end if
end sub

sub DelChannel()
	dim ChannelID
	ChannelID=trim(request("ChannelID"))
	if ChannelID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的频道ID</li>"
	else
		ChannelID=Clng(ChannelID)
	end if
	
	if FoundErr=True then
		exit sub
	end if

	sql="Select * From Channel Where ChannelID=" & ChannelID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,3
	if rs.bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的频道！</li>"

    else
		if rs("ChannelType")=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>不能删除系统频道</li>"
		end if
	end if
	if FoundErr=True then
		rs.Close
		set rs=Nothing
	else
		rs.delete
		rs.update
		rs.Close
		set rs=Nothing
		Response.Redirect "Admin_Channel.asp"  
	end if
end sub

sub UpOrder()
	dim ChannelID,sqlOrder,rsOrder,MoveNum,cOrderID,tOrderID,i,rs
	ChannelID=trim(request("ChannelID"))
	cOrderID=Trim(request("cOrderID"))
	MoveNum=trim(request("MoveNum"))
	if ChannelID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		ChannelID=CLng(ChannelID)
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
	set mrs=conn.execute("select max(OrderID) From Channel")
	MaxOrderID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
	
	'然后将位于当前栏目以上的栏目的OrderID依次加一，范围为要提升的数字
	sqlOrder="select * From Channel where OrderID<" & cOrderID & " order by OrderID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tOrderID=rsOrder("OrderID")       '得到要提升位置的OrderID，包括子栏目
		conn.execute("update Channel set OrderID=OrderID+1 where OrderID=" & tOrderID)
		i=i+1
		if i>MoveNum then
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)
	call CloseConn()      
	response.redirect "Admin_Channel.asp"
end sub

sub DownOrder()
	dim ChannelID,sqlOrder,rsOrder,MoveNum,cOrderID,tOrderID,i,rs,PrevID,NextID
	ChannelID=trim(request("ChannelID"))
	cOrderID=Trim(request("cOrderID"))
	MoveNum=trim(request("MoveNum"))
	if ChannelID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		ChannelID=CLng(ChannelID)
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
	set mrs=conn.execute("select max(OrderID) From Channel")
	MaxOrderID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update Channel set OrderID=" & MaxOrderID & " where ChannelID=" & ChannelID)
	
	'然后将位于当前栏目以下的栏目的OrderID依次减一，范围为要下降的数字
	sqlOrder="select * From Channel where OrderID>" & cOrderID & " order by OrderID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tOrderID=rsOrder("OrderID")       '得到要提升位置的OrderID，包括子栏目
		conn.execute("update Channel set OrderID=OrderID-1 where OrderID=" & tOrderID)
		i=i+1
		if i>MoveNum then
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update Channel set OrderID=" & tOrderID & " where ChannelID=" & ChannelID)
	call CloseConn()      
	response.redirect "Admin_Channel.asp"
end sub
%>
