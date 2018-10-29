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
dim Action,FoundErr,ErrMsg
dim rs,sql
dim SkinCount,LayoutCount
Action=trim(request("Action"))
%>
<html>
<head>
<title>课程管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>课 程 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td><a href="Admin_Course.asp">课程管理</a> | <a href="Admin_Course_Add.asp">添加新课程</a></td>
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
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from CourseList , Special , Admin where CourseList.SpecialID=Special.SpecialID and CourseList.TeacherName='" & session("AdminTrueName") &  "'"
	rs.Open sql,conn,1,1
%>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"><!-- 此行列出教师所开课程的列表的表头-->
    <td height="22"   width="300" align="center"><strong>课程名称</strong></td>
    <td width="100" align="center"><strong>课程说明</strong></td>
    <td width="100" align="center"><strong>开课时间</strong></td>
    <td width="200" align="center"><strong>上课班级</strong></td>
    <td width="110" align="center"><strong>上课班级所属学院</strong></td>
    <!--<td width="100" height="22" align="center"><strong> 常规操作</strong></td>-->
  </tr>
  <%do while not rs.EOF %>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td align="center"><a href="Admin_ArticleManageSpecial.asp?SpecialID=<%=rs("SpecialID")%>" title="点击进入管理此课程的文章"><%=rs("SpecialName")%></a></td>
    <td width="200"><%=dvhtmlencode(rs("ReadMe"))%></td>
    <td width="100" align="center">
      <%
	select case rs("BrowsePurview")
	case 9999
		response.write "游客"
	case 999
		response.write "注册用户"
	case 99
		response.write "收费用户"
	case 9
		response.write "VIP用户"
	case 5
		response.write "管理员"
	end select%>
    </td>
    <td width="100" align="center">
      <%
	select case rs("AddPurview")
	case 999
		response.write "注册用户"
	case 99
		response.write "收费用户"
	case 9
		response.write "VIP用户"
	case 5
		response.write "管理员"
	end select%>
    </td>
    <td width="100" align="center"><%
	response.write "<a href='Admin_Special.asp?action=Modify&SpecialID=" & rs("SpecialID") & "'>修改</a>&nbsp;&nbsp;"
	response.write "<a href='Admin_Special.asp?Action=Del&SpecialID=" & rs("SpecialID") & "' onClick=""return confirm('确定要删除此课程吗？删除此课程后原属于此课程的文章将不属于任何课程。');"">删除</a>&nbsp;&nbsp;" 
    response.write "<a href='Admin_Special.asp?Action=Clear&SpecialID=" & rs("SpecialID") & "' onClick=""return confirm('确定要清空此课程中的文章吗？本操作将原属于此课程的文章改为不属于任何课程。');"">清空</a>"
	%></td>
    <form action='Admin_Special.asp?Action=UpOrder' method='post'>
    </form>
    <form action='Admin_Special.asp?Action=DownOrder' method='post'>
    </form>
  </tr>
  <%
		rs.MoveNext
   	loop
  %>
</table> 
<%
	rs.Close
	set rs=Nothing
end sub

sub ShowOrder()
	dim iCount,i,j
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from Special"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
	j=1
%>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" colspan="4" align="center"><strong> 课 程 </strong><strong>排 
      序</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td align="center"><a href="Admin_ArticleManageSpecial.asp?SpecialID=<%=rs("SpecialID")%>" title="点击进入管理此课程的文章"><%=rs("SpecialName")%></a></td>
    <form action='Admin_Special.asp?Action=UpOrder' method='post'>
      <td width='120' align="center"> <% 
	if j>1 then 
		response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=SpecialID value="&rs("SpecialID")&">"
		response.write "<input type=hidden name=cOrderID value="&rs("OrderID")&">&nbsp;<input type=submit name=Submit value=修改>" 
	else 
		response.write "&nbsp;" 
	end if 
%> </td>
    </form>
    <form action='Admin_Special.asp?Action=DownOrder' method='post'>
      <td width='120' align="center"> <%
	if iCount>j then 
		response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=SpecialID value="&rs("SpecialID")&">"
		response.write "<input type=hidden name=cOrderID value="&rs("OrderID")&">&nbsp;<input type=submit name=Submit value=修改>" 
	else 
		response.write "&nbsp;" 
	end if 
%> </td>
      <td width='200' align="center">&nbsp;</td>
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

sub AddSpecial()
%>
<form method="post" action="Admin_Special.asp" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>添 加 新 课 程</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong> 课程名称：</strong></td>
      <td class="tdbg"><input name="SpecialName" type="text" id="SpecialName" size="49" maxlength="45"> 
        &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong> 课程简称：</strong>&nbsp;(四个汉字或以内)</td>
      <td class="tdbg"><input name="SpecialAbbreviation" type="text" id="SpecialAbbreviation" size="49" maxlength="8"> 
        &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>课程说明</strong><br>
        鼠标移至课程名称上时将显示设定的说明文字（不支持HTML）</td>
      <td class="tdbg"><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>默认配色模板：</strong><br>
        相关模板中包含CSS、颜色、图片等信息</td>
      <td class="tdbg">
        <%call Admin_ShowSkin_Option(0)%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>版面设计模板：</strong><br>相关模板中包含了版面设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色模板”失效。
        </td>
      <td class="tdbg">
        <%call Admin_ShowLayout_Option(4,0)%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>课程浏览权限：</strong><br>
        只有具有相应权限的人才能浏览此课程中的文章。</td>
      <td><select name="BrowsePurview" id="BrowsePurview">
          <option value="9999">游客</option>
          <option value="999">注册用户</option>
          <option value="99">收费用户</option>
          <option value="9">VIP用户</option>
          <option value="5">管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>课程发表文章权限：</strong><br>
        只有具有相应权限的人才能在此课程中发表文章。</td>
      <td><select name="AddPurview" id="AddPurview">
          <option value="999">注册用户</option>
          <option value="99">收费用户</option>
          <option value="9">VIP用户</option>
          <option value="5">管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input  type="submit" name="Submit" value=" 添 加 ">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Special.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
  </form>
<%
end sub

sub Modify()
	dim SpecialID
	SpecialID=trim(request("SpecialID"))
	if SpecialID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的课程ID！</li>"
		exit sub
	else
		SpecialID=Clng(SpecialID)
	end if
	sql="Select * From Special Where SpecialID=" & SpecialID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,3
	if rs.bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的课程，可能已经被删除！</li>"
	else

%>
<form method="post" action="Admin_Special.asp" name="form1">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>修 改 课 程</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong> 课程名称：</strong></td>
      <td class="tdbg"><input name="SpecialName" type="text" id="SpecialName" value="<%=rs("SpecialName")%>" size="49" maxlength="45">
        <input name="SpecialID" type="hidden" id="SpecialID" value="<%=rs("SpecialID")%>"> </td>
    </tr>
   <!-- 课程简称-->
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong> 课程名称：&nbsp;(四个汉字或以内)</strong></td>
      <td class="tdbg"><input name="SpecialAbbreviation" type="text" id="SpecialAbbreviation" value="<%=rs("SpecialAbbreviation")%>" size="49" maxlength="8">
      </td>
    </tr>
    <!--结束课程简称-->
    
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>课程说明</strong><br>
        鼠标移至课程名称上时将显示设定的说明文字（不支持HTML）</td>
      <td class="tdbg"><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"><%=rs("ReadMe")%>
</textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>默认配色模板：</strong><br>
        相关模板中包含CSS、颜色、图片等信息</td>
      <td class="tdbg">
        <%call Admin_ShowSkin_Option(rs("SkinID"))%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350" class="tdbg"><strong>版面设计模板：</strong><br>相关模板中包含了版面设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色模板”失效。
        </td>
      <td class="tdbg">
        <%call Admin_ShowLayout_Option(4,rs("SkinID"))%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>课程浏览权限：</strong><br>
        只有具有相应权限的人才能浏览此课程中的文章。</td>
      <td><select name="BrowsePurview" id="select">
          <option value="9999" <%if rs("BrowsePurview")=9999 then response.write " selected"%>>游客</option>
          <option value="999" <%if rs("BrowsePurview")=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if rs("BrowsePurview")=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if rs("BrowsePurview")=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if rs("BrowsePurview")=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>课程发表文章权限：</strong><br>
        只有具有相应权限的人才能在此课程中发表文章。</td>
      <td><select name="AddPurview" id="select2">
          <option value="999" <%if rs("AddPurview")=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if rs("AddPurview")=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if rs("AddPurview")=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if rs("AddPurview")=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveModify">
        <input  type="submit" name="Submit" value="保存修改结果">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Special.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub ShowUniteForm()
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" colspan="3" align="center"><strong>合 并 课 程</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td height="100"><form name="myform" method="post" action="Admin_Special.asp" onSubmit="return ConfirmUnite();">
        &nbsp;&nbsp;将课程 
        <select name="SpecialID" id="SpecialID">
        <%call ShowSpecial()%>
        </select>
        合并到
        <select name="TargetSpecialID" id="TargetSpecialID">
        <%call ShowSpecial()%>
        </select>
        <br> <br>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="UniteSpecial">
        <input type="submit" name="Submit" value=" 合并课程 " style="cursor:hand;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_ClassManage.asp'" style="cursor:hand;">
      </form>
	</td>
  </tr>
  <tr class="tdbg"> 
    <td height="60"><strong>注意事项：</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;所有操作不可逆，请慎重操作！！！<br>
      &nbsp;&nbsp;&nbsp;&nbsp;不能在同一个课程内进行操作。<br>
      &nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的课程将被删除，所有文章将转移到目标课程中。</td>
  </tr>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmUnite()
{
  if (document.myform.SpecialID.value==document.myform.TargetSpecialID.value)
  {
    alert("请不要在相同课程内进行操作！");
	document.myform.TargetSpecialID.focus();
	return false;
  }
}
</script>
<%
end sub
%>
</body>
</html>
<%
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
