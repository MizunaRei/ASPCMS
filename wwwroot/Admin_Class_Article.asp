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
<!--#include file="inc/admin_code_article.asp"-->
<%
dim Action,ParentID,SkinID,LayoutID,BrowsePurview,AddPurview,i,FoundErr,ErrMsg
dim SkinCount,LayoutCount
Action=trim(Request("Action"))
ParentID=trim(request("ParentID"))
if ParentID="" then
	ParentID=0
else
	ParentID=CLng(ParentID)
end if
%>
<html>
<head>
<title>文章栏目管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">

</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>文 章 栏 目 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Class_Article.asp">文章栏目管理首页</a> | <a href="Admin_Class_Article.asp?Action=Add">添加文章栏目</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Order">一级栏目排序</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=OrderN">N级栏目排序</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Reset">复位所有文章栏目</a>&nbsp;|&nbsp;<a href="Admin_Class_Article.asp?Action=Unite">文章栏目合并</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddClass()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Move" then
	call MoveClass()
elseif Action="SaveMove" then
	call SaveMove()
elseif Action="Del" then
	call DeleteClass()
elseif Action="Clear" then
	call ClearClass()
elseif Action="UpOrder" then 
	call UpOrder() 
elseif Action="DownOrder" then 
	call DownOrder() 
elseif Action="Order" then
	call Order()
elseif Action="UpOrderN" then 
	call UpOrderN() 
elseif Action="DownOrderN" then 
	call DownOrderN() 
elseif Action="OrderN" then
	call OrderN()
elseif Action="Reset" then
	call Reset()
elseif Action="SaveReset" then
	call SaveReset()
elseif Action="Unite" then
	call Unite()
elseif Action="SaveUnite" then
	call SaveUnite()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()


sub main()
	dim arrShowLine(10)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim sqlClass,rsClass,i,iDepth
	sqlClass="select * From ArticleClass order by RootID,OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
%>
<br> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" align="center"><strong>栏目名称</strong></td>
    <td width="100" align="center"><strong>管理员</strong></td>
    <td width="80" align="center"><strong>栏目属性</strong></td>
    <td width="60" align="center"><strong>浏览权限</strong></td>
    <td width="60" align="center"><strong>添加权限</strong></td>
    <td width="300" height="22" align="center"><strong>操作选项</strong></td>
  </tr>
  <% 
do while not rsClass.eof 
%>
  <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
    <td> <% 
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
	  response.write "<a href='Admin_Class_Article.asp?Action=Modify&ClassID=" & rsClass("ClassID") & "' title='" & rsClass("ReadMe") & "'>" & rsClass("ClassName") & "</a>"
	  if rsClass("Child")>0 then 
	  	response.write "（" & rsClass("Child") & "）" 
	  end if
	  
	  
	  'response.write "&nbsp;&nbsp;" & rsClass("ClassID") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
	  %> </td>
    <td> <%
	if rsClass("ClassMaster")<>"" then
		response.write rsClass("ClassMaster")
	else
		response.write "&nbsp;"
	end if
	%> </td>
    <td width="80" align="center"> <%
	if rsClass("LinkUrl")<>"" then
		response.write "<font color=red>外部</font>，"
	else
		response.write "<font color=green>系统</font>，"
	end if
	if rsClass("IsElite")=True then
		response.write "<font color=blue>推荐</font>"
	else
		response.write "普通"
	end if
	%> </td>
    <td align="center"> <%
	select case rsClass("BrowsePurview")
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
	end select%> </td>
    <td align="center">
      <%
	select case rsClass("AddPurview")
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
    <td align="center"><a href="Admin_Class_Article.asp?Action=Add&ParentID=<%=rsClass("ClassID")%>">添加子栏目</a> 
      | <a href="Admin_Class_Article.asp?Action=Modify&ClassID=<%=rsClass("ClassID")%>">修改设置</a> 
      | <a href="Admin_Class_Article.asp?Action=Move&ClassID=<%=rsClass("ClassID")%>">移动栏目</a> 
      | <a href="Admin_Class_Article.asp?Action=Clear&ClassID=<%=rsClass("ClassID")%>" onClick="return ConfirmDel3();">清空</a> 
      | <a href="Admin_Class_Article.asp?Action=Del&ClassID=<%=rsClass("ClassID")%>" onClick="<%if rsClass("Child")>0 then%>return ConfirmDel1();<%else%>return ConfirmDel2();<%end if%>">删除</a></td>
  </tr>
  <% 
	rsClass.movenext 
loop 
%>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmDel1()
{
   alert("此栏目下还有子栏目，必须先删除下属子栏目后才能删除此栏目！");
   return false;
}

function ConfirmDel2()
{
   if(confirm("删除栏目将同时删除此栏目中的所有文章，并且不能恢复！确定要删除此栏目吗？"))
     return true;
   else
     return false;
	 
}
function ConfirmDel3()
{
   if(confirm("清空栏目将把栏目（包括子栏目）的所有文章放入回收站中！确定要清空此栏目吗？"))
     return true;
   else
     return false;
	 
}
</script>
<br><br>
<%
end sub

sub AddClass()
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onsubmit="return check()">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>添 加 文 章 栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属栏目：</strong><br>
        不能指定为外部栏目 </td>
      <td> <select name="ParentID">
          <%call Admin_ShowClass_Option(0,ParentID)%>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="ClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目说明：<br>
        </strong> 鼠标移至栏目名称上时将显示设定的说明文字（不支持HTML）</td>
      <td><textarea name="Readme" cols="30" rows="4" id="Readme"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td><strong>是否是推荐栏目：</strong><br>
        推荐栏目将在首页及此栏目的父栏目上显示文章列表</td>
      <td><input name="IsElite" type="radio" value="Yes" checked>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="IsElite" value="No">
        否 </td>
    </tr>
    <tr class="tdbg"> 
      <td><strong>是否在顶部导航栏显示：</strong><br>
        此选项只对一级栏目有效。</td>
      <td><input name="ShowOnTop" type="radio" value="Yes" checked>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="ShowOnTop" value="No">
        否 </td>
    </tr>
    <tr class="tdbg">
      <td><strong>栏目内的文章在首页的显示样式：</strong><br>
        此选项只对一级栏目有效</td>
      <td><select name="Setting" id="Setting">
          <option value="0" selected>图片文章+普通文章</option>
          <option value="1">只显示普通文章</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目配色模板：</strong><br>
        相关模板中包含CSS、颜色、图片等信息</td>
      <td><%call Admin_ShowSkin_Option(SkinID)%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>版面设计模板：</strong><br>
        相关模板中包含了栏目设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色模板”失效。 </td>
      <td><%call Admin_ShowLayout_Option(2,LayoutID)%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目图片地址：</strong><br>
        图片会显示在栏目前面。注意图片大小。</td>
      <td><input name="ClassPicUrl" type="text" id="ClassPicUrl" size="37" maxlength="255">
        （预留功能）</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目编辑：</strong><br>
        多个编辑请用“|”分隔，如：webboy|dilys|sws2000<br>
        无需添加“文章总编”以上级别的管理员<br>
        管理员权限采用权限继承制度</td>
      <td><input name="ClassMaster" type="text" id="ClassMaster" size="37" maxlength="100" disabled> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目链接地址：</strong><br>
        如果想将栏目链接到外部地址，请输入完整的URL地址，否则请保持为空。</td>
      <td><input name="LinkUrl" type="text" id="LinkUrl" size="37" maxlength="255"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目浏览权限：</strong><br>
        只有具有相应权限的人才能浏览此栏目中的文章。</td>
      <td><select name="BrowsePurview" id="BrowsePurview">
          <option value="9999" <%if BrowsePurview=9999 then response.write " selected"%>>游客</option>
          <option value="999" <%if BrowsePurview=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if BrowsePurview=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if BrowsePurview=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if BrowsePurview=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目发表文章权限：</strong><br>
        只有具有相应权限的人才能在此栏目中发表文章。</td>
      <td><select name="AddPurview" id="AddPurview">
          <option value="999" <%if AddPurview=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if AddPurview=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if AddPurview=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if AddPurview=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input name="Add" type="submit" value=" 添 加 " <%if SkinCount=0 or LayoutCount=0 then response.write " disabled"%> style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor:hand;"> 
        <%if SkinCount=0 then response.write "<li><font color=red>请先添加栏目配色模板</font></li>"
		if SkinCount=0 then response.write "<li><font color=red>请先添加栏目栏目设计模板</font></li>" %></td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.ClassName.value=="")
  {
    alert("栏目名称不能为空！");
	document.form1.ClassName.focus();
	return false;
  }
}
</script>
<%
end sub

sub Modify()
	dim ClassID,sql,rsClass,i
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	sql="select * From ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onsubmit="return check()">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>修 改 文 章 栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属栏目：</strong><br>
        如果你想改变所属栏目，请<a href='Admin_Class_Article.asp?Action=Move&ClassID=<%=ClassID%>'>点此移动栏目</a></td>
      <td> <%
	if rsClass("ParentID")<=0 then
	  	response.write "无（作为一级栏目）"
	else
    	dim rsParentClass,sqlParentClass
		sqlParentClass="Select * From ArticleClass where ClassID in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParentClass=server.CreateObject("adodb.recordset")
		rsParentClass.open sqlParentClass,conn,1,1
		do while not rsParentClass.eof
			for i=1 to rsParentClass("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParentClass("Depth")>0 then
				response.write "└"
			end if
			response.write "&nbsp;" & rsParentClass("ClassName") & "<br>"
			rsParentClass.movenext
		loop
		rsParentClass.close
		set rsParentClass=nothing
	end if
	%> </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="ClassName" type="text" value="<%=rsClass("ClassName")%>" size="37" maxlength="20"> 
        <input name="ClassID" type="hidden" id="ClassID" value="<%=rsClass("ClassID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目说明：<br>
        </strong> 鼠标移至栏目名称上时将显示设定的说明文字（不支持HTML）</td>
      <td><textarea name="Readme" cols="30" rows="4" id="Readme"><%=rsClass("ReadMe")%></textarea></td>
    </tr>
    <tr class="tdbg">
      <td><strong>是否是推荐栏目：</strong><br>
        推荐栏目将在首页及此栏目的父栏目上显示文章列表</td>
      <td> <input name="IsElite" type="radio" value="Yes" <%if rsClass("IsElite")=True then response.write " checked"%>>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="IsElite" value="No" <%if rsClass("IsElite")=False then response.write " checked"%>>
        否 </td>
    </tr>
    <tr class="tdbg">
      <td><strong>是否在顶部导航栏显示：</strong><br>
        只选项只对一级栏目有效。</td>
      <td><input name="ShowOnTop" type="radio" value="Yes" <%if rsClass("ShowOnTop")=True then response.write " checked"%>>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="ShowOnTop" value="No" <%if rsClass("ShowOnTop")=False then response.write " checked"%>>
        否 </td>
    </tr>
    <tr class="tdbg">
      <td><strong>栏目内的文章在首页的显示样式：</strong><br>
        此选项只对一级栏目有效</td>
      <td><select name="Setting" id="Setting">
          <option value="0" <%if rsClass("Setting")=0 then response.write " selected"%>>图片文章+普通文章</option>
          <option value="1" <%if rsClass("Setting")=1 then response.write " selected"%>>只显示普通文章</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目配色模板：</strong><br>
        相关模板中包含CSS、颜色、图片等信息</td>
      <td><%call Admin_ShowSkin_Option(rsClass("SkinID"))%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>版面设计模板：</strong><br>
        相关模板中包含了栏目设计的版式等信息，如果是自行添加的设计模板，可能会导致“栏目配色模板”失效。 </td>
      <td><%call Admin_ShowLayout_Option(2,rsClass("LayoutID"))%></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目图片地址：</strong><br>
        图片会显示在栏目前面。注意图片大小。</td>
      <td><input name="ClassPicUrl" type="text" id="ClassPicUrl" value="<%=rsClass("ClassPicUrl")%>" size="37" maxlength="255">
        （预留功能）</td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目编辑：</strong><br>
        多个编辑请用“|”分隔，如：webboy|dilys|sws2000<br>
        无需添加“文章总编”以上级别的管理员<br>
        管理员权限采用权限继承制度</td>
      <td><input name="ClassMaster" type="text" id="ClassMaster" value="<%=rsClass("ClassMaster")%>" size="37" maxlength="100" disabled> 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目链接地址：</strong><br>
        如果想将栏目链接到外部地址，请输入完整的URL地址，否则请保持为空。</td>
      <td><input name="LinkUrl" type="text" id="LinkUrl" value="<%=rsClass("LinkUrl")%>" size="37" maxlength="255"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目浏览权限：</strong><br>
        只有具有相应权限的人才能浏览此栏目中的文章。</td>
      <td><select name="BrowsePurview" id="BrowsePurview">
          <option value="9999" <%if rsClass("BrowsePurview")=9999 then response.write " selected"%>>游客</option>
          <option value="999" <%if rsClass("BrowsePurview")=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if rsClass("BrowsePurview")=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if rsClass("BrowsePurview")=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if rsClass("BrowsePurview")=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目发表文章权限：</strong><br>
        只有具有相应权限的人才能在此栏目中发表文章。</td>
      <td><select name="AddPurview" id="AddPurview">
          <option value="999" <%if rsClass("AddPurview")=999 then response.write " selected"%>>注册用户</option>
          <option value="99" <%if rsClass("AddPurview")=99 then response.write " selected"%>>收费用户</option>
          <option value="9" <%if rsClass("AddPurview")=9 then response.write " selected"%>>VIP用户</option>
          <option value="5" <%if rsClass("AddPurview")=5 then response.write " selected"%>>管理员</option>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
        <input name="Submit" type="submit" value=" 保存修改结果 " <%if SkinCount=0 or LayoutCount=0 then response.write " disabled"%> style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor:hand;"> 
        <%if SkinCount=0 then response.write "<li><font color=red>请先添加栏目配色模板</font></li>"
		if SkinCount=0 then response.write "<li><font color=red>请先添加栏目栏目设计模板</font></li>" %></td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.ClassName.value=="")
  {
    alert("栏目名称不能为空！");
	document.form1.ClassName.focus();
	return false;
  }
}
</script>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub MoveClass()
	dim ClassID,sql,rsClass,i
	dim SkinID,LayoutID,BrowsePurview,AddPurview
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select * From ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>移 动 文 章 栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>栏目名称：</strong></td>
      <td><%=rsClass("ClassName")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rsClass("ClassID")%>"></td>
    </tr>
    <tr class="tdbg">
      <td width="200"><strong>当前所属栏目：</strong></td>
      <td>
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "无（作为一级栏目）"
	else
    	dim rsParent,sqlParent
		sqlParent="Select * From ArticleClass where ClassID in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParent=server.CreateObject("adodb.recordset")
		rsParent.open sqlParent,conn,1,1
		do while not rsParent.eof
			for i=1 to rsParent("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParent("Depth")>0 then
				response.write "└"
			end if
			response.write "&nbsp;" & rsParent("ClassName") & "<br>"
			rsParent.movenext
		loop
		rsParent.close
		set rsParent=nothing
	end if
	%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>移动到：</strong><br>
        不能指定为当前栏目的下属子栏目<br>
        不能指定为外部栏目</td>
      <td><select name="ParentID" size="2" style="height:300px;width:500px;"><%call Admin_ShowClass_Option(0,rsClass("ParentID"))%></select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveMove"> 
        <input name="Submit" type="submit" value=" 保存移动结果 " style="cursor:hand;">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor:hand;"></td></tr>
  </table>
</form>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub Order() 
	dim sqlClass,rsClass,i,iCount,j 
	sqlClass="select * From ArticleClass where ParentID=0 order by RootID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
	iCount=rsClass.recordcount 
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border"> 
  <tr class="title">  
    <td height="22" colspan="4" align="center"><strong>一 级 栏 目 排 序</strong></td> 
  </tr> 
  <% 
j=1 
do while not rsClass.eof 
%> 
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">  
      <td width="200">&nbsp;<%=rsClass("ClassName")%></td> 
<% 
	if j>1 then 
  		response.write "<form action='Admin_Class_Article.asp?Action=UpOrder' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修改>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
	if iCount>j then 
  		response.write "<form action='Admin_Class_Article.asp?Action=DownOrder' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修改>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
	</tr> 
  <% 
	j=j+1 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub OrderN() 
	dim sqlClass,rsClass,i,iCount,trs,UpMoveNum,DownMoveNum 
	sqlClass="select * From ArticleClass order by RootID,OrderID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border"> 
  <tr class="title">  
    <td height="22" colspan="4" align="center"><strong>N 级 栏 目 排 序</strong></td> 
  </tr> 
  <% 
do while not rsClass.eof 
%> 
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">  
      <td width="300"> 
	  <% 
	for i=1 to rsClass("Depth") 
	  	response.write "&nbsp;&nbsp;&nbsp;" 
	next 
	if rsClass("Child")>0 then 
		response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	else 
	  	response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	end if 
	if rsClass("ParentID")=0 then 
		response.write "<b>" 
	end if 
	response.write rsClass("ClassName") 
	if rsClass("Child")>0 then 
		response.write "(" & rsClass("Child") & ")" 
	end if 
	%></td> 
<% 
	if rsClass("ParentID")>0 then   '如果不是一级栏目，则算出相同深度的栏目数目，得到该栏目在相同深度的栏目中所处位置（之上或者之下的栏目数） 
		'所能提升最大幅度应为For i=1 to 该版之上的版面数 
		set trs=conn.execute("select count(ClassID) From ArticleClass where ParentID="&rsClass("ParentID")&" and OrderID<"&rsClass("OrderID")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		if UpMoveNum>0 then 
  			response.write "<form action='Admin_Class_Article.asp?Action=UpOrderN' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
			for i=1 to UpMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">&nbsp;<input type=submit name=Submit value=修改>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
		'所能降低最大幅度应为For i=1 to 该版之下的版面数 
		set trs=conn.execute("select count(ClassID) From ArticleClass where ParentID="&rsClass("ParentID")&" and orderID>"&rsClass("orderID")&"") 
		DownMoveNum=trs(0) 
		if isnull(DownMoveNum) then DownMoveNum=0 
		if DownMoveNum>0 then 
  			response.write "<form action='Admin_Class_Article.asp?Action=DownOrderN' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
			for i=1 to DownMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=ClassID value="&rsClass("ClassID")&">&nbsp;<input type=submit name=Submit value=修改>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
	else 
		response.write "<td colspan=2>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
	</tr> 
  <% 
	UpMoveNum=0 
	DownMoveNum=0 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub Reset() 
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border"> 
  <tr class="title">  
    <td height="22" colspan="3" align="center"><strong>复 位 所 有 文 章 栏 目</strong></td> 
  </tr> 
    <tr class="tdbg">  
    <td align="center">  
      <form name="form1" method="post" action="Admin_Class_Article.asp?Action=SaveReset"> 
        <table width="80%" border="0" cellspacing="0" cellpadding="0"> 
          <tr>  
            <td height="150"><font color="#FF0000"><strong>注意：</strong></font><br> 
              &nbsp;&nbsp;&nbsp;&nbsp;如果选择复位所有栏目，则所有栏目都将作为一级栏目，这时您需要重新对各个栏目进行归属的基本设置。不要轻易使用该功能，仅在做出了错误的设置而无法复原栏目之间的关系和排序的时候使用。  
            </td> 
          </tr> 
        </table> 
        <input type="submit" name="Submit" value="复位所有栏目"> &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor:hand;">
      </form></td>
    </tr>
</table>
<%
end sub

sub Unite()
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" colspan="3" align="center"><strong>文 章 栏 目 合 并</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td height="100"><form name="myform" method="post" action="Admin_Class_Article.asp" onSubmit="return ConfirmUnite();">
        &nbsp;&nbsp;将栏目 
        <select name="ClassID" id="ClassID">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
        合并到
        <select name="TargetClassID" id="TargetClassID">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
        <br> <br>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="SaveUnite">
        <input type="submit" name="Submit" value=" 合并栏目 " style="cursor:hand;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Class_Article.asp'" style="cursor:hand;">
      </form>
	</td>
  </tr>
  <tr class="tdbg"> 
    <td height="60"><strong>注意事项：</strong><br>
      &nbsp;&nbsp;&nbsp;&nbsp;所有操作不可逆，请慎重操作！！！<br>
      &nbsp;&nbsp;&nbsp;&nbsp;不能在同一个栏目内进行操作，不能将一个栏目合并到其下属栏目中。目标栏目中不能含有子栏目。<br>
      &nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的栏目（或者包括其下属栏目）将被删除，所有文章将转移到目标栏目中。</td>
  </tr>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmUnite()
{
  if (document.myform.ClassID.value==document.myform.TargetClassID.value)
  {
    alert("请不要在相同栏目内进行操作！");
	document.myform.TargetClassID.focus();
	return false;
  }
  if (document.myform.TargetClassID.value=="")
  {
    alert("目标栏目不能指定为含有子栏目的栏目！");
	document.myform.TargetClassID.focus();
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
	dim ClassID,ClassName,IsElite,ShowOnTop,Setting,Readme,ClassMaster,ClassPicUrl,LinkUrl,PrevOrderID
	dim sql,rs,trs
	dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,MaxClassID,MaxRootID
	dim PrevID,NextID,Child

	ClassName=trim(request("ClassName"))
	ClassMaster=trim(request("ClassMaster"))
	IsElite=trim(request("IsElite"))
	ShowOnTop=trim(request("ShowOnTop"))
	Setting=trim(request("Setting"))
	Readme=trim(request("Readme"))
	ClassPicUrl=trim(request("ClassPicUrl"))
	LinkUrl=trim(request("LinkUrl"))
	SkinID=Trim(request("SkinID"))
	LayoutID=trim(request("LayoutID"))
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if ClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目名称不能为空！</li>"
	end if
	if IsElite="Yes" then
		IsElite=True
	else
		IsElite=False
	end if
	if ShowOnTop="Yes" then
		ShowOnTop=True
	else
		ShowOnTop=False
	end if
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
	if ClassMaster<>"" then
		'call AddMaster(ClassMaster)
	end if
	if FoundErr=True then
		exit sub
	end if

	set rs = conn.execute("select Max(ClassID) From ArticleClass")
	MaxClassID=rs(0)
	if isnull(MaxClassID) then
		MaxClassID=0
	end if
	rs.close
	ClassID=MaxClassID+1
	set rs=conn.execute("select max(rootid) From ArticleClass")
	MaxRootID=rs(0)
	if isnull(MaxRootID) then
		MaxRootID=0
	end if
	rs.close
	RootID=MaxRootID+1
	
	if ParentID>0 then
		sql="select * From ArticleClass where ClassID=" & ParentID & ""
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属栏目已经被删除！</li>"
		else
			if rs("LinkUrl")<>"" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>不能指定外部栏目为所属栏目！</li>"
			end if
		end if
		if FoundErr=True then
			rs.close
			set rs=nothing
			exit sub
		else	
			RootID=rs("RootID")
			ParentName=rs("ClassName")
			ParentDepth=rs("Depth")
			ParentPath=rs("ParentPath")
			Child=rs("Child")
			ParentPath=ParentPath & "," & ParentID     '得到此栏目的父级栏目路径
			PrevOrderID=rs("OrderID")
			if Child>0 then		
				dim rsPrevOrderID
				'得到与本栏目同级的最后一个栏目的OrderID
				set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & ParentID & " and OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'得到同一父栏目但比本栏目级数大的子栏目的最大OrderID，如果比前一个值大，则改用这个值。
				set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentPath like '" & ParentPath & ",%'")
				if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
					if not IsNull(rsPrevOrderID(0))  then
				 		if rsPrevOrderID(0)>PrevOrderID then
							PrevOrderID=rsPrevOrderID(0)
						end if
					end if
				end if
			else
				PrevID=0
			end if

		end if
		rs.close
	else
		if MaxRootID>0 then
			set trs=conn.execute("select ClassID from ArticleClass where RootID=" & MaxRootID & " and Depth=0")
			PrevID=trs(0)
			trs.close
		else
			PrevID=0
		end if
		PrevOrderID=0
		ParentPath="0"
	end if

	sql="Select * From ArticleClass Where ParentID=" & ParentID & " AND ClassName='" & ClassName & "'"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		if ParentID=0 then
			ErrMsg=ErrMsg & "<br><li>已经存在一级栏目：" & ClassName & "</li>"
		else
			ErrMsg=ErrMsg & "<br><li>“" & ParentName & "”中已经存在子栏目“" & ClassName & "”！</li>"
		end if
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close

	sql="Select top 1 * From ArticleClass"
	rs.open sql,conn,1,3
    rs.addnew
	rs("ClassID")=ClassID
   	rs("ClassName")=ClassName
	rs("IsElite")=IsElite
	rs("ShowOnTop")=ShowOnTop
	rs("Setting")=Clng(Setting)
	'rs("ClassMaster")=ClassMaster
	rs("RootID")=RootID
	rs("ParentID")=ParentID
	if ParentID>0 then
		rs("Depth")=ParentDepth+1
	else
		rs("Depth")=0
	end if
	rs("ParentPath")=ParentPath
	rs("OrderID")=PrevOrderID
	rs("Child")=0
	rs("Readme")=Readme
	rs("ClassPicUrl")=ClassPicUrl
	rs("LinkUrl")=LinkUrl
	rs("SkinID")=SkinID
	rs("LayoutID")=LayoutID
	rs("BrowsePurview")=Cint(BrowsePurview)
	rs("AddPurview")=Cint(AddPurview)
	rs("PrevID")=PrevID
	rs("NextID")=0
	rs.update
	rs.Close
    set rs=Nothing
	
	'更新与本栏目同一父栏目的上一个栏目的“NextID”字段值
	if PrevID>0 then
		conn.execute("update ArticleClass set NextID=" & ClassID & " where ClassID=" & PrevID)
	end if
	
	if ParentID>0 then
		'更新其父类的子栏目数
		conn.execute("update ArticleClass set child=child+1 where ClassID="&ParentID)
		
		'更新该栏目排序以及大于本需要和同在本分类下的栏目排序序号
		conn.execute("update ArticleClass set OrderID=OrderID+1 where rootid=" & rootid & " and OrderID>" & PrevOrderID)
		conn.execute("update ArticleClass set OrderID=" & PrevOrderID & "+1 where ClassID=" & ClassID)
	end if
	
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub

sub SaveModify()
	dim ClassName,Readme,IsElite,ShowOnTop,Setting,ClassMaster,ClassPicUrl,LinkUrl,SkinID,LayoutID,BrowsePurview,AddPurview
	dim trs,rs
	dim ClassID,sql,rsClass,i
	dim SkinCount,LayoutCount
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		ClassID=CLng(ClassID)
	end if
	ClassName=trim(request("ClassName"))
	IsElite=trim(request("IsElite"))
	ShowOnTop=trim(request("ShowOnTop"))
	Setting=trim(request("Setting"))
	ClassMaster=trim(request("ClassMaster"))
	Readme=trim(request("Readme"))
	ClassPicUrl=trim(request("ClassPicUrl"))
	LinkUrl=trim(request("LinkUrl"))
	SkinID=Trim(request("SkinID"))
	LayoutID=trim(request("LayoutID"))
	BrowsePurview=trim(request("BrowsePurview"))
	AddPurview=trim(request("AddPurview"))
	if ClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目名称不能为空！</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	sql="select * From ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	if rsClass("Child")>0 and LinkUrl<>"" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>本栏目有子栏目，所以不能设为外部链接地址。</li>"
	end if
	if IsElite="Yes" then
		IsElite=True
	else
		IsElite=False
	end if
	if ShowOnTop="Yes" then
		ShowOnTop=True
	else
		ShowOnTop=False
	end if
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定栏目配色模板</li>"
	else
		SkinID=Clng(SkinID)
	end if
	if LayoutID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定版面设计模板</li>"
	else
		LayoutID=CLng(LayoutID)
	end if

	if ClassMaster<>"" and ClassMaster<>rsClass("ClassMaster") then
		'call AddMaster(ClassMaster)
	end if
	
	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
   	rsClass("ClassName")=ClassName
	rsClass("IsElite")=IsElite
	rsClass("ShowOnTop")=ShowOnTop
	rsClass("Setting")=Clng(Setting)
	'rsClass("ClassMaster")=ClassMaster
	rsClass("Readme")=Readme
	rsClass("ClassPicUrl")=ClassPicUrl
	rsClass("LinkUrl")=LinkUrl
	rsClass("SkinID")=SkinID
	rsClass("LayoutID")=LayoutID
	rsClass("BrowsePurview")=Cint(BrowsePurview)
	rsClass("AddPurview")=Cint(AddPurview)
	rsClass.update
	rsClass.close
	set rsClass=nothing
	
	set rs=nothing
	set trs=nothing
	
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub


sub DeleteClass()
	dim sql,rs,PrevID,NextID,ClassID
	ClassID=trim(Request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select ClassID,RootID,Depth,ParentID,Child,PrevID,NextID From ArticleClass where ClassID="&ClassID
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目不存在，或者已经被删除</li>"
	else
		if rs("Child")>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>该栏目含有子栏目，请删除其子栏目后再进行删除本栏目的操作</li>"
		end if
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	if rs("Depth")>0 then
		conn.execute("update ArticleClass set child=child-1 where ClassID=" & rs("ParentID"))
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'删除本栏目的所有文章和评论
	conn.execute("delete from Article where ClassID=" & ClassID)
	conn.execute("delete from ArticleComment where ClassID=" & ClassID)
	
	'修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
		
end sub

sub ClearClass()
	dim strClassID,rs,trs,SuccessMsg,ClassID
	ClassID=trim(Request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	strClassID=cstr(ClassID)
	set rs=conn.execute("select ClassID,Child,ParentPath from ArticleClass where ClassID=" & ClassID)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目不存在，或者已经被删除</li>"
		exit sub
	end if
	if rs(1)>0 then
		set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & rs(0))
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=conn.execute("select ClassID from ArticleClass where ParentPath like '" & rs(2) & "," & rs(0) & ",%'")
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=nothing
	end if
	rs.close
	set rs=nothing
	conn.execute("update Article set Deleted=True where ClassID in (" & strClassID & ")")
	conn.execute("delete from Article where ClassID in (" & strClassID & ")")	
	SuccessMsg="此栏目（包括子栏目）的所有文章已经被移到回收站中！"
	call WriteSuccessMsg(SuccessMsg)
end sub


sub SaveMove()
	dim ClassID,sql,rsClass,i
	dim rParentID
	dim trs,rs
	dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID
	ClassID=trim(request("ClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		ClassID=CLng(ClassID)
	end if
	
	sql="select * From ArticleClass where ClassID=" & ClassID
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	rParentID=trim(request("ParentID"))
	if rParentID="" then
		rParentID=0
	else
		rParentID=CLng(rParentID)
	end if
	
	if rsClass("ParentID")<>rParentID then   '更改了所属栏目，则要做一系列检查
		if rParentID=rsClass("ClassID") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属栏目不能为自己！</li>"
		end if
		'判断所指定的栏目是否为外部栏目或本栏目的下属栏目
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=conn.execute("select rootid From ArticleClass where LinkUrl='' and ClassID="&rParentID)
				if trs.bof and trs.eof then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>不能指定外部栏目为所属栏目</li>"
				else
					if rsClass("rootid")=trs(0) then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>不能指定该栏目的下属栏目作为所属栏目</li>"
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=conn.execute("select ClassID From ArticleClass where ParentPath like '"&rsClass("ParentPath")&"," & rsClass("ClassID") & "%' and ClassID="&rParentID)
			if not (trs.eof and trs.bof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>您不能指定该栏目的下属栏目作为所属栏目</li>"
			end if
			trs.close
			set trs=nothing
		end if
		
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
	if rsClass("ParentID")=0 then
		ParentID=rsClass("ClassID")
		iParentID=0
	else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	end if
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.close
	set rsClass=nothing
	
	
  '假如更改了所属栏目
  '需要更新其原来所属栏目信息，包括深度、父级ID、栏目数、排序、继承版主等数据
  '需要更新当前所属栏目信息
  '继承版主数据需要另写函数进行更新--取消，在前台可用ClassID in ParentPath来获得
  dim mrs,MaxRootID
  set mrs=conn.execute("select max(rootid) From ArticleClass")
  MaxRootID=mrs(0)
  set mrs=nothing
  if isnull(MaxRootID) then
	MaxRootID=0
  end if
  dim k,nParentPath,mParentPath
  dim ParentSql,ClassCount
  dim rsPrevOrderID
  if clng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '假如更改了所属栏目
	'更新原来同一父栏目的上一个栏目的NextID和下一个栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	if iParentID>0 and rParentID=0 then  	'如果原来不是一级分类改成一级分类
		'得到上一个一级分类栏目
		sql="select ClassID,NextID from ArticleClass where RootID=" & MaxRootID & " and Depth=0"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '得到新的PrevID
		rs(1)=ClassID     '更新上一个一级分类栏目的NextID的值
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'更新当前栏目数据
		conn.execute("update ArticleClass set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID)
		'如果有下属栏目，则更新其下属栏目数据。下属栏目的排序不需考虑，只需更新下属栏目深度和一级排序ID(rootid)数据
		if child>0 then
			i=0
			ParentPath=ParentPath & ","
			set rs=conn.execute("select * From ArticleClass where ParentPath like '%"&ParentPath & ClassID&"%'")
			do while not rs.eof
				i=i+1
				mParentPath=replace(rs("ParentPath"),ParentPath,"")
				conn.execute("update ArticleClass set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' where ClassID="&rs("ClassID"))
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		
		'更新其原来所属栏目的栏目数，排序相当于剪枝而不需考虑
		conn.execute("update ArticleClass set child=child-1 where ClassID="&iParentID)
		
	elseif iParentID>0 and rParentID>0 then    '如果是将一个分栏目移动到其他分栏目下
		'得到当前栏目的下属子栏目数
		ParentPath=ParentPath & ","
		set rs=conn.execute("select count(*) From ArticleClass where ParentPath like '%"&ParentPath & ClassID&"%'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'获得目标栏目的相关信息		
		set trs=conn.execute("select * From ArticleClass where ClassID="&rParentID)
		if trs("Child")>0 then		
			'得到与本栏目同级的最后一个栏目的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentID=" & trs("ClassID"))
			PrevOrderID=rsPrevOrderID(0)
			'得到与本栏目同级的最后一个栏目的ClassID
			sql="select ClassID,NextID from ArticleClass where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '得到新的PrevID
			rs(1)=ClassID     '更新上一个栏目的NextID的值
			rs.update
			rs.close
			set rs=nothing
			
			'得到同一父栏目但比本栏目级数大的子栏目的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentPath like '" & trs("ParentPath") & "," & trs("ClassID") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
		
		'在获得移动过来的栏目数后更新排序在指定栏目之后的栏目排序数据
		conn.execute("update ArticleClass set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		'更新当前栏目数据
		conn.execute("update ArticleClass set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("ClassID") & "',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID)
		
		'如果有子栏目则更新子栏目数据，深度为原来的相对深度加上当前所属栏目的深度
		set rs=conn.execute("select * From ArticleClass where ParentPath like '%"&ParentPath&ClassID&"%' order by OrderID")
		i=1
		do while not rs.eof
			i=i+1
			iParentPath=trs("ParentPath") & "," & trs("ClassID") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.execute("update ArticleClass set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' where ClassID="&rs("ClassID"))
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		
		'更新所指向的上级栏目的子栏目数
		conn.execute("update ArticleClass set child=child+1 where ClassID="&rParentID)
		
		'更新其原父类的子栏目数			
		conn.execute("update ArticleClass set child=child-1 where ClassID="&iParentID)
	else    '如果原来是一级栏目改成其他栏目的下属栏目
		'得到移动的栏目总数
		set rs=conn.execute("select count(*) From ArticleClass where rootid="&rootid)
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'获得目标栏目的相关信息		
		set trs=conn.execute("select * From ArticleClass where ClassID="&rParentID)
		if trs("Child")>0 then		
			'得到与本栏目同级的最后一个栏目的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentID=" & trs("ClassID"))
			PrevOrderID=rsPrevOrderID(0)
			sql="select ClassID,NextID from ArticleClass where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=ClassID
			rs.update
			set rs=nothing
			
			'得到同一父栏目但比本栏目级数大的子栏目的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From ArticleClass where ParentPath like '" & trs("ParentPath") & "," & trs("ClassID") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
	
		'在获得移动过来的栏目数后更新排序在指定栏目之后的栏目排序数据
		conn.execute("update ArticleClass set OrderID=OrderID+" & ClassCount &"+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		conn.execute("update ArticleClass set PrevID=" & PrevID & ",NextID=0 where ClassID=" & ClassID)
		set rs=conn.execute("select * From ArticleClass where rootid="&rootid&" order by OrderID")
		i=0
		do while not rs.eof
			i=i+1
			if rs("parentid")=0 then
				ParentPath=trs("ParentPath") & "," & trs("ClassID")
				conn.execute("update ArticleClass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" where ClassID="&rs("ClassID"))
			else
				ParentPath=trs("ParentPath") & "," & trs("ClassID") & "," & replace(rs("ParentPath"),"0,","")
				conn.execute("update ArticleClass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' where ClassID="&rs("ClassID"))
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		'更新所指向的上级栏目栏目数		
		conn.execute("update ArticleClass set child=child+1 where ClassID="&rParentID)

	end if
  end if
	
  call CloseConn()
  Response.Redirect "Admin_Class_Article.asp"  
end sub

sub UpOrder()
	dim ClassID,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=trim(request("ClassID"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=Cint(cRootID)
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

	'得到本栏目的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID from ArticleClass where ClassID=" & ClassID)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From ArticleClass")
	MaxRootID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update ArticleClass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前栏目以上的栏目的RootID依次加一，范围为要提升的数字
	sqlOrder="select * From ArticleClass where ParentID=0 and RootID<" & cRootID & " order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子栏目
		conn.execute("update ArticleClass set RootID=RootID+1 where RootID=" & tRootID)
		i=i+1
		if i>MoveNum then
			rsOrder("PrevID")=ClassID
			rsOrder.update
			conn.execute("update ArticleClass set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update ArticleClass set PrevID=0 where ClassID=" & ClassID)
	else
		rsOrder("NextID")=ClassID
		rsOrder.update
		conn.execute("update ArticleClass set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update ArticleClass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=Order"
end sub

sub DownOrder()
	dim ClassID,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=trim(request("ClassID"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=Cint(cRootID)
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

	'得到本栏目的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID from ArticleClass where ClassID=" & ClassID)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From ArticleClass")
	MaxRootID=mrs(0)+1
	'先将当前栏目移至最后，包括子栏目
	conn.execute("update ArticleClass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前栏目以下的栏目的RootID依次减一，范围为要下降的数字
	sqlOrder="select * From ArticleClass where ParentID=0 and RootID>" & cRootID & " order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前栏目已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子栏目
		conn.execute("update ArticleClass set RootID=RootID-1 where RootID=" & tRootID)
		i=i+1
		if i>MoveNum then
			rsOrder("NextID")=ClassID
			rsOrder.update
			conn.execute("update ArticleClass set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update ArticleClass set NextID=0 where ClassID=" & ClassID)
	else
		rsOrder("PrevID")=ClassID
		rsOrder.update
		conn.execute("update ArticleClass set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前栏目从最后移到相应位置，包括子栏目
	conn.execute("update ArticleClass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=Order"
end sub

sub UpOrderN()
	dim sqlOrder,rsOrder,MoveNum,ClassID,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=Trim(request("ClassID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		ClassID=CLng(ClassID)
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

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的栏目信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From ArticleClass where ClassID="&ClassID)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	if child>0 then
		set rs=conn.execute("select count(*) From ArticleClass where ParentPath like '%"&ParentPath&"%'")
		oldorders=rs(0)
		rs.close
		set rs=nothing
	else
		oldorders=0
	end if
	'先修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'和该栏目同级且排序在其之上的栏目------更新其排序，范围为要提升的数字
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From ArticleClass where ParentID="&ParentID&" and OrderID<"&OrderID&" order by OrderID desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=1
	do while not rs.eof
		tOrderID=rs(1)
		conn.execute("update ArticleClass set OrderID="&tOrderID+oldorders+i&" where ClassID="&rs(0))
		if rs(2)>0 then
			ii=i+1
			set trs=conn.execute("select ClassID,OrderID From ArticleClass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					conn.execute("update ArticleClass set OrderID="&tOrderID+oldorders+ii&" where ClassID="&trs(0))
					ii=ii+1
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		i=i+1
		if i>MoveNum then
			rs(4)=ClassID
			rs.update
			conn.execute("update ArticleClass set NextID=" & rs(0) & " where ClassID=" & ClassID)		
			exit do
		end if
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update ArticleClass set PrevID=0 where ClassID=" & ClassID)
	else
		rs(5)=ClassID
		rs.update
		conn.execute("update ArticleClass set PrevID=" & rs(0) & " where ClassID=" & ClassID)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的栏目的序号
	conn.execute("update ArticleClass set OrderID="&tOrderID&" where ClassID="&ClassID)
	'如果有下属栏目，则更新其下属栏目排序
	if child>0 then
		i=1
		set rs=conn.execute("select ClassID From ArticleClass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update ArticleClass set OrderID="&tOrderID+i&" where ClassID="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=OrderN"
end sub

sub DownOrderN()
	dim sqlOrder,rsOrder,MoveNum,ClassID,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=Trim(request("ClassID"))
	MoveNum=trim(request("MoveNum"))
	if ClassID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		ClassID=Cint(ClassID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要下降的数字！</li>"
			exit sub
		end if
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的栏目信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From ArticleClass where ClassID="&ClassID)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing

	'先修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'和该栏目同级且排序在其之下的栏目------更新其排序，范围为要下降的数字
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From ArticleClass where ParentID="&ParentID&" and OrderID>"&OrderID&" order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      '同级栏目
	ii=0     '同级栏目和子栏目
	do while not rs.eof
		conn.execute("update ArticleClass set OrderID="&OrderID+ii&" where ClassID="&rs(0))
		if rs(2)>0 then
			set trs=conn.execute("select ClassID,OrderID From ArticleClass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					ii=ii+1
					conn.execute("update ArticleClass set OrderID="&OrderID+ii&" where ClassID="&trs(0))
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		ii=ii+1
		i=i+1
		if i>=MoveNum then
			rs(5)=ClassID
			rs.update
			conn.execute("update ArticleClass set PrevID=" & rs(0) & " where ClassID=" & ClassID)		
			exit do
		end if
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update ArticleClass set NextID=0 where ClassID=" & ClassID)
	else
		rs(4)=ClassID
		rs.update
		conn.execute("update ArticleClass set NextID=" & rs(0) & " where ClassID=" & ClassID)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的栏目的序号
	conn.execute("update ArticleClass set OrderID="&OrderID+ii&" where ClassID="&ClassID)
	'如果有下属栏目，则更新其下属栏目排序
	if child>0 then
		i=1
		set rs=conn.execute("select ClassID From ArticleClass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update ArticleClass set OrderID="&OrderID+ii+i&" where ClassID="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	call CloseConn()
	response.Redirect "Admin_Class_Article.asp?Action=OrderN"
end sub

sub SaveReset()
	dim i,sql,rs,SuccessMsg,iCount,PrevID,NextID
	sql="select ClassID From ArticleClass order by RootID,OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	iCount=rs.recordcount
	i=1
	PrevID=0
	do while not rs.eof
		rs.movenext
		if rs.eof then
			NextID=0
		else
			NextID=rs(0)
		end if
		rs.moveprevious
		conn.execute("update ArticleClass set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & " where ClassID=" & rs(0))
		PrevID=rs(0)
		i=i+1
		rs.movenext
	loop
	rs.close
	set rs=nothing	
	
	SuccessMsg="复位成功！请返回<a href='Admin_Class_Article.asp'>栏目管理首页</a>做栏目的归属设置。"
	call WriteSuccessMsg(SuccessMsg)
end sub

sub SaveUnite()
	dim ClassID,TargetClassID,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i,SuccessMsg
	ClassID=trim(request("ClassID"))
	TargetClassID=trim(request("TargetClassID"))
	if ClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要合并的栏目！</li>"
	else
		ClassID=CLng(ClassID)
	end if
	if TargetClassID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标栏目！</li>"
	else
		TargetClassID=CLng(TargetClassID)
	end if
	if ClassID=TargetClassID then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请不要在相同栏目内进行操作</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	'判断目标栏目是否有子栏目，如果有，则报错。
	set rs=conn.execute("select Child from ArticleClass where ClassID=" & TargetClassID)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>目标栏目不存在，可能已经被删除！</li>"
	else
		if rs(0)>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>目标栏目中含有子栏目，不能合并！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到当前栏目信息
	set rs=conn.execute("select ClassID,ParentID,ParentPath,PrevID,NextID,Depth from ArticleClass where ClassID="&ClassID)
	iParentID=rs(1)
	Depth=rs(5)
	if iParentID=0 then
		ParentPath=rs(0)
	else
		ParentPath=rs(2) & "," & rs(0)
	end if
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'判断是否是合并到其下属栏目中
	set rs=conn.execute("select ClassID from ArticleClass where ClassID="&TargetClassID&" and ParentPath like '"&ParentPath&"%'")
	if not (rs.eof and rs.bof) then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>不能将一个栏目合并到其下属子栏目中</li>"
		exit sub
	end if
	
	'得到当前栏目的下属栏目ID
	set rs=conn.execute("select ClassID from ArticleClass where ParentPath like '"&ParentPath&"%'")
	i=0
	if not (rs.eof and rs.bof) then
		do while not rs.eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.movenext
		loop
	end if
	if i>0 then
		ParentPath=iParentPath
	else
		ParentPath=ClassID
	end if
	
	'先修改上一栏目的NextID和下一栏目的PrevID
	if PrevID>0 then
		conn.execute "update ArticleClass set NextID=" & NextID & " where ClassID=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update ArticleClass set PrevID=" & PrevID & " where ClassID=" & NextID
	end if
	
	'更新文章及评论所属栏目
	conn.execute("update Article set ClassID="&TargetClassID&" where ClassID in ("&ParentPath&")")
	conn.execute("update ArticleComment set ClassID="&TargetClassID&" where ClassID in ("&ParentPath&")")
	
	'删除被合并栏目及其下属栏目
	conn.execute("delete from ArticleClass where ClassID in ("&ParentPath&")")
	
	'更新其原来所属栏目的子栏目数，排序相当于剪枝而不需考虑
	if Depth>0 then
		conn.execute("update ArticleClass set Child=Child-1 where ClassID="&iParentID)
	end if
	
	SuccessMsg="栏目合并成功！已经将被合并栏目及其下属子栏目的所有数据转入目标栏目中。<br><br>同时删除了被合并的栏目及其子栏目。"
	call WriteSuccessMsg(SuccessMsg)
	set rs=nothing
	set trs=nothing
end sub

%> 
