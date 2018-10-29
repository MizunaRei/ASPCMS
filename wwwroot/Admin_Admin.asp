<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1
'response.write "此功能被WEBBOY暂时禁止了！"
'response.end
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!-- 两课网站所用函数-->
<!--#include file="inc/admin_code_article.asp"-->
<%
dim rs, sql, strPurview,iCount
dim Action,FoundErr,ErrMsg
Action=Trim(request("Action"))
%>
<html>
<head>
<title>管理员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>

//显示添加管理员的详细个人信息
function ModifyAdminPurview1()
{
	PurviewDetail.style.display='none';
	StudentAdminPurviewDetail.style.display='none';
	AdminPurviewDetail2.style.display='none';

}
function ModifyAdminPurview2()
{
	PurviewDetail.style.display='';
	StudentAdminPurviewDetail.style.display='none';
	AdminPurviewDetail2.style.display='';
}
function ModifyAdminPurview3()
{
	PurviewDetail.style.display='';
	StudentAdminPurviewDetail.style.display='';
	AdminPurviewDetail2.style.display='none';
}


function AddAdminPurview1()
{	
	PurviewDetail.style.display='none';
	StudentAdminPurviewDetail.style.display='none';
	AdminPurviewDetail2.style.display='none';
}
function AddAdminPurview2()
{
	PurviewDetail.style.display='';
	StudentAdminPurviewDetail.style.display='none';
	AdminPurviewDetail2.style.display='';
}


function AddAdminPurview3()
{
	PurviewDetail.style.display='';
	StudentAdminPurviewDetail.style.display='';
	AdminPurviewDetail2.style.display='none';
}
//结束两课网站代码
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
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}

function CheckAdd()
{
  if(document.form1.username.value=="")
    {
      alert("用户名不能为空！");
	  document.form1.username.focus();
      return false;
    }
    
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
  if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}

function CheckModifyPurview()
{
  if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
}

function GetClassPurview()
{
    document.form1.ClassInputer_Article.value="";
	document.form1.ClassChecker_Article.value="";
	document.form1.ClassMaster_Article.value="";
	if(document.form1.AdminPurview_Article[2].checked==true){
		for(var i=0;i<frmArticle.document.myform.Purview_Add.length;i++){
			if (frmArticle.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Article.value=="")
					document.form1.ClassInputer_Article.value=frmArticle.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Article.value+=","+frmArticle.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmArticle.document.myform.Purview_Check.length;i++){
			if (frmArticle.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Article.value=="")
					document.form1.ClassChecker_Article.value=frmArticle.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Article.value+=","+frmArticle.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmArticle.document.myform.Purview_Manage.length;i++){
			if (frmArticle.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Article.value=="")
					document.form1.ClassMaster_Article.value=frmArticle.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Article.value+=","+frmArticle.document.myform.Purview_Manage[i].value;
			}
		}
	}
    document.form1.ClassInputer_Soft.value="";
	document.form1.ClassChecker_Soft.value="";
	document.form1.ClassMaster_Soft.value="";
	if(document.form1.AdminPurview_Soft[2].checked==true){
		for(var i=0;i<frmSoft.document.myform.Purview_Add.length;i++){
			if (frmSoft.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Soft.value=="")
					document.form1.ClassInputer_Soft.value=frmSoft.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Soft.value+=","+frmSoft.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmSoft.document.myform.Purview_Check.length;i++){
			if (frmSoft.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Soft.value=="")
					document.form1.ClassChecker_Soft.value=frmSoft.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Soft.value+=","+frmSoft.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmSoft.document.myform.Purview_Manage.length;i++){
			if (frmSoft.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Soft.value=="")
					document.form1.ClassMaster_Soft.value=frmSoft.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Soft.value+=","+frmSoft.document.myform.Purview_Manage[i].value;
			}
		}
	}
	document.form1.ClassInputer_Photo.value="";
	document.form1.ClassChecker_Photo.value="";
	document.form1.ClassMaster_Photo.value="";
    if(document.form1.AdminPurview_Photo[2].checked==true){
		for(var i=0;i<frmPhoto.document.myform.Purview_Add.length;i++){
			if (frmPhoto.document.myform.Purview_Add[i].checked==true){
				if (document.form1.ClassInputer_Photo.value=="")
					document.form1.ClassInputer_Photo.value=frmPhoto.document.myform.Purview_Add[i].value;
				else
					document.form1.ClassInputer_Photo.value+=","+frmPhoto.document.myform.Purview_Add[i].value;
			}
		}
		for(var i=0;i<frmPhoto.document.myform.Purview_Check.length;i++){
			if (frmPhoto.document.myform.Purview_Check[i].checked==true){
				if (document.form1.ClassChecker_Photo.value=="")
					document.form1.ClassChecker_Photo.value=frmPhoto.document.myform.Purview_Check[i].value;
				else
					document.form1.ClassChecker_Photo.value+=","+frmPhoto.document.myform.Purview_Check[i].value;
			}
		}
		for(var i=0;i<frmPhoto.document.myform.Purview_Manage.length;i++){
			if (frmPhoto.document.myform.Purview_Manage[i].checked==true){
				if (document.form1.ClassMaster_Photo.value=="")
					document.form1.ClassMaster_Photo.value=frmPhoto.document.myform.Purview_Manage[i].value;
				else
					document.form1.ClassMaster_Photo.value+=","+frmPhoto.document.myform.Purview_Manage[i].value;
			}
		}
	}
}
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22" colspan="2" align="center"><strong>管 理 员 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Admin.asp">管理员管理首页</a>&nbsp;|&nbsp;<a href="Admin_Admin.asp?Action=Add">新增管理员</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddAdmin()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="ModifyPwd" then
	call ModifyPwd()
elseif Action="ModifyPurview" then
	call ModifyPurview()
elseif Action="SaveModifyPwd" then
	call SaveModifyPwd()
elseif Action="SaveModifyPurview" then
	call SaveModifyPurview()
elseif Action="Del" then
	call DelAdmin()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
if AdminPurview=1 then
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from admin order by id"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<br>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
    <form name="myform" method="Post" action="Admin_Admin.asp" onSubmit="return confirm('确定要删除选中的管理员吗？');">
      <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
          <tr align="center" class="title">
            <td  width="30"><strong>选中</strong></td>
            <td  width="30" height="22"><strong> 序号</strong></td>
            <td height="22"><strong> 用 户 名</strong></td>
            <td  width="100" height="22"><strong> 权 限</strong></td>
            <td width="100"><strong>最后登录IP</strong></td>
            <td width="120"><strong>最后登录时间</strong></td>
            <td  width="60"><strong>登录次数</strong></td>
            <td  width="150" height="22"><strong> 操 作</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td width="30"><input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" <%if rs("UserName")=AdminName then response.write " disabled"%> onClick="unselectall()"></td>
            <td width="30"><%=rs("ID")%></td>
            <td><%
	if rs("username")=AdminName then
		response.write "<font color=red><b>" & rs("UserName") & "</b></font>"
	else
		response.write rs("UserName")
	end if
	%></td>
            <td width="100"><%
		  select case rs("purview")
		    case 1
              strPurview="<font color=blue>超级管理员</font>"
            case 2
              strpurview="教师管理员"
			 case 3
			 	strpurview="学生管理员"
		  end select
		  response.write(strPurview)
         %>
            </td>
            <td width="100"><%
	if rs("LastLoginIP")<>"" then
		response.write rs("LastLoginIP")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="120"><%
	if rs("LastLoginTime")<>"" then
		response.write rs("LastLoginTime")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="60"><%
	if rs("LoginTimes")<>"" then
		response.write rs("LoginTimes")
	else
		response.write "0"
	end if
	%>
            </td>
            <td width="150"><%
	response.write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rs("ID") & "'>修改密码</a>&nbsp;&nbsp;"
	response.write "<a href='Admin_Admin.asp?Action=ModifyPurview&ID=" & rs("ID") & "'>修改权限</a>&nbsp;&nbsp;"
	if iCount>1 and rs("UserName")<>AdminName then
		response.write "<a href='Admin_Admin.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('确定要删除此管理员吗？');"">删除</a>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	%>
            </td>
          </tr>
          <%
	rs.MoveNext
	loop
  %>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有管理员</td>
            <td><input name="Action" type="hidden" id="Action" value="Del">
              <input name="Submit" type="submit" id="Submit" value="删除选中的管理员"></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<%
	rs.Close
	set rs=Nothing
end if
if AdminPurview=2 then
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from admin where Purview=3 and TeacherName='"  &  session("AdminTeacherName") & "' order by id"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<br>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
    <form name="myform" method="Post" action="Admin_Admin.asp" onSubmit="return confirm('确定要删除选中的管理员吗？');">
      <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
          <tr align="center" class="title">
            <td  width="30"><strong>选中</strong></td>
            <td  width="30" height="22"><strong> 序号</strong></td>
            <td height="22"><strong> 用 户 名</strong></td>
            <td  width="100" height="22"><strong> 权 限</strong></td>
            <td width="100"><strong>最后登录IP</strong></td>
            <td width="120"><strong>最后登录时间</strong></td>
            <td  width="60"><strong>登录次数</strong></td>
            <td  width="150" height="22"><strong> 操 作</strong></td>
          </tr>
          <%do while not rs.EOF %>
          <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
            <td width="30"><input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" <%if rs("UserName")=AdminName then response.write " disabled"%> onClick="unselectall()"></td>
            <td width="30"><%=rs("ID")%></td>
            <td><%
	if rs("username")=AdminName then
		response.write "<font color=red><b>" & rs("UserName") & "</b></font>"
	else
		response.write rs("UserName")
	end if
	%></td>
            <td width="100"><%
		  select case rs("purview")
		    case 1
              strPurview="<font color=blue>超级管理员</font>"
            case 2
              strpurview="教师管理员"
			 case 3
			 	strpurview="学生管理员"
		  end select
		  response.write(strPurview)
         %>
            </td>
            <td width="100"><%
	if rs("LastLoginIP")<>"" then
		response.write rs("LastLoginIP")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="120"><%
	if rs("LastLoginTime")<>"" then
		response.write rs("LastLoginTime")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="60"><%
	if rs("LoginTimes")<>"" then
		response.write rs("LoginTimes")
	else
		response.write "0"
	end if
	%>
            </td>
            <td width="150"><%
	response.write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rs("ID") & "'>修改密码</a>&nbsp;&nbsp;"
	response.write "<a href='Admin_Admin.asp?Action=ModifyPurview&ID=" & rs("ID") & "'>修改权限</a>&nbsp;&nbsp;"
	if iCount>1 and rs("UserName")<>AdminName then
		response.write "<a href='Admin_Admin.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('确定要删除此管理员吗？');"">删除</a>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	%>
            </td>
          </tr>
          <%
	rs.MoveNext
	loop
  %>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有管理员</td>
            <td><input name="Action" type="hidden" id="Action" value="Del">
              <input name="Submit" type="submit" id="Submit" value="删除选中的管理员"></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<%
	rs.Close
	set rs=Nothing
end if

end sub
'end sub main()

sub AddAdmin()
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckAdd();">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title">
      <td height="22" colspan="2"><div align="center"><strong>新 增 管 理 员</strong></div></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong> 用 户 名：</strong></td>
      <td width="65%" class="tdbg"><font size="2"><input name="username" type="text" maxlength="16"></font>
        </td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong> 初始密码： </strong></td>
      <td width="65%" class="tdbg"><font size="3">
        <input type="password" name="Password">
        </font></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong> 确认密码：</strong></td>
      <td width="65%" class="tdbg"><font size="3">
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
    <tr class="tdbg">
      <td width="35%" class="tdbg"><strong>权限设置： </strong></td>
      <td width="65%" class="tdbg"><table width="100%" border="0" cellspacing="1" cellpadding="2">
          
       <% if AdminPurview=1 then %> 
	    <tr>
            <td width="100"><input name="Purview" type="radio" value="1" checked onClick=AddAdminPurview1() >
              超级管理员：</td>
            <td>拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有</td>
          </tr> 
		  
          <tr>
            <td width="100"><input type="radio" name="Purview" value="2" onClick=AddAdminPurview2() >
              教师管理员：</td>
            <td>需要详细指定每一项管理权限</td>
          </tr>
		 <% end if %>
          <tr>
            <td width="100"><input type="radio" name="Purview" id="Purview3" value="3" onClick=AddAdminPurview3() >
              学生管理员：</td>
            <td>需要详细指定每一项管理权限</td>
          </tr>
        </table></td>
    </tr>
    <tr class="tdbg">
      <td colspan="2">
      <table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" style="display:none">
          <tr>
            <td colspan="2" align="center"><strong>管 理 员 权 限 详 细 设 置</strong></td>
          </tr>
          <tr valign="top">
          
          
                    <!-- 两课网站管理信息修改-->
          <tr>
            <td class="tdbg"  width="100%"  colspan="9">
            <fieldset>
              <legend>个人信息</legend>
              <table  width="100%">
            <tr>
            <td   ><strong>真实姓名：</strong>&nbsp;&nbsp; 
            <input name=TrueName type=text   value=""></td>
          </tr>  
         
         <tr><td>
          <%  call Admin_AddStudentAdminInformation() 

		 %>
         </td></tr></table></fieldset>
         <!--开始课程列表-->
         <tr><td colspan="9">
            <fieldset>
            <legend>管理的课程:</legend>
            <table><tbody><tr><td>
            <% 			  					
'	课程列表		  Response.Write("123456")
'		  Response.End()
			call  Admin_Add_ShowSpecial_Checkbox() 

			  
			  %>
              </td></tr></tbody></table>
            </fieldset>
          </td></tr><!--结束课程列表-->
          <!-- 结束两课网站管理员信息修改-->

          
          
          </td></tr>
           <tr> <td width="100%"  colspan="9">
            <fieldset>
              <legend>留言板</legend>
              <table width="100%" border="0" cellspacing="1" cellpadding="2">
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Reply">
                    回复留言 </td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Modify">
                    修改留言</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Del">
                    删除留言</td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Check">
                    审核留言</td>
                </tr>
              </table>
              </fieldset></td></tr>
              <!--<br>-->
             <tr> <td  width="100%" colspan="9">
             <table width="100%" border="0" cellspacing="1" cellpadding="2" id="AdminPurviewDetail2" style="display:none" ><tbody >
             
             <legend>网站管理权限</legend>
              <!--<table width="100%" border="0" cellspacing="1" cellpadding="2">-->
                
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Announce">
                    公告管理</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="FriendSite">
                    网站链接管理</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Vote">
                    调查管理&nbsp;</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Counter">
                    统计管理</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="User">
                    注册用户管理</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="MailList">
                    邮件列表管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Skin">
                    配色模板管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Layout">
                    版面设计模板管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="JS">
                    JS代码管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Database" disabled>
                    数据库管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="UpFile">
                    上传文件管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="ModifyPwd" checked>
                    修改自己的登录密码</td>
                </tr>
              <!--</table>-->
              </tbody>
              </table></td></tr>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td height="40" colspan="9" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input  type="submit" name="Submit" value=" 添 加 " style="cursor:hand;">
        &nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
end sub

sub ModifyPwd()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
	else
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckModifyPwd();">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title">
      <td height="22" colspan="2"><div align="center"><font size="2"><strong>修 改 管 理 员 密 码</strong></font></div></td>
    </tr>
    <tr>
      <td width="40%" class="tdbg"><strong>用 户 名：</strong></td>
      <td width="60%" class="tdbg"><%=rs("UserName")%>
        <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr>
      <td width="40%" class="tdbg"><strong>新 密 码：</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="Password">
      </td>
    </tr>
    <tr>
      <td width="40%" class="tdbg"><strong>确认密码：</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="PwdConfirm">
      </td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveModifyPwd">
        <input  type="submit" name="Submit" value="保存修改结果" style="cursor:hand;">
        &nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub ModifyPurview()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
	else
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:CheckModifyPurview();">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title">
      <td height="22" colspan="2"><div align="center"><font size="2"><strong>修 改 管 理 员 详 细 设 置</strong></font></div></td>
    </tr>
    <tr>
      <td width="300" class="tdbg"><strong>用 户 名：</strong></td>
      <td class="tdbg"><%=rs("UserName")%>
        <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr class="tdbg">
      <td width="25%" class="tdbg"><strong>权限设置： </strong></td>
      <td width="65%" class="tdbg"><table width="100%" border="0" cellspacing="1" cellpadding="2">
         <% if AdminPurview=1 then %>
		  <tr>
            <td width="100"><input name="Purview" type="radio" value="1" onClick=ModifyAdminPurview1() <%if rs("Purview")=1 then response.write "checked"%>>
              超级管理员：</td>
            <td>拥有所有权限。某些权限（如管理员管理、网站信息配置、网站选项配置等管理权限）只有超级管理员才有</td>
          </tr>
          <!--教师管理员选项-->
          <tr>
            <td width="100"><input type="radio" name="Purview" value="2" onClick=ModifyAdminPurview2() <%if rs("Purview")=2 then response.write "checked"%>>
              教师管理员：</td>
            <td>需要详细指定每一项管理权限</td>
          </tr>
          <!--结束教师管理员选项-->
          <% end if %>
          <!--学生管理员选项-->
         
          <tr>
            <td width="100"><input type="radio" name="Purview" value="3" onClick=ModifyAdminPurview3() <%if rs("Purview")=3 then response.write "checked"%>>
              学生管理员：</td>
            <td>需要详细指定每一项管理权限</td>
          </tr>
          <!--结束学生管理员选项-->
        </table></td>
    </tr>
    <tr class="tdbg">
      <td colspan="9">
      <table id="PurviewDetail" width="100%" border="0" cellspacing="10" cellpadding="0" <%if rs("Purview")=1 then response.write "style='display:none'"%>>
          <tr>
            <td colspan="2" align="center"><strong>管 理 员 详 细 设 置</strong></td>
          </tr>
          <!-- 两课网站管理信息修改-->
          <tr>
            <td class="tdbg" ><strong>真实姓名：</strong>&nbsp;&nbsp;&nbsp;&nbsp; 
            <input name=TrueName type=text   value="<%=rs("TrueName")%>"></td>
          </tr>
          <tr><td >
          <%  call Admin_ModifyStudentAdminInformation() 

		 %>
         </td></tr>
          <!-- 结束两课网站管理员信息修改-->
          <!--开始课程列表--><tr><td colspan="9">
            <fieldset>
            <legend>管理的课程:</legend>
            <table><tbody><tr><td>
            <% 			  					
'	课程列表		  Response.Write("123456")
'		  Response.End()
			call  Admin_ShowSpecial_Checkbox(rs("ID")) 

			  
			  %>
              </td></tr></tbody></table>
            </fieldset>
          </td></tr><!--结束课程列表-->
          <tr valign="top">
            <td width="100%" colspan="9">
            <fieldset>
              <legend>留言板</legend>
              <table width="100%" border="0" cellspacing="1" cellpadding="2">
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Reply" <%if CheckPurview(rs("AdminPurview_Guest"),"Reply")=True then response.write "checked"%>>
                    回复留言 </td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Modify" <%if CheckPurview(rs("AdminPurview_Guest"),"Modify")=True then response.write "checked"%>>
                    修改留言</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Del" <%if CheckPurview(rs("AdminPurview_Guest"),"Del")=True then response.write "checked"%>>
                    删除留言</td>
                  <td width="50%"><input name="AdminPurview_Guest" type="checkbox" id="AdminPurview_Guest" value="Check" <%if CheckPurview(rs("AdminPurview_Guest"),"Check")=True then response.write "checked"%>>
                    审核留言</td>
                </tr>
              </table>
              </fieldset></td>
              </tr><br>
             <tr><td  colspan="9" width="100%"><table width="100%"  border="0" cellspacing="1" cellpadding="2" id="AdminPurviewDetail2" <%if rs("Purview")=3 then response.write "style='display:none'"%> >
              
              <legend>网站管理权限</legend>
              <!--<table width="100%" border="0" cellspacing="1" cellpadding="2"><tbody >-->
                
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Announce" <%if CheckPurview(rs("AdminPurview_Others"),"Announce")=True then response.write "checked"%>>
                    公告管理</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="FriendSite" <%if CheckPurview(rs("AdminPurview_Others"),"FriendSite")=True then response.write "checked"%>>
                    网站链接管理</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Vote" <%if CheckPurview(rs("AdminPurview_Others"),"Vote")=True then response.write "checked"%>>
                    调查管理&nbsp;</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Counter" <%if CheckPurview(rs("AdminPurview_Others"),"Counter")=True then response.write "checked"%>>
                    统计管理</td>
                </tr>
                <tr>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="User" <%if CheckPurview(rs("AdminPurview_Others"),"User")=True then response.write "checked"%>>
                    注册用户管理</td>
                  <td width="50%"><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="MailList" <%if CheckPurview(rs("AdminPurview_Others"),"MailList")=True then response.write "checked"%>>
                    邮件列表管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Skin" <%if CheckPurview(rs("AdminPurview_Others"),"Skin")=True then response.write "checked"%>>
                    配色模板管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Layout" <%if CheckPurview(rs("AdminPurview_Others"),"Layout")=True then response.write "checked"%>>
                    版面设计模板管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="JS" <%if CheckPurview(rs("AdminPurview_Others"),"JS")=True then response.write "checked"%>>
                    JS代码管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="Database" <%if CheckPurview(rs("AdminPurview_Others"),"Database")=True then response.write "checked"%>  disabled>
                    数据库管理</td>
                </tr>
                <tr>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="UpFile" <%if CheckPurview(rs("AdminPurview_Others"),"UpFile")=True then response.write "checked"%>>
                    上传文件管理</td>
                  <td><input name="AdminPurview_Others" type="checkbox" id="AdminPurview_Others" value="ModifyPwd" <%if CheckPurview(rs("AdminPurview_Others"),"ModifyPwd")=True then response.write "checked"%>>
                    修改自己的登录密码</td>
                </tr>
            <!-- </tbody> </table>-->
              
              </table></td></tr>
              </td>
          </tr>
        </table></td>
    </tr>
    <%	
'	Response.End() 
	 %>
    <tr>
      <td colspan="9" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveModifyPurview">
        <input  type="submit" name="Submit" value="保存修改结果" style="cursor:hand;">
        &nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
'	Response.Write("==" & Request("TrueName") & "==")
'	Response.Write( "==" & Request("StudentNumber") & "==" )
'	Response.Write("==" & Request("TeacherName") & "==")
'	Response.Write("==" &  Request("AdminPurview_Special") & "==")
%>
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
if AdminPurview=3 then
'学生管理员不可添加管理员
	%>
	<script>
//alert("You are running the "+navigator.appName+navigator.appVersion+" browser.")
alert("对不起，你没有权限添加管理员。")
//alert("你的浏览器是: "+navigator.appName+navigator.appVersion+".")
</script>

	<%
	Call main()

end if
'学生管理员不可添加管理员

if AdminPurview=1 or AdminPurview=2  then 
	dim username, password,PwdConfirm, purview
	dim AdminPurview_Article,AdminPurview_Soft,AdminPurview_Photo,AdminPurview_Guest,AdminPurview_Others
	dim ClassInputer_Article,ClassChecker_Article,ClassMaster_Article
	dim ClassInputer_Soft,ClassChecker_Soft,ClassMaster_Soft
	dim ClassInputer_Photo,ClassChecker_Photo,ClassMaster_Photo

	username=trim(Request("username"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	purview=trim(Request("purview"))
		if  Trim(Request("AdminPurview_Article"))="" then
			AdminPurview_Article=1
			else 
			AdminPurview_Article=Clng(trim(Request("AdminPurview_Article")))
		end if
		if	Trim(Request("AdminPurview_Soft"))="" then
			AdminPurview_Soft=4
			else
			AdminPurview_Soft=Clng(trim(Request("AdminPurview_Soft")))
		end if 
		if  trim(Request("AdminPurview_Photo"))="" then
		    AdminPurview_Photo=4
			else
			AdminPurview_Photo=Clng(trim(Request("AdminPurview_Photo")))
		end if 
	AdminPurview_Guest=replace(replace(trim(request("AdminPurview_Guest"))," ",""),"'","")
	AdminPurview_Others=replace(replace(trim(request("AdminPurview_Others"))," ",""),"'","")
	ClassInputer_Article=replace(replace(trim(request("ClassInputer_Article"))," ",""),"'","")
	ClassChecker_Article=replace(replace(trim(request("ClassChecker_Article"))," ",""),"'","")
	ClassMaster_Article=replace(replace(trim(request("ClassMaster_Article"))," ",""),"'","")
	ClassInputer_Soft=replace(replace(trim(request("ClassInputer_Soft"))," ",""),"'","")
	ClassChecker_Soft=replace(replace(trim(request("ClassChecker_Soft"))," ",""),"'","")
	ClassMaster_Soft=replace(replace(trim(request("ClassMaster_Soft"))," ",""),"'","")
	ClassInputer_Photo=replace(replace(trim(request("ClassInputer_Photo"))," ",""),"'","")
	ClassChecker_Photo=replace(replace(trim(request("ClassChecker_Photo"))," ",""),"'","")
	ClassMaster_Photo=replace(replace(trim(request("ClassMaster_Photo"))," ",""),"'","")
	if AdminPurview_Guest<>"" then
		AdminPurview_Guest=AdminPurview_Guest & "," & "Manage"
	end if
	if username="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>初始密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与初始密码相同！</li>"
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户权限不能为空！</li>"
	else
		purview=CInt(purview)
	end if

	if FoundErr=True then
		exit sub
	end if
	sql="Select * from Admin where username='"&username&"'"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>数据库中已经存在此管理员！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
   	rs.addnew
 	rs("username")=username
   	rs("password")=md5(password)
    rs("purview")=purview
	if purview=1 then
		rs("AdminPurview_Article")=3
		rs("AdminPurview_Soft")=3
		rs("AdminPurview_Photo")=3
		rs("AdminPurview_Guest")=""
		rs("AdminPurview_Others")=""
	else
		rs("AdminPurview_Article")=AdminPurview_Article
		rs("AdminPurview_Soft")=AdminPurview_Soft
		rs("AdminPurview_Photo")=AdminPurview_Photo
		rs("AdminPurview_Guest")=AdminPurview_Guest
		rs("AdminPurview_Others")=AdminPurview_Others
	end if
	
			'两课网站代码
		
		rs("TrueName")=Trim(Request("TrueName"))
		rs("StudentNumber")=Trim(Request("StudentNumber"))
		if  rs("Purview")=2 then
			rs("TeacherName")=Trim(Request("TrueName"))
			end if
		if rs("Purview")=1 then
			rs("TeacherName")=Trim(Request("TrueName"))
			end if
		if rs("Purview")=3 then
			rs("TeacherName")=Trim(Request("TeacherName"))
			end if 
		if purview=1 then
			rs("AdminPurview_Special")="admin"
		else 
			rs("AdminPurview_Special")=replace(replace(trim(request("AdminPurview_Special"))," ",""),"'","")
		end if 		
		rs("AdminPurview_SpecialID")=Trim(Request("AdminPurview_SpecialID")) & ","	

		'rs("TeacherName")=Trim(Request("TeacherName"))
		'结束两课网站代码

	
	
	
	
	rs.update
    rs.Close
	set rs=Nothing
	if AdminPurview_Article=3 then
		call AddClassMaster("ArticleClass","ClassInputer",UserName,ClassInputer_Article)
		call AddClassMaster("ArticleClass","ClassChecker",UserName,ClassChecker_Article)
		call AddClassMaster("ArticleClass","ClassMaster",UserName,ClassMaster_Article)
	end if
	if AdminPurview_Soft=3 then
		call AddClassMaster("SoftClass","ClassInputer",UserName,ClassInputer_Soft)
		call AddClassMaster("SoftClass","ClassChecker",UserName,ClassChecker_Soft)
		call AddClassMaster("SoftClass","ClassMaster",UserName,ClassMaster_Soft)
	end if
	if AdminPurview_Photo=3 then
		call AddClassMaster("PhotoClass","ClassInputer",UserName,ClassInputer_Photo)
		call AddClassMaster("PhotoClass","ClassChecker",UserName,ClassChecker_Photo)
		call AddClassMaster("PhotoClass","ClassMaster",UserName,ClassMaster_Photo)
	end if
	Call main()
end if
end sub

sub SaveModifyPwd()
	dim UserID, UserName,password,PwdConfirm
	UserID=trim(Request("ID"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
	else
		UserID=Clng(UserID)
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>新密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与新密码相同！</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此管理员！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs("password")=md5(password)
 	rs.update
	rs.Close
   	set rs=Nothing
    call main()
end sub

sub SaveModifyPurview()
if AdminPurview=3 then
'学生管理员不可添加管理员
	%>
	<script>
//alert("You are running the "+navigator.appName+navigator.appVersion+" browser.")
alert("对不起，你没有权限添加管理员。")
//alert("你的浏览器是: "+navigator.appName+navigator.appVersion+".")
</script>

	<%
	Call main()

end if
if AdminPurview=1 or AdminPurview=2 then
	dim UserID,UserName,Purview
	dim AdminPurview_Article,AdminPurview_Soft,AdminPurview_Photo,AdminPurview_Guest,AdminPurview_Others,AdminPurview_Special
	dim ClassInputer_Article,ClassChecker_Article,ClassMaster_Article
	dim ClassInputer_Soft,ClassChecker_Soft,ClassMaster_Soft
	dim ClassInputer_Photo,ClassChecker_Photo,ClassMaster_Photo
	dim OldAdminPurview_Article,OldAdminPurview_Soft,OldAdminPurview_Photo
	UserID=trim(Request("ID"))
	purview=trim(Request("purview"))
		'两课网站代码
		if  Trim(Request("AdminPurview_Article"))="" then
			AdminPurview_Article=1
			else 
			AdminPurview_Article=Clng(trim(Request("AdminPurview_Article")))
		end if
		if	Trim(Request("AdminPurview_Soft"))="" then
			AdminPurview_Soft=4
			else
			AdminPurview_Soft=Clng(trim(Request("AdminPurview_Soft")))
		end if 
		if  trim(Request("AdminPurview_Photo"))="" then
		    AdminPurview_Photo=4
			else
			AdminPurview_Photo=Clng(trim(Request("AdminPurview_Photo")))
		end if 
		'两课网站代码
	AdminPurview_Guest=replace(replace(trim(request("AdminPurview_Guest"))," ",""),"'","")
	AdminPurview_Others=replace(replace(trim(request("AdminPurview_Others"))," ",""),"'","")
	ClassInputer_Article=replace(replace(trim(request("ClassInputer_Article"))," ",""),"'","")
	ClassChecker_Article=replace(replace(trim(request("ClassChecker_Article"))," ",""),"'","")
	ClassMaster_Article=replace(replace(trim(request("ClassMaster_Article"))," ",""),"'","")
	ClassInputer_Soft=replace(replace(trim(request("ClassInputer_Soft"))," ",""),"'","")
	ClassChecker_Soft=replace(replace(trim(request("ClassChecker_Soft"))," ",""),"'","")
	ClassMaster_Soft=replace(replace(trim(request("ClassMaster_Soft"))," ",""),"'","")
	ClassInputer_Photo=replace(replace(trim(request("ClassInputer_Photo"))," ",""),"'","")
	ClassChecker_Photo=replace(replace(trim(request("ClassChecker_Photo"))," ",""),"'","")
	ClassMaster_Photo=replace(replace(trim(request("ClassMaster_Photo"))," ",""),"'","")
	if AdminPurview_Guest<>"" then
		AdminPurview_Guest=AdminPurview_Guest & "," & "Manage"
	end if
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
	else
		UserID=Clng(UserID)
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户权限不能为空！</li>"
	else
		purview=CInt(purview)
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此管理员！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	UserName=rs("UserName")
	OldAdminPurview_Article=rs("AdminPurview_Article")
	OldAdminPurview_Soft=rs("AdminPurview_Soft")
	OldAdminPurview_Photo=rs("AdminPurview_Photo")
    rs("purview")=purview
	if purview=1 then
		rs("AdminPurview_Article")=3
		rs("AdminPurview_Soft")=3
		rs("AdminPurview_Photo")=3
		rs("AdminPurview_Guest")=""
		rs("AdminPurview_Others")=""
	else
		rs("AdminPurview_Article")=AdminPurview_Article
		rs("AdminPurview_Soft")=AdminPurview_Soft
		rs("AdminPurview_Photo")=AdminPurview_Photo
		rs("AdminPurview_Guest")=AdminPurview_Guest
		rs("AdminPurview_Others")=AdminPurview_Others
		'两课网站代码
		
		rs("TrueName")=Trim(Request("TrueName"))
		rs("StudentNumber")=Trim(Request("StudentNumber"))
		if  rs("Purview")=2 then
			rs("TeacherName")=Trim(Request("TrueName"))
			end if
		if rs("Purview")=1 then
			rs("TeacherName")=Trim(Request("TrueName"))
			end if
		if rs("Purview")=3 then
			rs("TeacherName")=Trim(Request("TeacherName"))
			end if 
		if purview=1 then
			rs("AdminPurview_Special")="admin"
		else 
			rs("AdminPurview_Special")=replace(replace(trim(request("AdminPurview_Special"))," ",""),"'","")
		end if	
		rs("AdminPurview_SpecialID")=Trim(Request("AdminPurview_SpecialID")) & ","	

		'rs("TeacherName")=Trim(Request("TeacherName"))
		'结束两课网站代码
		
		
	end if
 	rs.update
	rs.Close
   	set rs=Nothing
	if OldAdminPurview_Article=3 then
		call RemoveClassMaster("ArticleClass",UserName)
	end if
	if OldAdminPurview_Soft=3 then
		call RemoveClassMaster("SoftClass",UserName)
	end if
	if OldAdminPurview_Photo=3 then
		call RemoveClassMaster("PhotoClass",UserName)
	end if
	if AdminPurview_Article=3 then
		call AddClassMaster("ArticleClass","ClassInputer",UserName,ClassInputer_Article)
		call AddClassMaster("ArticleClass","ClassChecker",UserName,ClassChecker_Article)
		call AddClassMaster("ArticleClass","ClassMaster",UserName,ClassMaster_Article)
	end if
	if AdminPurview_Soft=3 then
		call AddClassMaster("SoftClass","ClassInputer",UserName,ClassInputer_Soft)
		call AddClassMaster("SoftClass","ClassChecker",UserName,ClassChecker_Soft)
		call AddClassMaster("SoftClass","ClassMaster",UserName,ClassMaster_Soft)
	end if
	if AdminPurview_Photo=3 then
		call AddClassMaster("PhotoClass","ClassInputer",UserName,ClassInputer_Photo)
		call AddClassMaster("PhotoClass","ClassChecker",UserName,ClassChecker_Photo)
		call AddClassMaster("PhotoClass","ClassMaster",UserName,ClassMaster_Photo)
	end if
    call main()
	end if
end sub

sub DelAdmin()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的管理员ID</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Select * from Admin where ID in (" & UserID & ")"
	else
		UserID=clng(UserID)
		sql="select * from Admin where ID=" & UserID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	do while not rs.eof
		if rs("Purview")=2 then
			if rs("AdminPurview_Article")=3 then
				call RemoveClassMaster("ArticleClass",rs("UserName"))
			end if
			if rs("AdminPurview_Soft")=3 then
				call RemoveClassMaster("SoftClass",rs("UserName"))
			end if
			if rs("AdminPurview_Photo")=3 then
				call RemoveClassMaster("PhotoClass",rs("UserName"))
			end if
		end if
		rs.delete
		rs.update
		rs.movenext
	loop
	rs.close
	set rs=nothing
	call main()
end sub

sub AddClassMaster(SheetName,FieldName,strUserName,arrClassID)
	if isNull(arrClassID) or arrClassID="" then
		exit sub
	end if
	dim sqlMaster,rsMaster,ClassMaster
	sqlMaster="select ClassID," & FieldName & " from " & SheetName & " where ClassID in (" & arrClassID & ") order by RootID,OrderID"
	Set rsMaster=Server.CreateObject("Adodb.RecordSet")
	rsMaster.open sqlMaster,conn,1,3
	do while not rsMaster.eof
		if isNull(rsMaster(1)) or rsMaster(1)="" then
			rsMaster(1)=strUserName
		else
			rsMaster(1)=rsMaster(1) & "|" & strUserName
		end if
		rsMaster.update
		rsMaster.movenext
	loop
	rsMaster.close
	set rsMaster=nothing
end sub

sub RemoveClassMaster(SheetName,strUserName)
	dim sqlMaster,rsMaster,ClassMaster,arrClassMaster,i
	sqlMaster="select ClassInputer,ClassChecker,ClassMaster from " & SheetName
	Set rsMaster=Server.CreateObject("Adodb.RecordSet")
	rsMaster.open sqlMaster,conn,1,3
	do while not rsMaster.eof
		rsMaster(0)=RemoveStr(rsMaster(0),strUserName)
		rsMaster(1)=RemoveStr(rsMaster(1),strUserName)
		rsMaster(2)=RemoveStr(rsMaster(2),strUserName)
		rsMaster.update
		rsMaster.movenext
	loop
	rsMaster.close
	set rsMaster=nothing
end sub

function RemoveStr(str1,str2)
	if isnull(str1) or str1="" then
		RemoveStr=""
		exit function
	end if
	if str2="" then
		RemoveStr=str1
		exit function
	end if
	if instr(str1,"|")>0 then
		dim arrStr,tempStr,i
		arrStr=split(str1,"|")
		for i=0 to ubound(arrStr)
			if arrStr(i)<>str2 then
				if tempStr="" then
					tempStr=arrStr(i)
				else
					tempStr=tempStr & "|" & arrStr(i)
				end if
			end if
		next
		RemoveStr=tempStr
	else
		if str1=str2 then
			RemoveStr=""
		else
			RemoveStr=str1
		end if
	end if	
end function

sub Admin_AddStudentAdminInformation()
        dim sqlAdmin_Teacher,rsAdmin_Teacher
		if   AdminPurview=1   then
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,3,1
				   	Response.Write("<table id='StudentAdminPurviewDetail' style='display:none' ><tr><td width='50%' colspan='9' > <strong>学&nbsp;&nbsp;&nbsp;号：</strong> &nbsp; " )
				  	Response.Write("<input name=StudentNumber type=text MaxLength=16 value=''></td></tr>")
					Response.Write("<tr> <td width='50%' colspan=9 ><strong>任课教师：</strong>&nbsp; " )		  
					Response.Write("<select name='TeacherName' id='TeacherName' ")
					 if AdminPurview=1 then 
					 	Response.Write(" >")
					 end if
					 if AdminPurview=2 then 
					    Response.Write(" disabled >")
					 end if
					'rsAdmin_Teacher.movefirst
						do while not rsAdmin_Teacher.eof
							if  ( rsAdmin_Teacher("Purview")=2 ) then 
								if session("AdminTeacherName")=rsAdmin_Teacher("TrueName") then
									Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "' selected>"   &  rsAdmin_Teacher("TrueName") & "</option>" )

								else
									Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "'>"   &  rsAdmin_Teacher("TrueName") & "</option>" )
								end if
							end if
						rsAdmin_Teacher.movenext
						loop
					response.Write("</select></td></tr></table>")
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接
		end if
        if   AdminPurview=2   then
				'dim sqlAdmin_Teacher,rsAdmin_Teacher
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,3,1
				   	Response.Write("<table id='StudentAdminPurviewDetail' style='display:none' ><tr><td width='50%' colspan='9' > <strong>学&nbsp;&nbsp;&nbsp;号：</strong> &nbsp; " )
				  	Response.Write("<input name=StudentNumber type=text MaxLength=16 value=''></td></tr>")
					Response.Write("<tr> <td width='50%' colspan=9 ><strong>任课教师：</strong>&nbsp; " )		  
					Response.Write("<input type=text name='TeacherName' id='TeacherName' value='" &  session("AdminTeacherName")  &   "' readonly >")
					response.Write("</td></tr></table>")
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接
		end if
end sub
sub Admin_ModifyStudentAdminInformation()
        dim sqlAdmin_Teacher,rsAdmin_Teacher
		if  AdminPurview=1    then
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,1,1
				   	Response.Write("<table  id='StudentAdminPurviewDetail' ")
					 if ( rs("Purview")=1 or rs("Purview")=2 )  then 
					 	response.write "style='display:none'"  
					 end if
					Response.Write("  ><tr><td> <strong>学&nbsp;&nbsp;&nbsp;号：</strong>&nbsp;&nbsp;&nbsp;&nbsp;  " )
				  	Response.Write("<input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>")
					Response.Write("<tr> <td width='50%'><strong>任课教师：</strong>&nbsp;&nbsp;&nbsp; " )		  
					Response.Write("<select name='TeacherName' id='TeacherName' ")
					if AdminPurview=1 then 
					 	Response.Write(" >")
					 end if
					 if AdminPurview=2 then 
					    Response.Write(" disabled >")
					 end if
					rsAdmin_Teacher.movefirst
						do while not rsAdmin_Teacher.eof
							if  ( rsAdmin_Teacher("Purview")=2 ) then 
								if session("AdminTeacherName")=rsAdmin_Teacher("TrueName") then
									Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "' selected>"   &  rsAdmin_Teacher("TrueName") & "</option>" )
								else
									Response.Write("<option value='" & rsAdmin_Teacher("TrueName") & "'>"   &  rsAdmin_Teacher("TrueName") & "</option>" )
								end if
							end if
						rsAdmin_Teacher.movenext
						loop
					response.Write("</select></td></tr></table>")
				'关闭数据库连接
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接
		end if
if  AdminPurview=2    then
				sqlAdmin_Teacher = "select TrueName,Purview from Admin where Purview=2"
				set rsAdmin_Teacher=server.CreateObject("adodb.recordset")
				rsAdmin_Teacher.open sqlAdmin_Teacher,conn,1,1
				   	Response.Write("<table  id='StudentAdminPurviewDetail' ")
					 if ( rs("Purview")=1 or rs("Purview")=2 )  then 
					 	response.write "style='display:none'"  
					 end if
					Response.Write("  ><tr><td> <strong>学&nbsp;&nbsp;&nbsp;号：</strong>&nbsp;&nbsp;&nbsp;&nbsp;  " )
				  	Response.Write("<input name=StudentNumber type=text value=" & rs("StudentNumber")  & "></td></tr>")
					Response.Write("<tr> <td width='50%'><strong>任课教师：</strong>&nbsp;&nbsp;&nbsp; " )		  
					Response.Write("<input type=text name='TeacherName' id='TeacherName' value='" &  session("AdminTeacherName")  &   "' readonly >")
					response.Write("</td></tr></table>")
				'关闭数据库连接
				rsAdmin_Teacher.close
   				 set rsAdmin_Teacher = nothing
				'关闭数据库连接
		end if
end sub

%>
