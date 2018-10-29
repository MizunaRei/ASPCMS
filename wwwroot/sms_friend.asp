<!--#include file="Inc/syscode_Article.asp"-->
<%
const ChannelID=2
Const ShowRunTime="Yes"
SkinID=0
PageTitle="用户短信服务"
dim membername
membername=Trim(Request.Cookies("asp163")("UserName"))
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<style type="text/css">
.sms_border
{
background:#6687BA;
}
</style>
<%call MenuJS()%>

</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<%

if CheckUserLogined()=False then
  	errmsg=errmsg+"<br><br>"+"<li>您还没有登录。<li>您没有使用此项操作的权限。"
	founderr=true
end if

if founderr then
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table><br>" & vbcrlf
	response.write strErr
else
	call mainnav()
	select case request("action")
	case "info"
		call info()
	case "addF"
		call addF()
	case "saveF"
		call saveF()
	case "删除"
		call DelFriend()
	case "清空好友"
		call AllDelFriend()
	case else
		call info()
	end select
	if founderr then call WriteErrMsg()
end if
if not founderr then 
call bottom()
end if
sub info()
%>
<br><table width="770" cellpadding=3 cellspacing=1 align=center class=sms_border>
<form action="sms_friend.asp" method=post name=inbox>
            <tr>
                <td class=txt_css valign=middle width="25%">姓名</td>
                <td class=txt_css valign=middle width="25%">邮件</td>
                <td class=txt_css valign=middle width="25%">主页</td>
                <td class=txt_css valign=middle width="10%">QQ</td>
                <td class=txt_css valign=middle width="10%">发短信</td>
                <td class=txt_css valign=middle width="5%">操作</td>
            </tr>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select F.*,U."&db_User_Email&",U."&db_User_Homepage&",U."&db_User_QQ&" from Friend F inner join "&db_User_Table&" U on F.F_Friend=U."&db_User_Name&" where F.F_username='"&trim(membername)&"' order by F.f_addtime desc"
	rs.open sql,Conn_User,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tdbg align=center valign=middle colspan=6>您的好友列表中没有任何内容。</td>
                </tr>
		
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tdbg><a href="dispuser.asp?name=<%=htmlencode(rs("F_friend"))%>" target=_blank><%=htmlencode(rs("F_friend"))%></a></td>
                    <td align=center valign=middle class=tdbg><a href="mailto:<%=htmlencode(rs("UserEmail"))%>"><%=htmlencode(rs("UserEmail"))%></a></td>
                    <td align=center class=tdbg><a href="<%=htmlencode(rs("homepage"))%>" target=_blank><%=htmlencode(rs("homepage"))%></a></td>
                    <td align=center class=tdbg><%=htmlencode(rs("Oicq"))%></td>
                    <td align=center class=tdbg><a href="JavaScript:openScript2('sms_main.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>',500,400)">发送</a></td>
                <td align=center class=tdbg><input type=checkbox name=id value=<%=rs("f_id")%>></td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
                
        <tr> 
          <td align=right valign=middle colspan=6 class=tdbg><input type=checkbox name=chkall value=on onClick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=button name=action onClick="location.href='sms_friend.asp?action=addF'" value="添加好友">&nbsp;<input type=submit name=action onClick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除">&nbsp;<input type=submit name=action onClick="{if(confirm('确定清除所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空好友"></td>
		</tr>
		</form>
		</table><BR>
<%
end sub

sub delFriend()
dim delid
delid=replace(request.form("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"请选择相关参数。"
founderr=true
exit sub
else
	Conn_User.execute("delete from Friend where F_username='"&trim(membername)&"' and F_id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除选定的好友记录。"
	call sms_suc()
end if
end sub
sub AllDelFriend()
	Conn_User.execute("delete from Friend where F_username='"&trim(membername)&"'")
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除了所有好友列表。"
	call sms_suc()
end sub

sub addF()
%>
<br>
<%call userlist()%>
<table width="760" cellpadding=3 cellspacing=1 align=center class=sms_border>
<form action="sms_friend.asp" method=post name=messager>
          <tr> 
            <td class=txt_css colspan=2 align=center> 
              <input type=hidden name="action" value="saveF">
              加入好友--请完整输入下列信息</td>
          </tr>
          <tr height=50> 
            <td class=tdbg valign=middle width=70><b>好友：</b></td>
            <td class=tdbg valign=middle>
              <input type=text name="touser" size=50 value="<%=request("myFriend")%>">
			  &nbsp;使用逗号（,）分开，最多5位用户
            </td>
          </tr>
          <tr> 
            <td valign=middle colspan=2 align=center class=tdbg> 
              <input type=Submit value="保存" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="清除">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub saveF()
dim incept,i
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>您忘记填写发送对象了吧。"
	founderr=true
	exit sub
else
	incept=checkStr(request("touser"))
	incept=split(incept,",")
end if

for i=0 to ubound(incept)
set rs=server.createobject("adodb.recordset")
sql="select "&db_User_Name&" from "&db_User_Table&" where "&db_User_Name&"='"&incept(i)&"'"
set rs=Conn_User.execute(sql)
if rs.eof and rs.bof then
	set rs=nothing
	sql="select username from admin where username='"&replace(incept(i),"'","")&"'"
        set rs=Conn.execute(sql)
        if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>系统没有这个用户，操作未成功。"
	founderr=true
	exit sub
	end if
end if
set rs=nothing

if membername=trim(incept(i)) then
	errmsg=errmsg+"<br>"+"<li>不能把自已添加为好友。"
	founderr=true
	exit sub
end if

sql="select F_friend from friend where F_username='"&membername&"' and  F_friend='"&incept(i)&"'"
set rs=Conn_User.execute(sql)
if rs.eof and rs.bof then
	sql="insert into friend (F_username,F_friend,F_addtime) values ('"&membername&"','"&trim(incept(i))&"',Now())"
	Conn_User.execute(sql)
end if
if i>4 then
	errmsg=errmsg+"<br>"+"<li>每次最多只能添加5位用户，您的名单5位以后的请重新填写。"
	founderr=true
	exit sub
	exit for
end if
next

sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，好友添加成功。"
call sms_suc()
end sub

sub userlist()
response.write "<table class=border align=center cellpadding=2 cellspacing=1 border=0 width=""760"" style=""word-break:break-all;""><tr>网站管理员组<br>"
dim admin_face
sql="select username from admin order by ID"
set rs=Conn.execute(sql)
i=0
do while not rs.eof
admin_face="<img src=""images/admin_face.gif"" width=24 height=30>"
if membername=rs(0) then
	response.write "<td width=""14%"">" & admin_face&"&nbsp;<a href=sms_friend.asp?action=saveF&touser="&rs(0)&" title=""网站管理员""><font color=""#0000ff"">"&rs(0)&"</font></a></td>"
else
	response.write "<td width=""14%"">" & admin_face&"&nbsp;<a href=sms_friend.asp?action=saveF&touser="&rs(0)&" title=""网站管理员"">"&rs(0)&"</a></td>"
end if

if i=6 then response.write "</tr><tr>"
if i>6 then 
	i=1
else
	i=i+1
end if
rs.movenext
loop
response.write "</tr></TABLE><br>"
set rs=nothing

response.write "<table class=border align=center cellpadding=2 cellspacing=1 border=0 width=""760"" style=""word-break:break-all;""><tr>普通用户组<br>"
dim user_face,user_info,sex,i
sql="select "&db_User_Name&","&db_User_Sex&","&db_User_QQ&","&db_User_Email&" from "&db_User_Table&" order by "&db_User_ID&""
set rs=Conn_User.execute(sql)
i=0
do while not rs.eof
if rs(1)=1 then
	sex="男"
else
	sex="女"
end if
user_info="性别："& sex & vbcrlf & "QQ："& rs(2) & vbcrlf &"Email："& rs(3)
user_face="<img src=""images/user_face.gif"" width=12 height=11>"
if membername=rs(0) then
	response.write "<td width=""14%"">" & user_face&"&nbsp;<a href=sms_friend.asp?action=saveF&touser="&rs(0)&" title="""& user_info &"""><font color=""#0000ff"">"&rs(0)&"</font></a></td>"
else
	response.write "<td width=""14%"">" & user_face&"&nbsp;<a href=sms_friend.asp?action=saveF&touser="&rs(0)&" title="""& user_info &""">"&rs(0)&"</a></td>"
end if

if i=6 then response.write "</tr><tr>"
if i>6 then 
	i=1
else
	i=i+1
end if
rs.movenext
loop
response.write "</tr></TABLE><br>"
set rs=nothing
end sub
%>