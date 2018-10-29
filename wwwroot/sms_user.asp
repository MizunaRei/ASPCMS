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
dim tablebody,strErr
dim boxName,smscount,smstype,readaction,turl
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
	call smsmain()
end if
call bottom()
%>

</body>
</html>

<%
sub smsmain()
smscount=1
select case request("action")
case "inbox"
	boxName="收件箱"
	smstype="inbox"
	readaction="read"
	turl="readsms"
	sql="select * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	call smsbox()
case "outbox"
	boxName="草稿箱"
	smstype="outbox"
	readaction="edit"
	turl="sms"
	sql="select * from message where sender='"&trim(membername)&"' and issend=0 and delS=0 order by sendtime desc"
	call smsbox()
case "issend"
	boxName="已发送的消息"
	smstype="issend"
	readaction="outread"
	turl="readsms"
	sql="select * from message where sender='"&trim(membername)&"' and issend=1 and delS=0 order by sendtime desc"
	call smsbox()
case "recycle"
	boxName="垃圾箱"
	smstype="recycle"
	readaction="read"
	turl="readsms"
	sql="select * from message where ((sender='"&trim(membername)&"' and delS=1) or (incept='"&trim(membername)&"' and delR=1)) and not delS=2 order by sendtime desc"
	call smsbox()
case else
	boxName="收件箱"
	smstype="inbox"
	readaction="read"
	turl="readsms"
	sql="select * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	call smsbox()
end select
end sub

sub smsbox()
dim newstyle
call mainnav()
%>
<br><table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td valign="top">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="sms_border">
<form action="sms_main.asp" method=post name=inbox>
<tr height='22'>
<td class=txt_css valign=middle width=30>已读</td>
<td class=txt_css valign=middle width=100>
<%if smstype="inbox" or smstype="recycle" then response.write "发件人" else response.write "收件人"%>
</td>
<td class=txt_css valign=middle width=300>主题</td>
<td class=txt_css valign=middle width=150>日期</td>
<td class=txt_css valign=middle width=50>大小</td>
<td class=txt_css valign=middle width=30>操作</td>
</tr>
<%
	set rs=server.createobject("adodb.recordset")
	rs.open sql,Conn_User,1,1
	if rs.eof and rs.bof then
%>
<tr>
<td class=tdbg align=center valign=middle colspan=6>您的<%=boxname%>中没有任何内容。</td>
</tr>
<%else%>
<%do while not rs.eof%>
<%
if rs("flag")=0 then
tablebody="tablebody2"
newstyle="font-weight:bold"
else
tablebody="tablebody1"
newstyle="font-weight:normal"
end if
%>
<tr>
<td class=tdbg align=center valign=middle>
<%
select case smstype
case "inbox"
	if rs("flag")=0 then
	response.write "<img src=""images/m_news.gif"">"
	else
	response.write "<img src=""images/m_olds.gif"">"
	end if
case "outbox"
	response.write "<img src=""images/m_issend_2.gif"">"
case "issend"
	response.write "<img src=""images/m_issend_1.gif"">"
case "recycle"
	if rs("flag")=0 then
	response.write "<img src=""images/m_news.gif"">"
	else
	response.write "<img src=""images/m_olds.gif"">"
	end if
end select
%>
</td>
<td class=tdbg align=center valign=middle style="<%=newstyle%>">
<%if smstype="inbox" or smstype="recycle" then%>
<%=htmlencode(rs("sender"))%>
<%else%>
<%=htmlencode(rs("incept"))%>
<%end if%>
</td>
<td class=tdbg align=left style="<%=newstyle%>"><a href="JavaScript:openScript2('sms_main.asp?action=<%=readaction%>&id=<%=rs("id")%>&sender=<%=rs("sender")%>',500,400)"><%=htmlencode(rs("title"))%></a>	</td>
<td class=tdbg style="<%=newstyle%>"><%=rs("sendtime")%></td>
<td class=tdbg style="<%=newstyle%>"><%=len(rs("content"))%>Byte</td>
<td align=center valign=middle width=30 class=tdbg><input type=checkbox name=id value=<%=rs("id")%>></td>
</tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
<tr> 
<td align=right valign=middle colspan=6 class=tdbg>节省每一分空间，请及时删除无用信息&nbsp;<input type=checkbox name=chkall value=on onClick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=submit name=action onClick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除<%=replace(boxname,"箱","")%>">&nbsp;<input type=submit name=action onClick="{if(confirm('确定清除<%=boxname%>所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空<%=boxname%>"></td>
</tr>
</form>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
  <td  height="13" align="center" valign="top"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr> 
		<td height="13" Class="tdbg_left2"></td>
	  </tr>
	</table></td>
</tr>
</table>
</td>
</tr>
</table>
<%
end sub
%>