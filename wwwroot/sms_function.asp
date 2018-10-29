<script language=javascript>
	function openScript(url)
	{
		var Win = window.open(url,"UserControlPad");
	}
	function openScript2(url, width, height)
	{
		var Win = window.open(url,"UserControlPad",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
	}
	function CheckAll(form)
	{
	  for (var i=0;i<form.elements.length;i++)    {
		var e = form.elements[i];
		if (e.name != 'chkall')       e.checked = form.chkall.checked; 
	   }
	}
</script>
<%
dim sucmsg
sub sms_suc()
	dim strErr
	strErr=strErr & "<br><table cellpadding=2 cellspacing=1 border=0 width=460 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>用户短信操作成功信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>操作成功：</b><br>" & sucmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table><br>" & vbcrlf
	response.write strErr
end sub


function newincept()
rs=Conn_User.execute("Select Count(id) From Message Where flag=0 and issend=1 and delR=0 And incept='"& Trim(Request.Cookies("asp163")("UserName")) &"'")
	newincept=rs(0)
	set rs=nothing
	if isnull(newincept) then newincept=0
end function

function inceptid(stype)
	set rs=Conn_User.execute("Select top 1 id,sender From Message Where flag=0 and issend=1 and delR=0 And incept='"& Trim(Request.Cookies("asp163")("UserName")) &"'")
	if stype=1 then
	inceptid=rs(0)
	else
	inceptid=rs(1)
	end if
	set rs=nothing
end function

function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    HTMLEncode = fString
end if
end function

function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

sub footer()
	dim strTemp
	strTemp="<table align='center' border='0' cellpadding='0' cellspacing='0' ><tr align='center' height='20' valign='bottom'><td>"
	strTemp= strTemp & Copyright
	strTemp= strTemp & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;站长：<a href='mailto:" & WebmasterEmail & "'>" & WebmasterName & "</a>"
	strTemp= strTemp & "</td></tr></table>"
	response.write strTemp
end sub

sub mainnav()
response.write "<TABLE width='770' cellpadding=6 cellspacing=1 align=center class=border><TBODY>"&_
				"<TR>"&_
				"<TD align=center class=tdbg><a href=sms_user.asp?action=inbox><img src=images/m_inbox.gif border=0 alt=收件箱></a> &nbsp; <a href=sms_user.asp?action=outbox><img src=images/M_outbox.gif border=0 alt=发件箱></a> &nbsp; <a href=sms_user.asp?action=issend><img src=images/M_issend.gif border=0 alt=已发送邮件></a>&nbsp; <a href=sms_user.asp?action=recycle><img src=images/M_recycle.gif border=0 alt=废件箱></a>&nbsp; <a href=sms_friend.asp><img src=images/M_address.gif border=0 alt=地址簿></a>&nbsp;<a href=JavaScript:openScript2('sms_main.asp?action=new',500,400)><img src=images/m_write.gif border=0 alt=发送消息></a>"&_
                           "</td></tr></TBODY></TABLE>"
end sub
%>