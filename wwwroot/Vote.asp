<!--#include file="inc/conn.asp" -->
<!--#include file="inc/function.asp" -->
<%
dim Action,ID,VoteType,VoteOption,arrOptions,sqlVote,rsVote
dim Voted,VotedID,arrVotedID,i
dim FoundErr,ErrMsg
Action=trim(Request("Action"))
ID=Trim(request("ID"))
VoteType=Trim(request("VoteType"))
VoteOption=trim(request("VoteOption"))
Voted=False
VotedID=session("VotedID")

if Id="" then
	founderr=true
	errmsg=errmsg+"<br><li>不能确定调查ID</li>"
	call WriteErrMsg()
	response.end
else
	ID=CLng(ID)
	if instr(VotedID,",")>0 then
		arrVotedID=split(VotedID,",")
		for i=0 to ubound(arrVotedID)
			if Clng(arrVotedID(i))=ID then
				Voted=True
				exit for
			end if
		next
	else
		if VotedID=ID then
			Voted=True
		end if
	end if
end if
if Action="" or VoteOption="" then
	Action="Show"
end if
If Action = "Vote" And VoteOption<>"" and Voted=False Then
	if VoteType="Single" then
		conn.execute "Update Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & ID
	else
		if instr(VoteOption,",")>0 then
			arrOptions=split(VoteOption,",")
			for i=0 to ubound(arrOptions)
				conn.execute "Update Vote set answer" & cint(trim(arrOptions(i)))  & "= answer" & cint(trim(arrOptions(i))) & "+1 where ID=" & Clng(ID)
			next
		else
			conn.execute "Update Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & Clng(ID)
		end if 
	end if
	if VotedID="" then
		session("VotedID")=ID
	else
		session("VotedID")=VotedID & "," & ID
	end if
End If
sqlVote="Select * from Vote Where ID=" & ID
Set rsVote = Server.CreateObject("ADODB.Recordset")
rsVote.open sqlVote,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>调查结果</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgcolor="#eeeeee">
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td valign="top" align="center"><strong><%
if Action="Vote" And VoteOption<>"" then
	response.write "<font color='#FF0000' size='3'>"
	if Session("UserName")<>"" then response.write Session("UserName") & "，"
	if Voted=True then
		response.write "==　你已经投过票了，请勿重复投票！　=="
    else	
		response.write "==　非常感谢您的投票！　=="
	end if
	response.write "</font><br>"
end if
		%></strong><br> 
      <table width="700" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF" class="border">
        <tr align="center" class="title"> 
          <td width="702" height="22" colspan="3"><strong><img src="Images/p1.GIF" width="16" height="16" align="absmiddle"> 
            网站关于<font color="#FF0000">“</font><font color="#FF0000"><%=rsVote("Title")%>”</font>的调查结果</strong></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <table width="700" border="0" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#000000" bordercolordark="#CCCCCC">
              <tr> 
                <td> <strong>　&middot;目前网友的总投票数为：</strong><font color="#FF0000"> 
                  <%
  dim totalVote
  totalVote=0
  for i=1 to 8
  	if rsVote("Select" & i)="" then exit for
	totalVote=totalVote+rsVote("answer"& i)
  next
  response.Write(totalVote & "票")
  if totalVote=0 then totalVote=1
  %>
                  </font> </td>
              </tr>
              <%
  for i=1 to 8
  	if trim(rsVote("Select" & i) & "")="" then exit for
  %>
              <% next %>
            </table>
            <table width="700" border="0" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#000000" bordercolordark="#CCCCCC">
              <%
  for i=1 to 8
  	if trim(rsVote("Select" & i) & "")="" then exit for
  %>
              <tr> 
                <td align="right"> <div align="left"><font color="#cc0000"></font> 
                    <table width="539" border="0" cellpadding="0" cellspacing="0" background="Images/dc/BG2.GIF">
                      <tr> 
                        <td height="23"><font color="#cc0000">&nbsp;</font>&nbsp;选项<%=i%>：<strong><%=rsVote("Select"& i)%></strong></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td><table border=0 cellspacing="0" cellpadding="0" height="30">
                    <tr> 
                      <td height="11" valign="top">&nbsp;&nbsp; 得票率：<img src="Images/dc/left.gif" width="4" height="21" border="0" align="top"><img src="Images/dc/greenbar.gif" width="1" height="21" align="top"><%dim perVote
	perVote=round(rsVote("answer"& i)/totalVote,4)
	response.write "<img src='Images/dc/greenbar.gif' width='" & int(360*pervote) & "' height='21' align='absmiddle'>"
	perVote=perVote*100%></td><td valign="top"><img src="Images/dc/mid.gif" width="6" height="21" align="top"><%dim perVote2
	perVote2=round(rsVote("answer"& i)/totalVote,4)
	response.write "<img src='Images/dc/whitebar.gif' width='" & 325-int(360*pervote2) & "' height='21' align='absmiddle'>"
	perVote2=perVote2*100%><img src="Images/dc/right.gif" width="6" height="21" border="0" align="top"></td></tr>
                    <tr><td></td>
                      <td height="19">占：<%	if perVote<1 and perVote<>0 then
		response.write "&nbsp;0" & perVote & "%"
	else
		response.write "&nbsp;" & perVote & "%"
	end if
%>
                        [得：<font color="#ff0000"><%response.write rsVote("answer"& i)%></font>票]</td></tr></table>
                </td>
              </tr>
              <% next %>
            </table></td>
        </tr>
      </table>
      
    </td>
  </tr>
  <tr> 
    <td valign="top"> 
      <table width="100%" border="0" cellpadding="1" cellspacing="4" bgcolor="#FFFFFF" class="border">
        <tr> 
          <td width="10%">&nbsp; </td>
          <td width="90%">
            <%
if Action="Show" and Voted=False then 
		if Session("UserName")<>"" then
			response.write Session("UserName") & "，"
		end if 
	    response.Write "<br><strong>您还没有投票，请您在此投下您宝贵的一票！</strong>"
		response.write "<form name='VoteForm' method='post' action='vote.asp'>"
		response.write "&nbsp;" & rsVote("Title") & "<br>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action' type='hidden' value='Vote'>"
		response.write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
		response.write "</form>"
end if

dim sqlOtherVote,rsOtherVote
if session("VoteID")<>"" then
	sqlOtherVote="Select * from Vote Where ID Not In (" & session("VotedID") & ") order by ID desc"
else
	sqlOtherVote="select * from Vote where ID<>" & ID
end if
Set rsOtherVote = Server.CreateObject("ADODB.Recordset")
rsOtherVote.open sqlOtherVote,conn,1,1
if rsOtherVote.bof and rsOtherVote.eof then
	response.write "<br>感谢您参加了本站的所有调查！！！"
else
	response.write "<br>欢迎你继续参加本站的其他调查：<br><br>"
	do while not rsOtherVote.eof
		response.write "<li><a href='Vote.asp?ID=" & rsOtherVote("ID") & "'>" & rsOtherVote("Title") & "</a></li>"
		rsOtherVote.movenext
	loop
end if
rsOtherVote.close
set rsOtherVote=nothing
%>
          </td>
        </tr>
      </table> </td>
  </tr>
  <tr>
    <td valign="top"> 
      <div align="center">【<a href="javascript:window.close();">关闭窗口</a>】</div></td>
  </tr>
</table>
</BODY>
</HTML>
<%
rsVote.Close()
Set rsVote = Nothing
call CloseConn()
%>
