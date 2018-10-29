<!--#include file="Inc/syscode_diary.asp"-->
<%
'请勿改动下面这三行代码
MaxPerPage=100
SkinID=0

dim UserGroupLevel,Users,UserGroupName,keywords
set Rs=server.createobject("adodb.recordset")

UserGroupLevel=trim(request("UserGroupLevel"))
if UserGroupLevel="" then
	UserGroupLevel=9
else
	UserGroupLevel=Clng(UserGroupLevel)
end if
keywords=trim(request("keywords"))
select case UserGroupLevel
	case 9 UserGroupName="VIP用户"
	case 99 UserGroupName="付费用户"
	case 999 UserGroupName="注册用户"
	case else UserGroupName="全部用户"
end select
PageTitle=UserGroupName&"列表"
strFileName="diary_selectUsers.asp?UserGroupLevel="&UserGroupLevel&"&keywords="&keywords
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
BODY {BACKGROUND-ATTACHMENT: fixed; BACKGROUND-IMAGE: diary_images/girl.gif; BACKGROUND-POSITION:center center;  BACKGROUND-REPEAT: no-repeat;scrollbar-track-color:#ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: 9pt; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #eeeeee;  SCROLLBAR-3DLIGHT-COLOR: #eeeeee; SCROLLBAR-ARROW-COLOR: #dddddd; FONT-FAMILY: "Verdana"; SCROLLBAR-DARKSHADOW-COLOR: #ffffff
}
select{font-size:8pt;font-family:verdana;background-color:#ffffff;border:1px dotted #cccccc; color:#333333;}
input{font-size:8pt;font-family:verdana;background-color:#ffffff;border-bottom:1px solid #cccccc;border-left:1px solid #cccccc; border-top:0px solid #cccccc;border-right:0px dotted #cccccc;color:#333333;}
textarea{font-size:8pt; font-family:verdana;background-color:#ffffff;border:1px dotted #cccccc;color:#333333;letter-spacing : 1pt ;line-height : 150%}
A {
	COLOR: #333333; TEXT-DECORATION: none ;border-bottom:1px dotted
}
A:hover {
	COLOR: #333333; background-color:#C0FFFF;
}
td {FONT-SIZE: 9pt;  FONT-FAMILY: "Verdana"; color:#3333333;letter-spacing : 1pt ;line-height : 150%}
</style>

</head>
<body>
 <table width="680" border="0" align="center" cellpadding="4" cellspacing="0">
  <form method="post" name="myform" ID="myform" action="">
  <tr>
	<td valign="top"><b>选定用户预览：</b><br>
		<textarea name="UserList" cols="98" rows="3" readonly="readonly"></textarea>
		</td>
  </tr>
  <tr align="center">
    <td>[<a href='diary_selectUsers.asp?UserGroupLevel=0'>全部用户</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='diary_selectUsers.asp?UserGroupLevel=999'>注册用户</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='diary_selectUsers.asp?UserGroupLevel=99'>会费用户</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='diary_selectUsers.asp?UserGroupLevel=9'>VIP用户</a>]
	&nbsp;&nbsp;&nbsp;&nbsp;<input name="keywords" type="text" size="12">&nbsp;&nbsp;<input type="submit" value="查找用户"></td>
  </tr></form>
</table>
<table width="700" border="0" align="center" cellpadding="2" cellspacing="0" class="title_main">
  <tr>
    <td width="32">&nbsp;</td>
    <td class="title_maintxt"><b><font color=red><%=UserGroupName%></font>列表:</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(点击用户即可把用户加入列表中)</td>
  </tr>
</table>
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr>
    <td valign="top"><table width="99%" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
        <tr>
          <td><% call ShowAllusers() %></td>
        </tr>
        <tr>
          <td align=center> <a href="" onclick="return adduser('<%=Users%>')">增加以上所有用户</a>&nbsp;&nbsp;<a href="" onclick="return deluser(1)">删除最后一个用户</a>&nbsp;&nbsp;<a href="" onclick="return deluser(0)">删除全部用户</a></td>
        </tr>
        <tr class="tdbg">
          <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tdbg_leftall">
                <td> <%
			  if TotalPut>0 then
				call showpage(strFileName,totalPut,MaxPerPage,true,true,"个用户")
			  end if
			  %> </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="1" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td  height="15" align="right" valign="top">&nbsp;[<a href="javascript:window.close();">关闭窗口</a>]</td>
        </tr>
      </table></td>
  </tr>
</table>

<script language="javascript">
myform.UserList.value=opener.myform.readers.value;
var oldusername="";
function adduser(obj)
{
	if(opener.myform.readers.value=="")
	{
		opener.myform.readers.value=obj;
		myform.UserList.value=opener.myform.readers.value;
		return false;
	}

	var username=obj.substring(1,obj.length);
	var singleuser=username.split("|");
	for(i=0;i<singleuser.length;i++)
	{
		if(checkuser(opener.myform.readers.value,singleuser[i]))
		{
			alert(singleuser[i]+" 用户已经存在，此操作已经忽略！");
		}
		else
		{
			opener.myform.readers.value=opener.myform.readers.value+"|"+singleuser[i];
			myform.UserList.value=opener.myform.readers.value;
		}
	}
	return false;
}

function deluser(num)
{
	if (num==0 || opener.myform.readers.value=="" || opener.myform.readers.value=="|")
	{
		opener.myform.readers.value="";
		myform.UserList.value="";
		return false;
	}

	var strDel=opener.myform.readers.value;
	strDel=strDel.substring(1,strDel.length);
	var s=strDel.split("|");
	opener.myform.readers.value="|"+strDel.substring(0,strDel.length-s[s.length-1].length-1);
	myform.UserList.value=opener.myform.readers.value;
	return false;
}

function checkuser(userlist,thisuser)
{
	var mystr=userlist.substring(1,userlist.length);
	var s=mystr.split("|");
	for (j=0;j<s.length;j++)
		{
		if(s[j]==thisuser)
			return true;
		}
	return false;
}
</script>

</body>
</html>
<%
call CloseConn()


'**************************************************
'过程名：ShowAllusers
'作  用：显示全部符合要求的用户名
'**************************************************
Function ShowAllusers()
	if UserGroupLevel=9 or UserGroupLevel=99 or UserGroupLevel=999 then
		sqlUser="select * from " & db_User_Table & " where UserLevel= "&UserGroupLevel&" and Username like '%"&keywords&"%' order by UserName"
	else
		sqlUser="select * from " & db_User_Table &" where Username like '%"&keywords&"%' order by UserName"
	end if
	Rs.open sqlUser,Conn_User,1,1
	if Rs.bof and Rs.eof then
		totalput=0
		response.write "<br><li>没有任何用户</li>"
	else
		totalput=Rs.recordcount
		if request("Page")="" then
			currentPage=1
		else
			currentPage=cint(request("Page"))
		end if
		if currentPage=1 then
			call ShowUserList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	Rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=Rs.bookmark
            	call ShowUserList()
        	else
	        	currentPage=1
           		call ShowUserList()
	    	end if
		end if
	end if
	Rs.close
	set Rs=nothing
end Function


'**************************************************
'过程名：ShowUserList
'作  用：依次输出每个用户名
'**************************************************
Function ShowUserList()
	dim i,strTemp

	strTemp = strTemp & "<table width='700' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>"
	do while not Rs.eof
		Users = Users&"|"&Rs(db_User_Name)
		if (i mod 10)=0 then strTemp = strTemp & "<tr>"
		strTemp = strTemp & "<td><a href='' onclick='return adduser(""|" &Rs(db_User_Name)& """)'>" & Rs(db_User_Name) & "</a></td><td align='center'>"
		i=i+1
		if i>=MaxPerPage then exit do
		if (i mod 10)=0 then strTemp = strTemp & "</tr>"
		Rs.movenext
	loop
	strTemp = strTemp & "</table>"
	response.write(strTemp)
end Function

%>