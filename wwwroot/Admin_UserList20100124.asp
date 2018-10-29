<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
dim OrderType
MaxPerPage=20
SkinID=0
OrderType=trim(request("OrderType"))
if OrderType="" then
	OrderType=1
else
	OrderType=Clng(OrderType)
end if
'按教师排序
dim OrderByOnlyATeacher,OrderByTeacherName
OrderByOnlyATeacher=Trim(Request("OrderByOnlyATeacher"))
OrderByTeacherName=Trim(Request("OrderByTeacherName"))
if OrderByOnlyATeacher="" then 
		OrderByOnlyATeacher=1
	else
		OrderByOnlyATeacher=Clng(OrderByOnlyATeacher)
end if
if OrderByTeacherName="" then 
		OrderByTeacherName=1
	else
		OrderByTeacherName=OrderByTeacherName
end if
'结束按教师排序
PageTitle="用户列表"
strFileName="Admin_UserList.asp?OrderType=" & OrderType
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
      

    <td width="5" bgcolor="#949693"></td>
    <td  valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#949693" class="border">
        <tr> 
          <td valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"> <% call Admin_ShowAllUser() %> </td>
              </tr>
            </table></td>
        </tr>
        <tr class="tdbg"> 
          <td> <table width="100%" border="0" cellspacing="5" cellpadding="0">
              <tr class="tdbg_leftall"> 
                <td> <%
		  if TotalPut>0 then
		  	call showpage(strFileName,totalPut,MaxPerPage,true,true,"个用户")
		  end if
		  %> </td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="5"  valign="top" bgcolor="#949693">&nbsp;</td>
  </tr>
</table>


<% call PopAnnouceWindow(400,300) %>
</body>
</html>
<%
set rsArticle=nothing
set rsPic=nothing
call CloseConn()
%>