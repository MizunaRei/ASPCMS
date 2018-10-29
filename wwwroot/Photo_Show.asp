<!--#include file="Inc/syscode_Photo.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=4
Const ShowRunTime="Yes"
MaxPerPage=20
SkinID=0
PageTitle="图片信息"
%>
<html>
<head>
<title><%=PhotoTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="Top.asp"-->
<%
dim sqlRoot,rsRoot,trs,arrClassID,TitleStr
sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.ParentPath From PhotoClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ClassID & " or C.ParentPath like '%" & ParentPath & "," & ClassID & "%' and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.OrderID"
Set rsRoot= Server.CreateObject("ADODB.Recordset")
rsRoot.open sqlRoot,conn,1,1	
arrClassID=ClassID
do while not rsRoot.eof
	arrClassID=arrClassID & "," & rsRoot(0)
	rsRoot.movenext
loop
rsRoot.close
set rsRoot=nothing
FoundErr=False
if rs("PhotoLevel")<=999 then
	if UserLogined<>True then
		FoundErr=True
		ErrMsg=ErrMsg & "对不起，本图片为收费图片，要求至少是本站的注册用户才能欣赏！<br>您还没注册或者没有登录？所以不能欣赏本图片。请赶紧 <a href='User_Reg.asp'><font color=red><b>注册</b></font></a> 或 <a href='User_Login.asp'><font color=red><b>登录</a></font></a>吧！"
	else
		if UserLevel>rs("PhotoLevel") then
			FoundErr=True
			ErrMsg=ErrMsg & "对不起，本图片为收费图片，并且只有 <font color=blue>"
			if rs("PhotoLevel")=999 then
				ErrMsg=ErrMsg & "注册用户"
			elseif rs("PhotoLevel")=99 then
				ErrMsg=ErrMsg & "收费用户"
			elseif rs("PhotoLevel")=9 then
				ErrMsg=ErrMsg & "VIP用户"
			elseif rs("PhotoLevel")=5 then
				ErrMsg=ErrMsg & "管理员"
			end if
			ErrMsg=ErrMsg & "级别的用户</font> 才能欣赏。你目前的权限级别不够，所以不能欣赏。"
		else
			if ChargeType=1 and rs("PhotoPoint")>0 then
				if UserPoint<rs("PhotoPoint") then
					FoundErr=True
					ErrMsg=ErrMsg &"对不起，本图片为收费图片，并且欣赏本图片需要消耗 <b><font color=red>" & rs("PhotoPoint") & "</font></b> 点！"
					ErrMsg=ErrMsg &"而你目前只有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。点数不足，无法欣赏本图片。请与我们联系进行充值。"
				else
					ErrMsg=ErrMsg & "<font color=red><b>注意</b></font>：本图片为收费图片，并且欣赏本图片需要消耗 <font color=red><b>" & rs("PhotoPoint") & "</b></font> 点！&nbsp;&nbsp;若点击任何一个图片地址，意味着你同意消耗相应点数来欣赏此图片。若你不同意，请不要点击任何图片地址！"
				end if
			elseif ChargeType=2 then
				if ValidDays<=0 then
					FoundErr=True
					ErrMsg=ErrMsg & "<font color=red>对不起，本图片为收费图片，而您的有效期已经过期，所以无法欣赏本图片。请与我们联系进行充值。</font>"
				end if
			end if
		end if
	end if							
end if
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" align="left" valign="top" class="tdbg_leftall"><table width="180" border="0" cellspacing="0" cellpadding="0">
               <tr> 
          <td height="5"></td>
        </tr>
		<tr> 
          <td background="Images/left03.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr class="title_left"> 
                <td align="center" class="title_lefttxt"> <strong>图片搜索</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"> <div align="center"> 
                    <%call ShowSearchForm("Photo_Search.asp",1)%>
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <%if Child>0 then%>
        <tr> 
          <td background="Images/left12.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td align="center" class="title_lefttxt"> <strong><%=ClassName%>分类</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td height="80" valign="top"> <%call ShowChildClass(1)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <%end if%>
        <tr> 
          <td background="Images/left18.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td align="center" class="title_lefttxt"><strong>热门图片</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td height="80" valign="top"> <%call ShowHot(10,100)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left08.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="title_left"> 
                <td align="center" class="title_lefttxt"><strong>推荐图片</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td height="80" valign="top"> <%call ShowElite(10,100)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td class="title_left2"></td>
        </tr>
      </table></td>
    <td width="5" valign="top"><br> </td>
    <td width="575" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
      </table>
      <table width="99%" border="0" cellpadding="3" cellspacing="1" bgcolor="#666666">
        <tr>
          <td align="center" valign="top" bgcolor="#FFFFFF"><TABLE 
width="100%" border=0 align=center cellPadding=2 cellSpacing=0 class="tdbg_rightall">
              <TBODY>
                <TR> 
                  <TD align=middle><div align="center"><strong><%= dvhtmlencode(rs("PhotoName")) %></strong></div></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="100%" height="5" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td></td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td width="70">图片大小：</td>
                <td width="249"><%= rs("PhotoSize") & " K" %></td>
                <td width="240" rowspan="7" align="center" valign="middle"> 
                  <table width="150" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="10"><img src="Images/bg_0ltop.gif" width="10" height="10"></td>
                      <td height="10" background="Images/bg_01.gif"></td>
                      <td height="10"><img src="Images/bg_0rtop.gif" width="10" height="10"></td>
                    </tr>
                    <tr> 
                      <td width="10" background="Images/bg_03.gif">&nbsp;</td>
                      <td width="130" align="center" valign="middle" bgcolor="#FFFFFF">
                        <%
					  if FoundErr<>True then
						  response.write "<a href='Photo_Viewer.asp?UrlID=1&PhotoID=" & rs("PhotoID") & "'><img src='" & rs("PhotoUrl_Thumb") & "' border=0 width=150 height=113></a>"
					  else
						  response.write "<img src='" & rs("PhotoUrl_Thumb") & "' border=0 width=150 height=113>"
					  end if%>
                      </td>
                      <td width="10" background="Images/bg_04.gif">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="10"><img src="Images/bg_0lbottom.gif" width="10" height="10"></td>
                      <td height="10" background="Images/bg_02.gif"></td>
                      <td height="10"><img src="Images/bg_0rbottom.gif" width="10" height="10"></td>
                    </tr>
                  </table>
                  <%response.write "<a href='Photo_Viewer.asp?UrlID=1&PhotoID=" & rs("PhotoID") & "'>" & dvhtmlencode(rs("PhotoName")) & "</a>"%> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">图片作者：</td>
                <td><%
		if rs("Author")="" then
			response.write "佚名"
		else
			response.write dvhtmlencode(rs("Author")) 
		end if%></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">更新时间：</td>
                <td><%= rs("UpdateTime") %></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">推荐等级：</td>
                <td><%= string(rs("Stars"),"★") %></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">查看次数：</td>
                <td>本日：<%=rs("DayHits")%>&nbsp;&nbsp;本周：<%=rs("WeekHits")%>&nbsp;&nbsp;本月：<%=rs("MonthHits")%>&nbsp;&nbsp;总计：<%=rs("Hits")%></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">图片添加：</td>
                <td><%= rs("Editor") %></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="70">相关评论：</td>
                <td>无</td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td>图片简介：</td>
                <td colspan="2"><%= ubbcode(dvhtmlencode(rs("PhotoIntro"))) %></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td>查看图片：</td>
                <td colspan="2"><%
				if FoundErr=True then
					response.write ErrMsg
				else
					response.write "<a href='Photo_Viewer.asp?UrlID=1&PhotoID=" & rs("PhotoID") & "'>图片地址一" & "</a>"
					if rs("PhotoUrl2")<>"" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='Photo_Viewer.asp?UrlID=2&PhotoID=" & rs("PhotoID") & "'>图片地址二" & "</a>"
					if rs("PhotoUrl3")<>"" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='Photo_Viewer.asp?UrlID=3&PhotoID=" & rs("PhotoID") & "'>图片地址三" & "</a>"
					if rs("PhotoUrl4")<>"" then response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='Photo_Viewer.asp?UrlID=4&PhotoID=" & rs("PhotoID") & "'>图片地址四" & "</a>"
		  		end if
				%></td>
              </tr>
            </table>
              <BR>
            <TABLE align=center border=0 cellPadding=2 cellSpacing=0 width="100%">
              <TBODY> 
              <TR> 
                <TD class="tdbg_rightall"><strong>&nbsp;<img src="Images/TEAM.gif" width="16" height="16" align="absmiddle">&nbsp;网友评论：</strong>（只显示最新5条。评论内容只代表网友观点，与本站立场无关！）</TD>
                <TD class="tdbg_rightall" align="right">【<a href="Photo_Comment.asp?PhotoID=<%=rs("PhotoID")%>" target="_blank">发表评论</a>】</TD>
              </TR>
              </TBODY> 
            </TABLE>
              <TABLE align=center bgColor=#CCCCCC border=0 cellPadding=2 
            cellSpacing=1 style="WORD-BREAK: break-all" width="100%">
                <TBODY>
                  <TR> 
                    <TD bgColor=#fffffe style="LINE-HEIGHT: 16px"> 
                      <%call ShowComment(5)%>
                    </TD>
                  </TR>
                </TBODY>
              </TABLE>
			</td>
        </tr>
      </table>
      <table width="99%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td  height="13" align="center" valign="top"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
  <tr> 
    <td  height="13" align="center" valign="top"><table width="756" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="13" Class="tdbg_left2"></td>
        </tr>
      </table></td>
  </tr>
</table>
<% call Bottom() %>
<% call PopAnnouceWindow(400,300) %>
</body>
</html>
<%
rs.close
set rs=nothing
call CloseConn()
%>