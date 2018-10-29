<!--#include file="Inc/syscode_Photo.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=4
Const ShowRunTime="Yes"
MaxPerPage=20
SkinID=0
PageTitle="查看图片"
dim UrlID
UrlID=trim(request("UrlID"))
if UrlID="" then
	UrlID=1
else
	UrlID=Clng(UrlID)
end if
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
				if Request.Cookies("asp163")("Pay_Photo" & PhotoID)<>"yes" then
					if UserPoint<rs("PhotoPoint") then
						FoundErr=True
						ErrMsg=ErrMsg &"对不起，本图片为收费图片，并且欣赏本图片需要消耗 <b><font color=red>" & rs("PhotoPoint") & "</font></b> 点！"
						ErrMsg=ErrMsg &"而你目前只有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。点数不足，无法欣赏本图片。请与我们联系进行充值。"
					else
						if lcase(trim(request("Pay")))="yes" then
							Conn_User.execute "update " & db_User_Table & " set " & db_User_UserPoint & "=" & db_User_UserPoint & "-" & rs("PhotoPoint") & " where " & db_User_Name & "='" & UserName & "'"
							response.Cookies("asp163")("Pay_Photo" & PhotoID)="yes"
						else
							FoundErr=True
							ErrMsg=ErrMsg & "<font color=red><b>注意</b></font>：欣赏本图片需要消耗 <font color=red><b>" & rs("PhotoPoint") & "</b></font>"
							ErrMsg=ErrMsg &"你目前尚有 <b><font color=blue>" & UserPoint & "</font></b> 点可用。阅读本文后，你将剩下 <b><font color=green>" & UserPoint-rs("PhotoPoint") & "</font></b> 点"
							ErrMsg=ErrMsg &"<br><br>你确实愿意花费 <b><font color=red>" & rs("PhotoPoint") & "</font></b> 点来欣赏本图片吗？"
							ErrMsg=ErrMsg &"<br><br><a href='Photo_Viewer.asp?Pay=yes&UrlID=" & UrlID & "&PhotoID=" & PhotoID & "'>我愿意</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='index.asp'>我不愿意</a></p>"
						end if
					end if
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
if FoundErr<>True then
		rs("Hits")=rs("Hits")+1
		if datediff("D",rs("LastHitTime"),now())<=0 then
			rs("DayHits")=rs("DayHits")+1
		else
			rs("DayHits")=1
		end if
		if datediff("ww",rs("LastHitTime"),now())<=0 then
			rs("WeekHits")=rs("WeekHits")+1
		else
			rs("WeekHits")=1
		end if
		if datediff("m",rs("LastHitTime"),now())<=0 then
			rs("MonthHits")=rs("MonthHits")+1
		else
			rs("MonthHits")=1
		end if
		rs("LastHitTime")=now()
		rs.update
end if
%>
<html>
<head>
<title><%=PhotoTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
<script language=JavaScript>
function smallit(){            
	var height1=PhotoViewer.images1.height;            
	var width1=PhotoViewer.images1.width;            
	PhotoViewer.images1.height=height1/1.2;            
	PhotoViewer.images1.width=width1/1.2;           
}             
          
function bigit(){            
	var height1=PhotoViewer.images1.height;            
	var width1=PhotoViewer.images1.width;            
	PhotoViewer.images1.height=height1*1.2;          
	PhotoViewer.images1.width=width1*1.2;           
}             
function fullit()
{
	var width_s=screen.width-10;
	var height_s=screen.height-30;
	window.open("Photo_View.asp?UrlID=<%=UrlID%>&PhotoID=<%=PhotoID%>", "PhotoView", "width="+width_s+",height="+height_s+",left=0,top=0,location=no,toolbar=no,status=no,resizable=no,scrollbars=yes,menubar=no,directories=no");
}
function realsize()
{
	PhotoViewer.images1.height=PhotoViewer.images2.height;     
	PhotoViewer.images1.width=PhotoViewer.images2.width;
	PhotoViewer.block1.style.left = 0;
	PhotoViewer.block1.style.top = 0;
	
}
function featsize()
{
	var width1=PhotoViewer.images2.width;            
	var height1=PhotoViewer.images2.height;            
	var width2=760;            
	var height2=500;            
	var h=height1/height2;
	var w=width1/width2;
	if(height1<height2&&width1<width2)
	{
		PhotoViewer.images1.height=height1;            
		PhotoViewer.images1.width=width1;           
	}
	else
	{
		if(h>w)
		{
			PhotoViewer.images1.height=height2;          
			PhotoViewer.images1.width=width1*height2/height1;           
		}
		else
		{
			PhotoViewer.images1.width=width2;           
			PhotoViewer.images1.height=height1*width2/width1;          
		}
	}
	PhotoViewer.block1.style.left = 0;
	PhotoViewer.block1.style.top = 0;
}
</script>         
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="Top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr>
    <td height="40" align="center"><strong><font size="4"><%= dvhtmlencode(rs("PhotoName")) %></font></strong></td>
  </tr>
  <tr class="tdbg_rightall">
    <td align="center"><table width="100%" border="0" cellspacing="3" cellpadding="0">
        <tr> 
          <td><div align="center">图片大小：<%= rs("PhotoSize") & " K" %></div></td>
          <td><div align="center">图片作者： 
              <%
		if rs("Author")="" then
			response.write "佚名"
		else
			response.write dvhtmlencode(rs("Author")) 
		end if%>
            </div></td>
          <td><div align="center">更新时间：<%= rs("UpdateTime") %></div></td>
          <td><div align="center">推荐等级：<font color="#009900"><%= string(rs("Stars"),"★") %></font></div></td>
          <td>查看次数：<%=rs("Hits")%></td>
        </tr>
      </table></td>
  </tr>
  <%if rs("PhotoUrl2")<>"" or rs("PhotoUrl3")<>"" or rs("PhotoUrl3")<>"" then%>
  <tr>
    <td align="center" valign="middle" class="tdbg_leftall"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr align="center">
        <td><%
		if UrlID=1 then
			response.write "<font color=red>图片地址一</font>"
		else
			response.write "<a href='Photo_Viewer.asp?UrlID=1&PhotoID=" & rs("PhotoID") & "'>图片地址一" & "</a>"
		end if
		  %>
        </td>
		<%if rs("PhotoUrl2")<>"" then %>
        <td><%
		if UrlID=2 then
			response.write "<font color=red>图片地址二</font>"
		else
			response.write "<a href='Photo_Viewer.asp?UrlID=2&PhotoID=" & rs("PhotoID") & "'>图片地址二" & "</a>"
		end if
		  %>
        </td>
		<%
		end if
		if rs("PhotoUrl3")<>"" then %>
        <td><%
		if UrlID=3 then
			response.write "<font color=red>图片地址三</font>"
		else
			response.write "<a href='Photo_Viewer.asp?UrlID=3&PhotoID=" & rs("PhotoID") & "'>图片地址三" & "</a>"
		end if
		  %>
        </td>
		<%
		end if
		if rs("PhotoUrl4")<>"" then %>
        <td><%
		if UrlID=4 then
			response.write "<font color=red>图片地址四</font>"
		else
			response.write "<a href='Photo_Viewer.asp?UrlID=4&PhotoID=" & rs("PhotoID") & "'>图片地址四" & "</a>"
		end if
		  %>
        </td>
		<%end if%>
      </tr>
    </table></td>
  </tr>
  <%end if%>
  <tr>
    <td height="30" align="center" valign="middle" class="tdbg_leftall"><input name="smallit" type="button" id="smallit" onclick="smallit();" value="- 缩小 -">    
&nbsp;&nbsp;
<input name="bigit" type="button" id="bigit" onclick="bigit();" value="+ 放大 +">              
          
&nbsp;
<input name="fullit" type="button" id="fullit" value="全屏显示" onClick="fullit();">
&nbsp;
<input name="realsize" type="button" id="realsize" value="实际大小" onClick="realsize();">
&nbsp;
<input name="featsize" type="button" id="featsize" value="最合适大小" onClick="featsize();"></td>
  </tr>
  <tr> 
    <td height="500" align="center" valign="middle" class="tdbg_leftall"><%
	if FoundErr=True then
		response.write ErrMsg
	else
		response.write "<iframe id='PhotoViewer' width='99%' height='500' scrolling='no' src='Photo_View.asp?UrlID=" & UrlID & "&PhotoID=" & PhotoID &"'></iframe>"
	end if%></td>
  </tr>
  <tr>
    <td height="30" align="center" valign="middle" class="tdbg_leftall"><input name="smallit" type="button" id="smallit" onclick="smallit();" value="- 缩小 -">    
&nbsp;&nbsp;
<input name="bigit" type="button" id="bigit" onclick="bigit();" value="+ 放大 +">              
          
&nbsp;
<input name="fullit" type="button" id="fullit" value="全屏显示" onClick="fullit();">
&nbsp;
<input name="realsize" type="button" id="realsize" value="实际大小" onClick="realsize();">
&nbsp;
<input name="featsize" type="button" id="featsize" value="最合适大小" onClick="featsize();"></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="topborder">
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