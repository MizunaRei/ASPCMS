<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
SkinID=0
dim sqlLink,rsLink,LinkType
LinkType=trim(request("LinkType"))
strFileName="FriendSite.asp?LinkType=" & LinkType
MaxPerPage=20
PageTitle="友情链接"

sqlLink="select * from FriendSite where IsOK=True "
if LinkType<>"" then
	LinkType=CInt(LinkType)
	if LinkType=1 then
		sqlLink=sqlLink & " and LinkType=1 "
	elseif LinkType=2 then
		sqlLink=sqlLink & " and LinkType=2 "
	end if
end if
sqlLink=sqlLink & " order by id desc"
set rsLink=server.createobject("adodb.recordset")
rsLink.open sqlLink,conn,1,1
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<div align="center">
<table width="760" border="0" align="center" valign="center" cellpadding="0" cellspacing="0" background="images/m_bg.gif">
  <tr> 
    <td>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="404"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                  <td>&nbsp;&nbsp;&nbsp;&nbsp;我 们 的 朋 友 &gt;&gt;</td>
              </tr>
            </table>          
          </td>
          <td width="356" align="right"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_maintxt"><div align="center"><a href="FriendSiteReg.asp">申请链接</a>&nbsp;|&nbsp;<a href="FriendSite.asp?LinkType=2">文字链接</a>&nbsp;|&nbsp;<a href="FriendSite.asp?LinkType=1">LOGO链接</a>&nbsp;|&nbsp;<a href="FriendSite.asp">所有链接</a></div></td>
              </tr>
            </table>
          </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="102"><%
  	if rsLink.eof and rsLink.bof then
		response.write "<table width='760'align='center'><tr><td height='50'>目前共有 0 个友情链接</td></tr></table>"
	else
    	totalPut=rsLink.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个站点"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsLink.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsLink.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个站点"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个站点"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
      <table width="740" border="0" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
        <tr background="images/fcbg1.gif"> 
          
          <td width="109" height="22" align="center" background="images/fcbg1.gif"><font color="#FF0000">网站名称</font></td>
          <td width="153" height="22" align="center" background="images/fcbg1.gif"><font color="#FF0000">网站LOGO</font></td>
          <td width="417" height="22" align="center" background="images/fcbg1.gif"><font color="#FF0000">网站简介</font></td>
          <td width="60" height="22" align="center" background="images/fcbg1.gif"><font color="#FF0000">站长</font></td>
         
        </tr>
        <%
	do while not rsLink.eof
%>
        <tr> 
         
          <td width="109" align="center" background="images/fcbg2.gif"><a href="<%=rsLink("SiteUrl")%>" target='blank' title="<%=rsLink("SiteUrl")%>"><%=rsLink("SiteName")%></a></td>
          <td width="153" align="center" background="images/fcbg2.gif"> 
            <%
if rsLink("LinkType")=1 then
	if rsLink("LogoUrl")<>"" and rsLink("LogoUrl")<>"http://" then
		if lcase(right(rsLink("LogoUrl"),3))="swf" then
			Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='88' height='31'><param name='movie' value='" & rsLink("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsLink("LogoUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='88' height='31'></embed></object>"
		else
			response.write "<a href='" & rsLink("SiteUrl") & "' target='_blank' title='" & rsLink("LogoUrl") & "'><img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'></a>"
		  end if
	else
		response.write "<a href='" & rsLink("SiteUrl") & "' target='_blank'><img src='images/nologo.gif' width='88' height='31' border='0'></a>"
	end if
else
	response.write "&nbsp;"
end if
%>
          </td>
          <td background="images/fcbg2.gif"><%=rsLink("SiteIntro")%></td>
          <td width="60" align="center" background="images/fcbg2.gif"><a href="mailto:<%=rsLink("Email")%>"><%=rsLink("SiteAdmin")%></a></td>
          
        </tr>
        <%
	i=i+1
	if i>=MaxPerPage then exit do
	rsLink.movenext
loop
%>
      </table>
      <table width="740" border="0" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
        <tr>
          <td height="19" background="images/fcbg1.gif">
            <%
end sub 
%>
          </td>
        </tr>
      </table><br>
      </td>
  </tr>
</table>


</div>
</body>
</html>
<%
rsLink.close
set rsLink=nothing
call CloseConn()
%>
