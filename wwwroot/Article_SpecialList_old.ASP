<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
MaxPerPage=20
strFileName="Article_SpecialList.asp"
SkinID=0
PageTitle="专题列表"
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<% call MenuJS() %>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="Top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" align="left" valign="top" class="tdbg_leftall"> <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR class="title_left"> 
          <TD align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_lefttxt"><div align="center"><strong>用 户 登 录</strong></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="80" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top"> 
                  <% call ShowUserLogin() %>
                </td>
              </tr>
            </table></TD>
        </TR>
      </table>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR class="title_left"> 
          <TD align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_lefttxt"><div align="center"><b>热 门 文 章</b></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="80" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top">
                  <% call ShowHot(5,14) %>
                </td>
              </tr>
            </table></TD>
        </TR>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border="0" style="word-break:break-all">
        <TR class="title_left"> 
          <TD align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="title_lefttxt"><div align="center"><b>推 荐 文 章</b></div></td>
              </tr>
            </table></TD>
        </TR>
        <TR> 
          <TD height="80" valign="top" class="tdbg_left"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td valign="top">
                  <% call ShowElite(5,14) %>
                </td>
              </tr>
            </table></TD>
        </TR>
      </table>
	</td>
    <td width="5"></td>
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0" class="title_main">
        <tr> 
          <td width="40">&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0" >
              <tr> 
                <td class="title_maintxt"> 专题列表</td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" class="border">
        <tr> 
          <td valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
              <tr> 
                <td height="150" valign="top"> 
                  <% call ShowAllSpecial() %>
                </td>
              </tr>
            </table></td>
        </tr>
        <tr class="tdbg"> 
          <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tdbg_leftall"> 
                <td> 
                  <%
		  if TotalPut>0 then
		  	call showpage(strFileName,TotalPut,MaxPerPage,false,true,"个专题")
		  end if
		  %>
           </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
	  <table width='100%' border='0'cellpadding='0' cellspacing='0' style="word-break:break-all">
        <tr class="title_main" align='center'> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="title_main">
                    <tr> 
                      <td width="40">&nbsp;</td>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0" >
                          <tr> 
                            <td class="title_maintxt">最新热门图片 </td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        <tr> 
          <td height="80" valign="top"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5" class="border">
              <tr> 
                <td valign="top"> 
                  <% call ShowPicArticle(0,4,20,1,4,130,90,200,true,false) %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td  height="15" align="center" valign="top"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table></td>
        </tr>
      </table>
      
      <table width='100%' border='0'cellpadding='2' cellspacing='0'>
        <tr class='tdbg_leftall'> 
          <td width="18%"> <div align="center"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td class="title_maintxt"><div align="center">文章搜索</div></td>
                </tr>
              </table>
            </div></td>
          <td width="82%"> 
            <% call ShowSearchForm("Article_Search.asp",2) %>
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td  height="13" align="center" valign="top"><table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
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
rsSpecial.close
set rsSpecial=nothing
call CloseConn()
%>