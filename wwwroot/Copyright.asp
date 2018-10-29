<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=0
Const ShowRunTime="Yes"
MaxPerPage=20
SkinID=0
PageTitle="碧聊文学原创文学网"
Set rsArticle= Server.CreateObject("ADODB.Recordset")
Set rsPic= Server.CreateObject("ADODB.Recordset")
%>
<html>
<head>
<title><%=strPageTitle & " >> " & PageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center"> 
        <p class="tdbg_rightall"><font color="#CC0000" size="3">【&nbsp;关于碧聊文学原创文学家园&nbsp;】</font> 
        </p>
      </div></td>
  </tr>
  <tr> 
    <td height="347"> 
      <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="347"> 
            <div align="center"> </div>
            <p>1、<font color="#CC0000" size="3">碧聊文学原创文学家园</font>是一个以原创网络文学为主，其他娱乐为辅的综合文学站点。</p>
            <p>2、<font color="#CC0000" size="3">碧聊文学原创文学家园</font>目前所有的服务都是免费并且完善的。</p>
            <p>3、<font color="#CC0000" size="3">碧聊文学原创文学家园</font>的目标是成为国内知名的文学网站之一。</p>
            <p>4、<font color="#CC0000" size="3">碧聊文学原创文学家园</font>成立于2006年1月1日。虽然时间不长，但是我们一直用心在做，服务内容并不压于同类网站。</p>
            <p>5、<font color="#CC0000" size="3">碧聊文学原创文学家园</font>提供社区、原创相册、下载等服务，以后还将推出在线书店。 
            </p>
            <p>6、如果您<font color="#CC0000" size="3">碧聊文学原创文学家园</font>有任何意见和建议请您按以下地址发给我们提议。谢谢！</p>
            <p>7、未尽事宜以<font color="#CC0000" size="3">碧聊文学原创文学家园</font>最新公告和国家相关法律为准。<br>
            </p>
            <p>**********************************************************************<br>
              * 主编信箱: <a href="mailto:10000@26265.cn">antishy</a><br>
              * 主页地址: <a href="http://www.26265.cn/">http://www.26265.cn/</a><br>
              * 论坛地址: <a href="http://www.26265.cn/bbs">http://www.26265.cn/bbs</a><br>
              **********************************************************************<br>
              <br>
              <br>
            </p>
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tdbg">
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
set rsArticle=nothing
set rsPic=nothing
call CloseConn()
%>