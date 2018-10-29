<!--#include file="Inc/syscode_Photo.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=4
Const ShowRunTime="Yes"
MaxPerPage=20
strFileName="Photo_Class.asp?ClassID=" & ClassID
Set rsPhoto= Server.CreateObject("ADODB.Recordset")
Set rsPic= Server.CreateObject("ADODB.Recordset")
SkinID=0
%>
<html>
<head>
<title><%=strPageTitle%></title>
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
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="border2">
  <tr> 
    <td width="180" align="left" valign="top" class="tdbg_leftall"><table width="180" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"></td>
        </tr>
		<%if Child>0 then%>
        <tr>
          <td background="Images/left12.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
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
              <tr> 
                <td align="center" class="title_lefttxt"><strong>热门图片</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <%call ShowHot(10,100)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
        <tr> 
          <td background="Images/left08.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center" class="title_lefttxt"><strong>推荐图片</strong></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top" class="tdbg_left"><table width="100%" height="100%" border="0" cellpadding="8">
              <tr> 
                <td valign="top"> <%call ShowElite(10,100)%> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td class="title_left2"></td>
        </tr>
      </table></td>
    <td width="5" valign="top">&nbsp;</td>
    <td width="575" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="title_main">
              <tr> 
                <td width="40">&nbsp;</td>
                <td valign="bottom" class="title_maintxt"><%=ClassName%> 图片列表</td>
              </tr>
            </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="5" class="border">
              <tr> 
                <td height="100" valign="top"><%call ShowPhoto(100,arrClassID)%></td>
              </tr>
              <tr>
                <td valign="top">
                  <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"张图片")
		  end if
		  %>
                </td>
              </tr>
            </table></td>
	</tr>
	<tr>
	    <td>
	      <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="13" Class="tdbg_left2"></td>
              </tr>
            </table>
		</td>
    </tr>
    <tr>
       <td><table width='100%' border='0' align="center"cellpadding='2' cellspacing='0' class="tdbg_rightall">
              <tr> 
                <td width="100" align="center" class="title_maintxt"><img src="Images/checkphoto.gif" width="15" height="15" align="absmiddle">&nbsp;&nbsp;图片搜索：</td>
                <td> <div align="center"> 
                    <% call ShowSearchForm("Photo_Search.asp",2) %>
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table>
      
    </td>
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
set rsPhoto=nothing
set rsPic=nothing
call CloseConn()
%>