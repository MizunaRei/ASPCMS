<!--#include file="Inc/syscode_article.asp"-->
<%
'请勿改动下面这三行代码
const ChannelID=2
Const ShowRunTime="Yes"
MaxPerPage=20
strFileName="Article_Class.asp?ClassID=" & ClassID & "&SpecialID=" & SpecialID
Set rsArticle= Server.CreateObject("ADODB.Recordset")
Set rsPic= Server.CreateObject("ADODB.Recordset")
%>
<html>
<head>
<title><%=strPageTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Skin_CSS.asp"-->
<%call MenuJS()%>
</head>
<body <%=Body_Label%> onmousemove='HideMenu()'>
<!--#include file="Top1.asp"-->
<%
dim sqlRoot,rsRoot,trs,arrClassID,TitleStr
sqlRoot="select C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID,C.Child,C.ParentPath From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & ClassID & " and C.IsElite=True and C.LinkUrl='' and C.BrowsePurview>=" & UserLevel & " order by C.OrderID"
Set rsRoot= Server.CreateObject("ADODB.Recordset")
rsRoot.open sqlRoot,conn,1,1
%>
<table width="760" height="393" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="180" height="393" align="left" valign="top" background="skin/3/tdbg_left.gif">
	 <br><table width="178" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="175" background="images/fc1.gif">&nbsp;&nbsp;推 荐 文 章</td>
        </tr><tr>
          <td width="175" background="images/fc3.gif">&nbsp;</td>
        </tr>
        <tr>
          <td height="18" background="images/fc.gif" align="top">
            <table width="83%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="18">
                  <% call ShowElite(10,16) %>
                </td>
              </tr>
            </table>
            
          </td>
        </tr>
        <tr>
          <td height="31" background="images/fc2.gif">&nbsp;</td>
        </tr>
      </table><table width="178" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="175" background="images/fc01.gif">&nbsp;&nbsp;图 片 文 学</td>
        </tr><tr>
          <td width="175" background="images/fc3.gif">&nbsp;
		 </td>
		 
        </tr>
        <tr>
          <td background="images/fc.gif">&nbsp;<% call ShowPicArticle(ClassID,2,10,1,2,70,60,200,true,false) %>
          </td>
        </tr>
        <tr>
          <td height="39" background="images/fc4.gif">&nbsp;</td>
       
        </tr>
      </table>
    
    </td>
    <td width="5"></td>
    <td width="575" valign="top" background="images/m_bg.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0" valign="top" align="top"><tr><td height="393" valign="top">
<%
	if rsRoot.bof and rsRoot.eof then
	%>
      <table width="97%" border="0" valign="top" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
        <tr> 
          
          <td background="images/fcbg1.gif" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;<%=ClassName%> 文章列表</td>
        </tr>
      </table>
      <table width="97%" border="0" valign="top" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
        <tr> 
          <td height="198" valign="top" background="images/fcbg2.gif"> 
            <%call ShowArticle(30)%>
          </td>
        </tr>
      </table>
     
      <%
		  if totalput>0 then
		  	call showpage(strFileName,totalput,MaxPerPage,false,true,"篇文章")
		  end if
		  %>
      <%
	else
		do while not rsRoot.eof
%>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="5" bgcolor="#FFFFFF">
              <tr> 
                
                <td background="images/fcbg1.gif">&nbsp;&nbsp;&nbsp;&nbsp;<%
				arrClassID=rsRoot(0)
				response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>" & rsRoot(1) & "</a>"
				if rsRoot(5)>0 then
					response.write "："
					set trs=conn.execute("select top 4 C.ClassID,C.ClassName,C.RootID,L.LayoutFileName,L.LayoutID From ArticleClass C inner join Layout L on C.LayoutID=L.LayoutID where C.ParentID=" & rsRoot(0) & " and C.IsElite=True and C.LinkUrl=''  and C.BrowsePurview>=" & UserLevel & " order by C.OrderID")
					do while not trs.eof
						response.write "&nbsp;&nbsp;<a href='" & trs(3) & "?ClassID=" & trs(0) & "'>" & trs(1) & "</a>"
						trs.movenext
					loop
					set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & rsRoot(0) & " or ParentPath like '%" & rsRoot(6) & "," & rsRoot(0) & ",%' and Child=0 and LinkUrl='' and BrowsePurview>=" & UserLevel)
					do while not trs.eof
						arrClassID=arrClassID & "," & trs(0)
						trs.movenext
					loop
				end if
				%> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="57" align="center"><table bgcolor="#FFFFFF" width="97%" border="0" align="center" cellpadding="0" cellspacing="5">
                    <tr> 
                      <td height="24" colspan="2" valign="top"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td background="images/fcbg2.gif">
                              <%
sql="select top 5 A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,"
sql=sql & "A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True and A.ClassID in (" & arrClassID & ")  order by A.ArticleID desc"
rsArticle.open sql,conn,1,1
if rsArticle.bof and  rsArticle.eof then
	response.write "<li>没有任何文章</li>"
else
	call ArticleContent(30,True,True,True,1,True,True)
end if
rsArticle.close
				%>
                            </td>
                          </tr>
                        </table> 
                        
                      </td>
                    </tr>
                    <tr> 
                      <td width="510" height="18" valign="top" background="images/fcbg2.gif"><div align="right"> 
                        </div></td>
                      <td width="53" valign="top" background="images/fcbg2.gif">
                        <%response.write "<a href='" & rsRoot(3) & "?ClassID=" & rsRoot(0) & "'>more...</a>"%>
                      </td>
                    </tr>
                  </table></td>
        </tr>
      </table>
     
            <%
			rsRoot.movenext
		loop
	end if
	rsRoot.close
	set rsRoot=nothing
%>
            <table width='97%' border='0' align="center" cellpadding='0' cellspacing='5' bgcolor="#FFFFFF">
        <tr> 
          <td width="19%" height="18" background="images/fcbg1.gif"> <div align="center">站内文章搜索</div></td>
          <td width="81%"戀彧敬瑦????	???????? background="images/fcbg1.gif"> 
            <% call ShowSearchForm("Article_Search.asp",2) %>
          </td>
        </tr>
      </table>
	  </td></tr></table>
	  </td>
  </tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" class="tdbg">
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
<% call ShowAD(0) %>
<% call ShowAD(4) %>
<% call ShowAD(5) %>
</body>
</html>
<%
set rsArticle=nothing
set rsPic=nothing
call CloseConn()
%>