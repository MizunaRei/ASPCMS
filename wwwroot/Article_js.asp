<!--#include file="inc/conn.asp" -->
<!--#include file="inc/function.asp" -->
<!--#include file="inc/ubbcode.asp" -->
<!--#include file="inc/config.asp" -->
<%
dim ClassID,IncludeChild,SpecialID,ArticleNum,TitleMaxLen,ContentMaxLen,ShowType,ShowCols
dim ShowProperty,ShowClassName,ShowIncludePic,ShowTitle,ShowUpdateTime,ShowHits,ShowAuthor,ShowHot,ShowMore
dim Hot,Elite,DateNum,OrderField,OrderType,ImgWidth,ImgHeight
dim SystemPath,rs,sql,str,str1,topicLen,topic
dim i,FileType
dim tClass,trs,arrClassID
dim tLayout,LayoutFileName_Class,LayoutFileName_Article
dim Author,AuthorName,AuthorEmail
ClassID=trim(request.querystring("ClassID"))
IncludeChild=trim(request.QueryString("IncludeChild"))
SpecialID=trim(request.querystring("SpecialID"))
ArticleNum=trim(request.querystring("ArticleNum"))
TitleMaxLen=trim(request.querystring("TitleMaxLen"))
ContentMaxLen=trim(request.querystring("ContentMaxLen"))
ShowType=trim(request.querystring("ShowType"))
ShowCols=trim(request.querystring("ShowCols"))
ShowProperty=trim(request.querystring("ShowProperty"))
ShowClassName=trim(request.querystring("ShowClassName"))
ShowIncludePic=trim(request.querystring("ShowIncludePic"))
ShowTitle=trim(request.querystring("ShowTitle"))
ShowUpdateTime=trim(request.querystring("ShowUpdateTime"))
ShowHits=trim(request.querystring("ShowHits"))
ShowAuthor=trim(request.querystring("ShowAuthor"))
ShowHot=trim(request.querystring("ShowHot"))
ShowMore=trim(request.querystring("ShowMore"))
Hot=trim(request.querystring("Hot"))
Elite=trim(request.querystring("Elite"))
DateNum=trim(request.querystring("DateNum"))
OrderField=trim(request.querystring("OrderField"))
OrderType=trim(request.querystring("OrderType"))
ImgWidth=trim(request.querystring("ImgWidth"))
ImgHeight=trim(request.querystring("ImgHeight"))

SystemPath="http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"article_js.asp","")
if ShowType<>"" then
	ShowType=Cint(ShowType)
else
	ShowType=1
end if
if ShowCols<>"" then
	ShowCols=Cint(ShowCols)
else
	ShowCols=1
end if
if ContentMaxLen<>"" then
	ContentMaxLen=Cint(ContentMaxLen)
else
	ContentMaxLen=200
end if
if ClassID="" then
	ClassID=0
else
	ClassID=Clng(ClassID)
end if
if SpecialID="" then
	SpecialID=0
else
	SpecialID=Clng(SpecialID)
end if

set tLayout=conn.execute("select LayoutFileName from Layout where LayoutType=2 and IsDefault=True")
LayoutFileName_Class=tLayout(0)
if isNull(LayoutFileName_Class) then LayoutFileName_Class="ShowClass.asp"
set tLayout=conn.execute("select LayoutFileName from Layout where LayoutType=3 and IsDefault=True")
LayoutFileName_Article=tLayout(0)
if isNull(LayoutFileName_Article) then LayoutFileName_Article="ShowArticle.asp"
set tLayout=nothing

sql="select"
if ArticleNum<>"" then sql=sql & " top " & Cint(ArticleNum)
if ShowClassName="true" then
	sql=sql & " A.ArticleID,A.ClassID,C.ClassName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,A.Content,"
	sql=sql & " A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sql=sql & " inner join ArticleClass C on A.ClassID=C.ClassID where A.Deleted=False and A.Passed=True"
else
	sql=sql & " A.ArticleID,A.ClassID,L.LayoutID,L.LayoutFileName,A.Title,A.Key,A.Author,A.CopyFrom,A.UpdateTime,A.Editor,A.TitleFontColor,A.TitleFontType,A.Content,"
	sql=sql & " A.Hits,A.OnTop,A.Hot,A.Elite,A.Passed,A.IncludePic,A.Stars,A.PaginationType,A.ReadLevel,A.ReadPoint,A.DefaultPicUrl from Article A"
	sql=sql & " inner join Layout L on A.LayoutID=L.LayoutID where A.Deleted=False and A.Passed=True"
end if

if ClassID>0 then
	set tClass=conn.execute("select ClassID,ParentPath,Child From ArticleClass where ClassID=" & ClassID)
	if tClass.bof and tClass.eof then
		response.write "document.write (" & Chr(34) & "找不到指定的栏目，可能已经被管理员删除！请重新生成JS调用代码。" & Chr(34) & ");"
		response.end
	else
		if IncludeChild="true" then
			if tClass(2)>0 then
				arrClassID=tClass(0)
				set trs=conn.execute("select ClassID from ArticleClass where ParentID=" & tClass(0) & " or ParentPath like '%" & tClass(1) & "," & tClass(0) & ",%' and Child=0 and LinkUrl=''")
				do while not trs.eof
					arrClassID=arrClassID & "," & trs(0)
					trs.movenext
				loop
				set trs=nothing	
				sql=sql & " and A.ClassID in (" & arrClassID & ")"
			else
				sql=sql & " and A.ClassID=" & Clng(ClassID)
			end if
		else
			sql=sql & " and A.ClassID=" & Clng(ClassID)
		end if
	end if
	set tClass=nothing	
end if
if SpecialID>0 then sql=sql & " and A.SpecialID=" & SpecialID
if ShowType=3 or ShowType=4 then sql=sql & " and A.DefaultPicUrl<>''"
if Hot="true" then sql=sql & " and A.Hits>=" & HitsOfHot
if Elite="true" then sql=sql & " and A.Elite=True"
if DateNum<>"" then sql=sql & " and DATEDIFF('d',UpdateTime,Date())<=" & Cint(DateNum)
sql=sql & " order by A.OnTop asc"
if OrderField<>"" then sql=sql & " , A." & OrderField
if OrderType<>"" then
	sql=sql & " " & OrderType
else
	sql=sql & " asc"
end if

set rs=server.createObject("Adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then 
	response.write "document.write (" & Chr(34) & "没有符合条件的文章" & Chr(34) & ");"	
else
	if ShowType=3 or ShowType=4 then
		response.write "document.write (" & Chr(34) & "<table cellspacing='5' align='center'><tr valign='top'>" & Chr(34) & ");" & vbcrlf
	end if
	do while not rs.eof
		if TitleMaxLen<>"" then
			topic=gotTopic(rs("title"),Cint(TitleMaxLen))
		else
			topic=rs("title")
		end if
		if rs("TitleFontType")=1 then
			topic="<b>" & topic & "</b>"
		elseif rs("TitleFontType")=2 then
			topic="<em>" & topic & "</em>"
		elseif rs("TitleFontType")=3 then
			topic="<b><em>" & topic & "</em></b>"
		end if
		if rs("TitleFontColor")<>"" then
			topic="<font color='" & rs("TitleFontColor") & "'>" & topic & "</font>"
		end if
		Author=rs("Author")
		if instr(Author,"|")>0 then
			AuthorName=left(Author,instr(Author,"|")-1)
			AuthorEmail=right(Author,len(Author)-instr(Author,"|")-1)
		else
			AuthorName=Author
			AuthorEmail=""
		end if

		if ShowType=1 or ShowType=2 then
			str=""
		else
			str="<td width='" & ImgWidth & "' align='center'>"
		end if
		
		if ShowType=1 or ShowType=2 then
			if ShowProperty="true" then
				if rs("OnTop")=true then
					str=str & "<img src='images/soul.gif' alt='固顶文章'>&nbsp;"
				elseif rs("Elite")=true then
					str=str & "<img src='images/soul.gif' alt='推荐文章'>&nbsp;"
				else
					str=str & "<img src='images/soul.gif' alt='普通文章'>&nbsp;"
				end if
			end if
			if ShowIncludePic="true" and rs("IncludePic")=true then
				str=str & "<font color=blue>[图文]</font>"
			end if
			if ShowClassName="true" then
				str=str & "[<a href='" & SystemPath & LayoutFileName_Class & "?ClassID=" & rs("ClassID") & "'>" & rs("ClassName") & "</a>]"
				str=str & "<a href='" & SystemPath & LayoutFileName_Article & "?ArticleID=" & rs("articleid") & "' title='文章标题：" & rs("Title") & "\n" & "作    者：" & AuthorName & "\n" & "更新时间：" & rs("UpdateTime") & "\n" & "点击次数：" & rs("Hits") & " ' target='_blank'>"
			else
	      		str=str & "<a href='" & SystemPath & rs("LayoutFileName") & "?ArticleID=" & rs("articleid") & "' title='文章标题：" & rs("Title") & "\n" & "作    者：" & AuthorName & "\n" & "更新时间：" & rs("UpdateTime") & "\n" & "点击次数：" & rs("Hits") & " ' target='_blank'>"
			end if
			str=str & Topic & "</a>"
		else
			if ShowClassName="true" then
				str=str & "[<a href='" & SystemPath & LayoutFileName_Class & "?ClassID=" & rs("ClassID") & "'>" & rs("ClassName") & "</a>]"
				str=str & "<a href='" & SystemPath & LayoutFileName_Article & "?ArticleID=" & rs("articleid") & "' title='文章标题：" & rs("Title") & "\n" & "作    者：" & AuthorName & "\n" & "更新时间：" & rs("UpdateTime") & "\n" & "点击次数：" & rs("Hits") & " ' target='_blank'>"
			else
	      		str=str & "<a href='" & SystemPath & rs("LayoutFileName") & "?ArticleID=" & rs("articleid") & "' title='文章标题：" & rs("Title") & "\n" & "作    者：" & AuthorName & "\n" & "更新时间：" & rs("UpdateTime") & "\n" & "点击次数：" & rs("Hits") & " ' target='_blank'>"
			end if
			
			FileType=right(lcase(rs("DefaultPicUrl")),3)
			if FileType="swf" then
				str=str & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='" & ImgWidth & "' height='" & ImgHeight & "'><param name='movie' value='" & rs("DefaultPicUrl") & "'><param name='quality' value='high'><embed src='" & rs("DefaultPicUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='" & ImgWidth & "' height='" & ImgHeight & "'></embed></object>"
			elseif fileType="jpg" or fileType="bmp" or fileType="png" or fileType="gif" then
				str=str & "<img src='" & rs("DefaultPicUrl") & "' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'>"
			else
				str=str & "<img src='http://www.fanchen.com/images/NoPic2.jpg' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'>"
			end if
			str=str & Topic & "</a>"
		end if
		if ShowAuthor="true" or ShowUpdateTime="true" or ShowHits="true" then
			str=str & "&nbsp;/&nbsp;"
			if ShowAuthor="true" then
				if AuthorEmail="" then
					str=str & AuthorName
				else
					str=str & "<a href='mailto:" & AuthorEmail & "'>" & AuthorName & "</a>"
				end if
			end if
			if ShowUpdateTime="true" then
				if ShowAuthor="true" then
					str=str & "，"
				end if
				if CDate(FormatDateTime(rs("UpdateTime"),2))=date() then
					str=str & "<font color=red>"
				else
					str=str & "<font color=#999999>"
				end if
				str=str & FormatDateTime(rs("UpdateTime"),1) & "</font>"
			end if
			if ShowHits="true" then
				if ShowAuthor="true" or ShowUpdateTime="true" then
					str=str & "，"
				end if
				str=str & rs("Hits")
			end if
			str=str & ""
		end if
		if ShowHot="true" and rs("Hits")>=HitsOfHot then
			str=str & "<img src='" & SystemPath & "images/hot.gif' align='absmiddle' alt='热点文章'>"
		end if
		if ShowType=1 then
			str=str & "<br>"
		elseif ShowType=2 then
			str=str & "<br><div style='padding:0px 20px'>" & left(replace(replace(replace(nohtml(rs("content")),chr(34),""),chr(10),"\n"),chr(13),"\n"),ContentMaxLen) & "……</div>"
		elseif ShowType=3 then
			str=str & "</td>"
		elseif ShowType=4 then
			str=str & "</td><td>" & left(replace(replace(replace(nohtml(rs("content")),chr(34),""),chr(10),"\n"),chr(13),"\n"),ContentMaxLen) & "……</td>"
		end if
		rs.movenext
		if ShowType=3 or ShowType=4 then
			i=i+1
			if ((i mod ShowCols=0) and (not rs.eof)) then
				str=str & "</tr><tr valign='top'>"
			end if
		end if
		response.write "document.write (" & Chr(34) & str & Chr(34) & ");" & vbcrlf
	loop
	if ShowType=1 or ShowType=2 then
		if ShowMore="true" then
			if ClassID>0 then
				str="<div align='right'><a href='" & SystemPath & LayoutFileName_Class & "?ClassID=" & ClassID & "'>more...</a></div>"
			else
				str="<div align='right'><a href='http://www.fanchen.com/index.asp'>more...</a></div>"
			end if
			response.write "document.write (" & Chr(34) & str & Chr(34) & ");" & vbcrlf
		end if
	else
		str="</tr>"
		if ShowMore="true" then
			if ClassID>0 then
				str=str & "<tr><td colspan='" & ShowCols*2 & "' align='right'><a href='" & SystemPath & LayoutFileName_Class & "?ClassID=" & ClassID & "'>更多……</a></td>"
			else
				str=str & "<tr><td colspan='" & ShowCols*2 & "' align='right'><a href='index.asp'>更多……</a></td>"
			end if
		end if
		str=str & "</table>"
		response.write "document.write (" & Chr(34) & str & Chr(34) & ");" & vbcrlf
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>