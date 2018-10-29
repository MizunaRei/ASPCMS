<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/conn.asp"-->
<%
dim ChannelID
ChannelID=trim(request("ChannelID"))
if ChannelID="" then
	ChannelID=0
else
	ChannelID=Clng(ChannelID)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目树形导航</title>
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<%call ShowClass_Tree()%>
</body>
</html>
<%
sub ShowClass_Tree()
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim rsClass,sqlClass,tmpDepth,i
	sqlClass="select C.ClassID,C.ClassName,C.Depth,L.LayoutFileName,C.NextID,C.LinkUrl,C.Child"
	if ChannelID=2 then
		sqlClass= sqlClass & " From ArticleClass C"
	elseif ChannelID=3 then
		sqlClass= sqlClass & " From SoftClass C"
	elseif ChannelID=4 then
		sqlClass= sqlClass & " From PhotoClass C"
	end if
	sqlClass= sqlClass & " inner join Layout L on C.LayoutID=L.LayoutID order by C.RootID,C.OrderID"
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
	if rsClass.bof and rsClass.bof then
		strClassTree="没有任何栏目"
	else
		strClassTree=""
		do while not rsClass.eof
			tmpDepth=rsClass(2)
			if rsClass(4)>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
			if tmpDepth>0 then
				for i=1 to tmpDepth
					if i=tmpDepth then
						if rsClass(4)>0 then
							strClassTree=strClassTree & "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>"
						else
							strClassTree=strClassTree & "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>"
						end if
					else
						if arrShowLine(i)=True then
							strClassTree=strClassTree & "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>"
						else
							strClassTree=strClassTree & "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>"
						end if
					end if
				next
			end if
			if rsClass(6)>0 then 
				strClassTree=strClassTree & "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
			else 
				strClassTree=strClassTree & "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
			end if 
			if rsClass(5)="" then
				strClassTree=strClassTree & "<a href='" & rsClass(3) & "?ClassID=" & rsClass(0) & "' target='_top'>"
			else
				strClassTree=strClassTree & "<a href='" & rsClass(5) & "' target='_blank'>"
			end if
			if rsClass(2)=0 then 
				strClassTree=strClassTree & "<b>"  & rsClass(1) & "</b>"
			else
				strClassTree=strClassTree & rsClass(1)
			end if 
			'if rsClass(5)<>"" then
			'	strClassTree=strClassTree & "(外)"
			'end if
			strClassTree=strClassTree & "</a>"
			if rsClass(6)>0 then 
				strClassTree=strClassTree & "（" & rsClass(6) & "）" 
			end if 
			strClassTree=strClassTree & "<br>"
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
	response.write strClassTree
end sub
%>
