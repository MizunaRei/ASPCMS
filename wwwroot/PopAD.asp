<!--#include file="Inc/conn.asp"-->
<html>
<head>
<title>µ¯³ö¹ã¸æ</title>
</head>
<body leftMargin="0" topMargin="0">
<%
dim ID,sqlAD,rsAD,AD
ID=Trim(request("ID"))
if ID="" then
	sqlAD="select * from Advertisement where  ADType=0 order by ID desc"
else
	sqlAD="select * from Advertisement where ADType=0 and ID="&Clng(ID)
end if
set rsAD=server.createobject("adodb.recordset")
rsAD.open sqlAD,conn,1,1
if not rsAd.bof and not rsAD.eof then
	if rsAD("isflash")=true then
		AD= "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
		if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
		if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
		AD = AD & "><param name='movie' value='" & rsAD("ImgUrl") & "'><param name='quality' value='high'><embed src='" & rsAD("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
		if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
		if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
		AD = AD & "></embed></object>"
	else
		AD ="<a href='" & rsAD("SiteUrl") & "' target='_blank' title='" & rsAD("SiteName") & "£º" & rsAD("SiteUrl") & "'><img src='" & rsAD("ImgUrl") & "'"
		if rsAD("ImgWidth")>0 then AD = AD & " width='" & rsAD("ImgWidth") & "'"
		if rsAD("ImgHeight")>0 then AD = AD & " height='" & rsAD("ImgHeight") & "'"
		AD = AD & " border='0'></a>"
	end if
	response.write AD
end if
rsAD.close
set rsAD=nothing
%>
</body>
</html>