<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/function.asp"-->

<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
dim ComeUrl
ComeUrl=Request.ServerVariables("HTTP_REFERER")

if Action="SetSkin" then
	call SetSkin()
end if
if FoundErr=True then
	call WriteErrMsg()
end if

sub SetSkin()
	dim ClassID,SkinID
	ClassID=trim(request("ClassID"))
	SkinID=trim(request("SkinID"))
	if ClassID="" then
		ClassID=0
	else
		ClassID=Clng(ClassID)
	end if
	if SkinID="" then
		SkinID=0
	else
		SkinID=Clng(SkinID)
	end if
	response.Cookies("asp163")("SkinID")=SkinID
	response.Redirect ComeUrl
end sub
%>