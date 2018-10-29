<%@language=vbscript codepage=936 %>
<%
session("AdminName")=""
Response.Cookies("asp163")("UserName")=""
Response.Cookies("asp163")("UserLevel")=""
Response.Redirect "index.asp"
%>