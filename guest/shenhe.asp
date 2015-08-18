<!--#include file="gsconn.asp"--> 
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<%
if request("page") <> "" then
	str="?page="&request("page")
else
	str=""
end if
id=request("id")
set rsput=server.createobject("adodb.recordset")
if request("action")="kai" then
sqlput="UPDATE book SET shenhe=true WHERE id = " + id
else
sqlput="UPDATE book SET shenhe=false WHERE id = " + id
end if
rsput.open sqlput,conngs,1,3 
response.redirect "manager.asp"&str
rsput.close
set rsput=nothing
conngs.close
set conngs=nothing
%>