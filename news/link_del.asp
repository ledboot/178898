<!--#include file="../inc/conn.asp"-->
<%
if session("AdminName") = "" then
    response.Redirect "logout.asp"
end if
lx = Request("lx")
%>
<%
set rs=server.createobject("adodb.recordset")   
sql="select * from link where id="&request("id")
rs.open sql,conn,1,3   
rs.delete
rs.close
set rs=nothing
response.redirect "link_index.asp?lx=" & lx
%>