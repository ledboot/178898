<!--#include file="gsconn.asp"-->
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<%
   dim sql 
   dim rs
   set rs=server.createobject("adodb.recordset")
   sql="delete from book where id="&request("id")
   rs.open sql,conngs,1,1
   rs.close
   set rs=nothing  
   conn.close
   set conn=nothing

   response.redirect "manager.asp"

%>
