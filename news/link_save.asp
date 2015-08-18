<!--#include file="../inc/conn.asp"-->
<%
if session("AdminName") = "" then
    response.Redirect "logout.asp"
end if
%>
<%

set rs=server.createobject("adodb.recordset")   
sql="select * from link"   
rs.open sql,conn,1,3   
rs.addnew

rs("linkname")=trim(request("linkname"))
rs("linkwww")=trim(request("linkwww"))
lx = Request("lx")
If lx = 0 Then
	rs("classname") = Trim(Request("indeximg"))
End If
rs("lx") = lx
rs.update
rs.close
set rs=nothing
%>
<script language=javascript>
alert("链接添加成功！");
location.href='link_index.asp?lx=<%=lx%>';
</script>