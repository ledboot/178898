<!--#include file="../inc/conn.asp"-->
<%
if session("AdminName") = "" then
    response.Redirect "logout.asp"
end if
%>
<%
lx = Request("lx")
set rs=server.createobject("adodb.recordset")   
sql="select * from link where id="&request("id")
rs.open sql,conn,1,3   

rs("linkname")=trim(request("linkname"))
rs("linkwww")=trim(request("linkwww"))
If lx = 0 Then
rs("classname")=trim(request("indeximg"))
End If
rs("lx") = lx

rs.update
rs.close
set rs=nothing
%>
<script language=javascript>
alert("修改链接成功！");
location.href='link_index.asp?lx=<%=lx%>'
</script>