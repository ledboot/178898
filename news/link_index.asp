<!--#include file="../inc/conn.asp"-->
<%
if session("AdminName") = "" then
    response.Redirect "logout.asp"
end if
%>
<%
lx = SafeRequest("lx", 1)
set rs=server.createobject("adodb.recordset")   
sql="select * from link where lx=" & lx & " order by id desc"
rs.open sql,conn,1,1
%>
<html>
<head>
<title><%=webtitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/index.CSS" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript">
<!--
function check(id,lx) { //v2.0
if (confirm('请确定是否要删除这条记录'))
location.href='link_del.asp?id='+id + '&lx=' + lx

}
//-->
</script>
<body bgcolor="#FFFFFF" text="#000000">
<center>
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"><a href="link_index.asp?lx=<%=lx%>">管理链接</a> <a href="link_add.asp?lx=<%=lx%>">添加链接</a></td>
    </tr>
    <tr> 
      <td>
	    <table width="100%" border="0" cellspacing="1" cellpadding="4" bgcolor="#FF6600">
          <tr> 
            <td width="33%" align="center" bgcolor="#FFFFFF">链接名称</td>
            <td width="36%" align="center" bgcolor="#FFFFFF">链接地址</td>
            <td width="12%" align="center" bgcolor="#FFFFFF">类别</td>
            <td width="19%" align="center" bgcolor="#FFFFFF">操作</td>
          </tr>
          <%
	pageno=request.querystring("pageno")
	if pageno="" then pageno=1
    if pageno=0 then pageno=1 		
			if not rs.bof and not rs.eof then
		rs.pagesize=20
		rs.absolutepage=pageno
    end if 
		  If not rs.bof and not rs.eof Then                      
			    For i=1 to rs.pagesize                      
			    If rs.eof then exit for
		  %>
          <tr> 
            <td bgcolor="#FFFFFF" width="33%"><%=rs("linkname")%></td>
            <td bgcolor="#FFFFFF" width="36%"><a href="<%=rs("linkwww")%>" target="_blank"><%=rs("linkwww")%></a></td>
            <td bgcolor="#FFFFFF" width="12%" align="center">
				<%
					Select Case lx
						Case 0
							lxStr = "合作伙伴"
						Case 1
							lxStr = "友情链接"
					End Select
					Response.Write(lxStr)
				%>
			</td>
            <td bgcolor="#FFFFFF" width="19%" align="center"><a href="link_edit.asp?id=<%=rs("id")%>&lx=<%=lx%>">修改</a> 
              <a href="#"  onclick="return check(<%=rs("id")%>,<%=lx%>)">删除</a></td>
          </tr>
          <%
                      rs.movenext
                      next
					  else
					  response.write ""
					 'response.write "内容正在整理中..."
                      end if
                      %>
          <tr> 
            <td colspan="4" align="center" bgcolor="#FFFFFF"><a href="link_index.asp?pageno=1&lx=<%=lx%>">首页</a> 
              <%if pageno>1 then %> <a href="link_index.asp?pageno=<%=pageno-1%>&lx=<%=lx%>">上一页</a> 
              <%else%> <font color="#C0C0C0">上一页</font> <%end if %> <%if cint(rs.pagecount)>cint(pageno) then %> <a href="link_index.asp?pageno=<%=pageno+1%>&lx=<%=lx%>">下一页</a> 
              <%else%> <font color="#C0C0C0">下一页</font> <%end if %> <a href="link_index.asp?pageno=<%=rs.pagecount%>&lx=<%=lx%>">尾页</a> 
              页次：<font color="#FF0000"><%=pageno%></font>/<font color="#FF0000"><%=rs.pagecount%></font> 页&nbsp;<font color="#FF0000"><%=rs.pagesize %></font> 条/页 共 <font color="#FF0000"><%=rs.recordcount%></font> 条 </td>
          </tr>
        </table>
        
	  </td>
    </tr>
  </table>
</center>
</body>
</html>