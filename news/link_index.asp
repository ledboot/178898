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
if (confirm('��ȷ���Ƿ�Ҫɾ��������¼'))
location.href='link_del.asp?id='+id + '&lx=' + lx

}
//-->
</script>
<body bgcolor="#FFFFFF" text="#000000">
<center>
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"><a href="link_index.asp?lx=<%=lx%>">��������</a> <a href="link_add.asp?lx=<%=lx%>">�������</a></td>
    </tr>
    <tr> 
      <td>
	    <table width="100%" border="0" cellspacing="1" cellpadding="4" bgcolor="#FF6600">
          <tr> 
            <td width="33%" align="center" bgcolor="#FFFFFF">��������</td>
            <td width="36%" align="center" bgcolor="#FFFFFF">���ӵ�ַ</td>
            <td width="12%" align="center" bgcolor="#FFFFFF">���</td>
            <td width="19%" align="center" bgcolor="#FFFFFF">����</td>
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
							lxStr = "�������"
						Case 1
							lxStr = "��������"
					End Select
					Response.Write(lxStr)
				%>
			</td>
            <td bgcolor="#FFFFFF" width="19%" align="center"><a href="link_edit.asp?id=<%=rs("id")%>&lx=<%=lx%>">�޸�</a> 
              <a href="#"  onclick="return check(<%=rs("id")%>,<%=lx%>)">ɾ��</a></td>
          </tr>
          <%
                      rs.movenext
                      next
					  else
					  response.write ""
					 'response.write "��������������..."
                      end if
                      %>
          <tr> 
            <td colspan="4" align="center" bgcolor="#FFFFFF"><a href="link_index.asp?pageno=1&lx=<%=lx%>">��ҳ</a> 
              <%if pageno>1 then %> <a href="link_index.asp?pageno=<%=pageno-1%>&lx=<%=lx%>">��һҳ</a> 
              <%else%> <font color="#C0C0C0">��һҳ</font> <%end if %> <%if cint(rs.pagecount)>cint(pageno) then %> <a href="link_index.asp?pageno=<%=pageno+1%>&lx=<%=lx%>">��һҳ</a> 
              <%else%> <font color="#C0C0C0">��һҳ</font> <%end if %> <a href="link_index.asp?pageno=<%=rs.pagecount%>&lx=<%=lx%>">βҳ</a> 
              ҳ�Σ�<font color="#FF0000"><%=pageno%></font>/<font color="#FF0000"><%=rs.pagecount%></font> ҳ&nbsp;<font color="#FF0000"><%=rs.pagesize %></font> ��/ҳ �� <font color="#FF0000"><%=rs.recordcount%></font> �� </td>
          </tr>
        </table>
        
	  </td>
    </tr>
  </table>
</center>
</body>
</html>