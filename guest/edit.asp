<!--#include file="gsconn.asp"-->
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>
<%
dim sql
dim rs
 sql="select * from book where id="&request("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conngs,1,1
                %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="guestbook.css">
<title>�޸�����</title>
</head>

<body>

<form method="POST" action="saveedit.asp?id=<%=request("id")%>">
  <div align="center"><center>
      <table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td colspan="2" height="30"><font color="#006600">ǩ д �� ��</font></td>
        </tr>
        <tr> 
          <td width="15%" align="right"><font color="#0080FF">�������֣�</font></td>
          <td> 
            <input type="text" name="name" size="30" value="<%=rs("name")%>">
          </td>
        </tr>
        <tr> 
          <td align="right"><font color="#0080FF">���ĵ��ʣ�</font></td>
          <td> 
            <input type="text" name="email" size="30" value="<%=rs("email")%>">
          </td>
        </tr>
        <tr> 
          <td align="right"><font color="#0080FF">��ҳ��ַ��</font></td>
          <td> 
            <input type="text" name="homepage" size="30" value="<%=rs("homepage")%>">
          </td>
        </tr>
        <tr> 
          <td align="right"><font color="#0080FF">����OICQ��</font></td>
          <td> 
            <input type="text" name="oicq" size="30" value="<%=rs("oicq")%>">
          </td>
        </tr>
        <tr> 
          <td align="right"><font color="#0080FF">�������ԣ�</font></td>
          <td> 
            <textarea           cols="60" name="content" rows="8"><%content=replace(rs("content"),"<BR>",chr(13))
            content=replace(content,"&nbsp;"," ")
            response.write content%></textarea>
          </td>
        </tr>
      </table>
      </center></div><div align="center"><center>
      <p> 
        <input type="submit" value="ȷ���޸�" name="cmdok" >
        &nbsp;&nbsp; 
        <input type="reset" value="�ָ�ԭ��" name="cmdcancel" >
		&nbsp;&nbsp;
        <input type="button" value="�����޸�" name="cmdcancel2" onclick="javascript:history.back()">
      </p>
  </center></div>
</form>
</body>
</html>
<% 
 rs.close 
set rs=nothing 
  conn.close  
  set conn=nothing  
%>
 