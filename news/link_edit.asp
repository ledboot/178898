<!--#include file="../inc/conn.asp"-->
<%
if session("AdminName") = "" then
    response.Redirect "logout.asp"
end if
lx = Request("lx")
%>
<html>
<head>
<title><%=webtitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/index.CSS" rel="stylesheet" type="text/css">
</head>
<script language="vbscript">
sub chat()

if form1.linkname.value="" then
msgbox "�������������ƣ�",48,"ע�⣡"
form1.linkname.focus()
exit sub
end if

if form1.linkwww.value="" then
msgbox "���������ӵ�ַ��",48,"ע�⣡"
form1.linkwww.focus()
exit sub
end if

form1.submit()
end sub

-->
</script>
<%
set rs=server.createobject("adodb.recordset")   
sql="select * from link where id="&request("id")   
rs.open sql,conn,1,1  

%>

<body bgcolor="#FFFFFF" text="#000000">
<center>
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"><a href="link_index.asp">��������</a> <a href="link_add.asp">�������</a></td>
    </tr>
    <tr> 
      <td> 
        <table width="100%" border="0" cellspacing="0" cellpadding="4">
          <form name="form1" method="post" action="link_editsave.asp?id=<%=request("id")%>">
            <tr> 
              <td width="20%" align="right">�������ƣ�</td>
              <td width="80%"> 
                <input type="text" name="linkname" size="40" value="<%=rs("linkname")%>">
                <font color="#FF0000">*</font> </td>
            </tr>
            <tr> 
              <td width="20%" align="right">���ӵ�ַ��</td>
              <td width="80%"> 
                <input type="text" name="linkwww" value="<%=rs("linkwww")%>" size="40">
                <font color="#FF0000">*<br> ����http://www.0733168.com/</font></td>
            </tr>
			<%
				If lx = 0 Then
			%>
			<tr height="25"> 
				<td align="right">�ĵ��ϴ���</td>
				<td><iframe name="indeximg_1" src="photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=0&ThumbnailSize=0" width="94%" height=25 scrolling="no" marginWidth=1 frameSpacing=0 marginHeight=0 frameBorder=0></iframe></td>
			</tr>
			<tr height="25"> 
				<td align="right">�ϴ���ַ��</td>
				<td><input class="INPUT1" type="text" name="indeximg" size="45" value="<%=rs("classname")%>"></td>
			</tr>
			<%
				End If
			%>
            <tr> 
              <td width="20%" align="right">�������</td>
              <td width="80%"> <input type="text" value="<%If lx = 0 Then%>�������<%ElseIf lx = 1 Then%>��������<%End If%>" readonly><input type="hidden" value="<%=lx%>" name="lx" id="lx"></td>
            </tr>
            <tr> 
              <td align="right" width="20%">&nbsp;</td>
              <td align="center" width="80%"> 
                <input type="button" name="Button" value="��  ��" onClick="chat()">
                <input type="reset" name="Submit2" value="��  ��">
              </td>
            </tr>
          </form>
        </table>
      </td>
    </tr>
  </table>
</center>
</body>
</html>