<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
	flag = SafeRequest("flag", 0)
	papaerId  = SafeRequest("papaerId", 0)
	quesId = SafeRequest("quesId",0)
	zt = SafeRequest("zt",0)
%>
<%
	if flag = "man" and zt = "del" then
		conn.execute("delete from GW_SubjectLibrary where id=" & quesId)
		Response.Redirect("Admin_Questions_List.asp?flag=man&papaerId="&papaerId)
		Response.End()
	end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>�����Ϣ</title>
	</head>
	<body>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr> 
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">�Ծ���� �� ��Ϣ����</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;��������</b><a href="Admin_Paper.asp?flag=man">�������Ŀ�µ���Ϣ</a> | <a href="Admin_Paper.asp?flag=man">����</a></td>
				</tr>
			</table>
		<form method="POST" action="Admin_Questions_List.asp?Flag=man&pageno=<%=pageno%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;��Ϣ������<font color="#FF0000">�������</font> ȫ����Ϣ</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td width="4%" background="images/admin_top_bg.jpg" align="center">���</td>
									<td width="26%" background="images/admin_top_bg.jpg" align="center">�Ծ�����</td>
                                    <td width="34%" background="images/admin_top_bg.jpg" align="center">��Ŀ��</td>
                                    <td width="6%" background="images/admin_top_bg.jpg" align="center">��Ŀ����</td>
									<td width="23%" background="images/admin_top_bg.jpg" align="center">����ѡ��</td>
								</tr>
                                <%
								  Set rs = rsSubjectLibrary3(papaerId)
								%>
								<%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=20
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
								%>
								<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
                                    <td width="4%" align="center"><%=rs("GW_SubjectLibrary.ID")%></td>
									<td width="26%" align="center"><%=rs("GW_Paper.Title")%></td>
                                    <td width="34%" align="center"><%=Server.HtmlEncode(rs("GW_SubjectLibrary.Title"))%></td>
									<td width="6%" align="center">
										<%
											if rs("Types") = 1 then
												response.Write("��ѡ��")
											elseif rs("Types") = 2 then
												response.Write("��ѡ��(A-D)")
											elseif rs("Types") = 3 then
												response.Write("��ѡ��(A-E)")
											elseif rs("Types") = 4 then
												response.Write("��ѡ��(A-F)")
											elseif rs("Types") = 5 then
												response.Write("�ж���")
											end if
										%>
                                    </td>
									<td width="23%" align="center">
										&nbsp;<a href="Admin_Questions_Update.asp?flag=update&quesId=<%=rs("GW_SubjectLibrary.ID")%>&tmtype=<%=rs("Types")%>&papaerId=<%=papaerId%>">�޸�</a>
										 | <a href="Admin_Questions_List.asp?flag=man&zt=del&quesId=<%=rs("GW_SubjectLibrary.ID")%>&pageno=<%=pageno%>&papaerId=<%=papaerId%>">ɾ��</a>
									</td>
								</tr>
								<%
											rs.movenext
											If rs.EOF Then
												Exit For
											End If
					  					next
									End If
								%>
								<tr height="25" > 
									<td colspan="7" align="center">
										<a href="Admin_Questions_List.asp?pageno=1&papaerId=<%=papaerId%>">��ҳ</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Questions_List.asp?pageno=<%=pageno-1%>&papaerId=<%=papaerId%>">��һҳ</a> 
										<%
											else
										%>
										<font color="#C0C0C0">��һҳ</font>
										<%
											end if
										%>
										<%
											if cint(rs.pagecount)>cint(pageno) then
										%>
										<a href="Admin_Questions_List.asp?pageno=<%=pageno+1%>&papaerId=<%=papaerId%>">��һҳ</a> 
										<%
											else
										%>
										<font color="#C0C0C0">��һҳ</font>
										<%
											end if
										%>
										<a href="Admin_Questions_List.asp?pageno=<%=rs.pagecount%>&papaerId=<%=papaerId%>">βҳ</a>
										&nbsp;<font class="font_navigation">ҳ�Σ�</font>
										<font color="#FF0000"><%=pageno%></font>
										<font class="font_navigation">/</font>
										<font color="#FF0000"><%=rs.pagecount%></font>
										<font class="font_navigation">ҳ</font>
										&nbsp;<font color="#FF0000"><%=rs.pagesize %></font>
										<font class="font_navigation">��/ҳ ��</font>
									  <font color="#FF0000"><%=rs.recordcount%></font>
										<font class="font_navigation">��</font>
								  </td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</form>
            <%
				rs.close()
				set rs =nothing
			%>
	</body>
</html>
<%
	call CloseDatabase()
%>