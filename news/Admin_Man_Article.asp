<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	classid = Request("classid")
	flag = Request("flag")
	zt = Request("zt")
	Articleid = Request("Articleid")
	pageno = Request("pageno")
	ArticleID_1 = Request("ArticleID_1")
	AI1Array = Split(ArticleID_1, ",")
%>
<%
	If flag = "del" Then
		If zt = 1 Then
			For i = 0 To UBound(AI1Array)
				Set rs = rsArticle4(AI1Array(i))
				rs("Deleted") = 1
				rs.Update
				rs.Close
				Set rs = Nothing
			Next
			Response.Redirect("Admin_Man_Article.asp?classid=" & classid)
			Response.End()
		End If
		If zt = 2 Then
			Set rs = rsArticle4(Articleid)
			rs("Deleted") = 1
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Redirect("Admin_Man_Article.asp?classid=" & classid)
			Response.End()
		End If
		If zt = 4 Then
			Set rs = rsArticle4(Articleid)
			If rs("OnTop") = 0 Then
				rs("OnTop") = 1
			Else
				rs("OnTop") = 0
			End If
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Redirect("Admin_Man_Article.asp?classid=" & classid)
			Response.End()
		End If
		If zt = 6 Then
			Set rs = rsArticle4(Articleid)
			If rs("Hot") = 0 Then
				rs("Hot") = 1
			Else
				rs("Hot") = 0
			End If
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Redirect("Admin_Man_Article.asp?classid=" & classid)
			Response.End()
		End If
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>������</title>
		<script language="javascript">
			function SelectAllItemNS(status, chkname){
				var a = document.forms[1].elements;
				var i;
				for(i=0;i<a.length;i++) {
					if(a[i].name == chkname) a[i].checked = status;
				}
			}
			function SelectAllItem(status, chkname){
				if(navigator.appName=="Netscape") SelectAllItemNS(status, chkname);
				else SelectAllItemIE(status, chkname);
			}
			function SelectAllItemIE(status, chkname){
				var a = document.all.item(chkname);
				var i;
				if(a!=null) {
					if(a.length!=null) {
						for (i=0; i<a.length; i++) {
							a(i).checked = status;
						}
					}
					else{
						a.checked = status;
					}
				}
			}
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
		<center>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<%
					Set rs = rsClass7(classid)
					classname = rs("classname")
					rs.Close
					Set rs = Nothing
				%>
				<tr> 
					<td align="center" height="25" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;"><%=classname%> �� ��Ϣ����</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						<b>&nbsp;��������</b>
						<a href="Admin_Add_Article.asp?classid=<%=classid%>">��Ӹ���Ŀ�µ���Ϣ</a>
						 |&nbsp;<a href="Admin_Man_Article.asp?classid=<%=classid%>">�������Ŀ�µ���Ϣ</a>
						  | <a href="Admin_HszM_Article.asp?classid=<%=classid%>">��Ϣ����վ����</a>
						   | <a href="javascript:history.back()">����</a>
					</td>
				</tr>
			</table>
			<form method="POST" action="Admin_Man_Article.asp?Flag=del&zt=1&pageno=<%=pageno%>&classid=<%=classid%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;��Ϣ������<font color="#FF0000"><%=classname%></font> ȫ����Ϣ</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td width="4%" background="images/admin_top_bg.jpg">&nbsp;</td>
									<td width="6%" background="images/admin_top_bg.jpg" align="center">���</td>
									<td width="44%" background="images/admin_top_bg.jpg">&nbsp;��Ϣ����</td>
									<td width="8%" background="images/admin_top_bg.jpg" align="center">�������</td>
									<td width="10%" background="images/admin_top_bg.jpg" align="center">��Ϣ����</td>
									<td width="5%" background="images/admin_top_bg.jpg" align="center">�Ƿ�<br>���</td>
									<td width="23%" background="images/admin_top_bg.jpg" align="center">����ѡ��</td>
								</tr>
								<%
									Set rs = rsArticle5(classid)
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
									<td width="4%">&nbsp;<input type="checkbox" name="ArticleID_1" value="<%=rs("ArticleID")%>"></td>
									<td width="6%" align="center"><%=rs("ArticleID")%></td>
									<td width="44%">&nbsp;<%=rs("Title")%></td>
									<td width="8%" align="center"><%=rs("Hits")%></td>
									<td width="10%" align="center">
										<font color="#FF0000">
											<%
												If rs("OnTop") = 0 Then
													Response.Write("&nbsp;")
												Else
													Response.Write("�̶�")
												End If
												If rs("Hot") = 0 Then
													Response.Write("&nbsp;")
												Else
													Response.Write("&nbsp;�Ƽ�")
												End If
											%>
										</font>
									</td>
									<td width="5%" align="center">
										<font color=#FF0000>
											<%
												If rs("Passed") = 1 Then
													Response.Write("��")
												Else
													Response.Write("��")
												End If
											%>
										</font>
									</td>
									<td width="23%">
										&nbsp;<a href="Admin_Edit_Article.asp?Articleid=<%=rs("ArticleID")%>&classid=<%=classid%>">�޸�</a>
										 | <a href="Admin_Man_Article.asp?Flag=del&zt=2&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>">ɾ��</a>
										  | <a href="Admin_Man_Article.asp?Flag=del&zt=4&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>"><%If rs("OnTop") = 0 Then%>�̶�<%Else%>���<%End If%></a>
										   | <a href="Admin_Man_Article.asp?Flag=del&zt=6&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>"><%If rs("Hot") = 0 Then%>�Ƽ�<%Else%>���Ƽ�<%End If%></a> 
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
										<a href="Admin_Man_Article.asp?pageno=1&Classid=<%=classid%>">��ҳ</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Man_Article.asp?pageno=<%=pageno-1%>&Classid=<%=classid%>">��һҳ</a> 
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
										<a href="Admin_Man_Article.asp?pageno=<%=pageno+1%>&Classid=<%=classid%>">��һҳ</a> 
										<%
											else
										%>
										<font color="#C0C0C0">��һҳ</font>
										<%
											end if
										%>
										<a href="Admin_Man_Article.asp?pageno=<%=rs.pagecount%>&Classid=<%=classid%>">βҳ</a>
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
								<tr height="25"> 
									<td colspan="7">
										&nbsp;<input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'ArticleID_1'); else SelectAllItem(false, 'ArticleID_1');" type="checkbox" value="1" name="select_all" > ȫѡ 
										<input type="button" value="ɾ��ѡ������Ϣ" name="B1" onClick="javaScript:if (confirm('ȷ��Ҫɾ��ѡ�е���Ϣ�𡣱�������ѡ�е���Ϣ��ʱ��ŵ�����վ�У���Ҫʱ�����Դӻ���վ�лָ��� ')) form1.submit();" > 
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr> 
						<td>&nbsp;</td>
					</tr>
				</table>
			</form>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>