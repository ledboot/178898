<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	flag = SafeRequest("flag", 0)
	id = SafeRequest("id", 1)
	zt = SafeRequest("zt", 0)
%>
<%
	If flag = "man" And zt = "del" Then
		conn.execute("delete from GW_Profession where id=" & id)
		Response.Redirect("Admin_Profession.asp?flag=man")
		Response.End()
	ElseIf flag = "edit" And zt = "save" Then
		ProfessionName = SafeRequest("ProfessionName", 0)
		If strlen(ProfessionName) < 3 Or strlen(ProfessionName) > 27 Then
			Error1 "�Բ���רҵ���Ʋ��Ϸ�������������!", "Admin_Profession.asp?flag=man"
		End If
		set rs = rsProfesion3(id)
		rs("ProfessionName") = ProfessionName
		rs.update
		rs.close
		set rs = nothing
		Response.Redirect("Admin_Profession.asp?flag=man")
		Response.End()
	Elseif zt = "save" then
		ProfessionName = SafeRequest("ProfessionName", 0)
		Set rs = rsProfesion()
		rs.AddNew
		rs("ProfessionName") = ProfessionName
		rs.Update
		rs.Close
		Set rs = Nothing 
	end if
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
		<script language="vbscript">
			sub chat()
				if form2.ProfessionName.value = "" then
					msgbox "������רҵ����",48,"ע�⣡"
					form2.ProfessionName.focus()
					exit sub
				end if
				if len(form2.ProfessionName.value) <3 or len(form2.ProfessionName.value) >27 then
					msgbox "רҵ������Ҫ����3С��27��",48,"ע�⣡"
					form2.ProfessionName.focus()
					exit sub
				end if
				form2.submit()
			end sub
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
		<center>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td align="center" height="25" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">רҵ���� �� ��Ϣ����</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						<b>&nbsp;��������</b>
						<a href="Admin_Profession_add.asp">���רҵ</a>
						   | <a href="javascript:history.back()">����</a>
					</td>
				</tr>
			</table>
			<%
				if flag ="man" then
				Set rs = rsProfesion()
			%>
			<form method="POST" action="Admin_Profession.asp?Flag=man&pageno=<%=pageno%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;��Ϣ������<font color="#FF0000">רҵ����</font> ȫ����Ϣ</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td width="6%" background="images/admin_top_bg.jpg" align="center">���</td>
									<td width="44%" background="images/admin_top_bg.jpg" align="center">&nbsp;רҵ����</td>
									<td width="23%" background="images/admin_top_bg.jpg" align="center">����ѡ��</td>
								</tr>
								<%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=20
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
								%>
								<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
									<td width="6%" align="center"><%=rs("ID")%></td>
									<td width="44%" align="center">&nbsp;<%=rs("ProfessionName")%></td>
									<td width="23%" align="center">
										&nbsp;<a href="Admin_Profession.asp?flag=edit&id=<%=rs("ID")%>">�޸�</a>
										 | <a href="Admin_Profession.asp?flag=man&zt=del&id=<%=rs("ID")%>&pageno=<%=pageno%>">ɾ��</a>
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
										<a href="Admin_Profession.asp?pageno=1&Classid=<%=classid%>">��ҳ</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Profession.asp?pageno=<%=pageno-1%>&Classid=<%=classid%>">��һҳ</a> 
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
										<a href="Admin_Profession.asp?pageno=<%=pageno+1%>&Classid=<%=classid%>">��һҳ</a> 
										<%
											else
										%>
										<font color="#C0C0C0">��һҳ</font>
										<%
											end if
										%>
										<a href="Admin_Profession.asp?pageno=<%=rs.pagecount%>&Classid=<%=classid%>">βҳ</a>
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
					<tr> 
						<td>&nbsp;</td>
					</tr>
				</table>
			</form>
			<%
				Elseif flag = "edit" then
				set rs = rsProfesion3(id)
			%>
			<form method="POST" action="Admin_Profession.asp?flag=edit&zt=save&id=<%=id%>" name="form2">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td>&nbsp;</td>
					</tr>
					<tr> 
						<td height="25" style="border: 1 solid #C0C0C0" background="images/admin_top_bg.jpg">&nbsp;��ǰ״̬��רҵ����&nbsp;<font color="#FF0000">�����Ϣ</font></td>
					</tr>
					<tr> 
						<td align="center" bgcolor="#F7F7F7" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0"> 
							<form method="POST" action="?Flag=save" name="form1">
								<table border="0" cellpadding="0" cellspacing="0">
									<tr height="25"> 
										<td width="80">������Ŀ��</td>
										<td><strong><font color="#800000">רҵ����</font></strong></td>
									</tr>
									<tr height="25"> 
										<td>רҵ���ƣ�</td>
										<td ><input class="INPUT1" name="ProfessionName" type="text" style="width: 420;" maxlength="80"  value="<%=rs("ProfessionName")%>"> <font color="#FF0000">*</font></td>
									</tr>
									<tr height="25" align="center"> 
										<td colspan="2"><input type="button" value="�޸�" name="B1" onClick="chat()"></td>
									</tr>
								</table>
							</form>
						</td>
					</tr>
				</table>
			</form>
			<%
				rs.close
				set rs = nothing
				end if
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>