<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"-->
<!--#include file="../inc/md5.asp" -->
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
	flag = SafeRequest("flag", 0)
	id = SafeRequest("id", 1)
	zt = SafeRequest("zt", 0)
	If flag = "man" And zt = "del" Then
		conn.execute("delete from admin where id=" & id)
		Response.Redirect("Admin_Add_User.asp?flag=man")
		Response.End()
	ElseIf flag = "Edit" And zt = "save" Then
		Username = SafeRequest("Username", 0)
		password = SafeRequest("password", 0)
		Purview = SafeRequest("Purview", 1)
		If strlen(username) < 3 Or strlen(username) > 27 Then
			Error1 "�Բ����û��������벻�Ϸ�������������!", "Admin_Add_User.asp?flag=man"
		End If
		If strlen(password) < 3 Or strlen(password) > 27 Then
			Error1 "�Բ����û��������벻�Ϸ�������������!", "Admin_Add_User.asp?flag=man"
		End If
		If Purview = 0 Then
			Error1 "������������������!", "Admin_Add_User.asp?flag=man"
		End If
		If Purview = 2 Then
			AdminSc_Purview_ClassID = SafeRequest("AdminSc_Purview_ClassID", 0)
			AdminSh_Purview_ClassID = SafeRequest("AdminSh_Purview_ClassID", 0)
			If strlen(AdminSc_Purview_ClassID) < 3 Or strlen(AdminSc_Purview_ClassID) > 270 Then
				Error1 "������������������!", "Admin_Add_User.asp?flag=man"
			End If
			If strlen(AdminSh_Purview_ClassID) < 3 Or strlen(AdminSh_Purview_ClassID) > 270 Then
				Error1 "������������������!", "Admin_Add_User.asp?flag=man"
			End If
		Else
			AdminSc_Purview_ClassID = "all"
			AdminSh_Purview_ClassID = "all"
		End If
		Set rs = rsAdmin3(id)
		rs("Username") = Username
		rs("password") = password
		rs("Purview") = Purview
		rs("AdminSc_Purview_ClassID") = AdminSc_Purview_ClassID
		rs("AdminSh_Purview_ClassID") = AdminSh_Purview_ClassID
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("Admin_Add_User.asp?flag=man")
		Response.End()
	ElseIf zt = "save" Then
		Username = SafeRequest("Username", 0)
		password = SafeRequest("password", 0)
		Purview = SafeRequest("Purview", 1)
		If strlen(username) < 3 Or strlen(username) > 27 Then
			Error1 "�Բ����û��������벻�Ϸ�������������!", "Admin_Add_User.asp?flag=man"
		End If
		If strlen(password) < 3 Or strlen(password) > 27 Then
			Error1 "�Բ����û��������벻�Ϸ�������������!", "Admin_Add_User.asp?flag=man"
		End If
		If Purview = 0 Then
			Error1 "������������������!", "Admin_Add_User.asp?flag=man"
		End If
		If Purview = 2 Then
			AdminSc_Purview_ClassID = SafeRequest("AdminSc_Purview_ClassID", 0)
			AdminSh_Purview_ClassID = SafeRequest("AdminSh_Purview_ClassID", 0)
			If strlen(AdminSc_Purview_ClassID) < 3 Or strlen(AdminSc_Purview_ClassID) > 270 Then
				Error1 "������������������!", "Admin_Add_User.asp?flag=man"
			End If
			If strlen(AdminSh_Purview_ClassID) < 3 Or strlen(AdminSh_Purview_ClassID) > 270 Then
				Error1 "������������������!", "Admin_Add_User.asp?flag=man"
			End If
		Else
			AdminSc_Purview_ClassID = "all"
			AdminSh_Purview_ClassID = "all"
		End If
		Set rs = rsAdmin2()
		rs.AddNew
		rs("Username") = Username
		rs("password") = password
		rs("Purview") = Purview
		rs("AdminSc_Purview_ClassID") = AdminSc_Purview_ClassID
		rs("AdminSh_Purview_ClassID") = AdminSh_Purview_ClassID
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("Admin_Add_User.asp?flag=man")
		Response.End()
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>�û�����</title>
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
					else a.checked = status;
				}
			}
		</script>
		<script language="vbscript">
			sub chat()
				if form1.Username.value="" then
					msgbox "����������ע���û�����",48,"ע�⣡"
					form1.Username.focus()
					exit sub
				end if
				if len(form1.Username.value )<3 or len(form1.Username.value )>27 then
					msgbox "ע���û�������Ҫ����3С��27",48,"ע�⣡"
					form1.Username.focus()
					exit sub
				end if
				if form1.password.value="" then
					msgbox "���벻��Ϊ�գ�",48,"ע�⣡"
					form1.password.focus()
					exit sub
				end if
				if len(form1.password.value )<5 or len(form1.password.value )>27 then
					msgbox "���볤��Ҫ����5С��27",48,"ע�⣡"
					form1.password.focus()
					exit sub
				end if
				if form1.apassword.value<>form1.password.value then
					msgbox "ȷ�����벻��Ϊ�գ�����"&vbcrlf&"��������������벻һ����",48,"ע�⣡"
					form1.apassword.focus()
					exit sub
				end if
				form1.submit()
			end sub
		</script>
		<script language="vbscript">
				sub chat1()
					if form2.Username.value="" then
						msgbox "����������ע���û�����",48,"ע�⣡"
						form2.Username.focus()
						exit sub
					end if
					if len(form2.Username.value )<3 or len(form2.Username.value )>27 then
						msgbox "ע���û�������Ҫ����3С��27",48,"ע�⣡"
						form2.Username.focus()
						exit sub
					end if
				
					if form2.password.value="" then
						msgbox "���벻��Ϊ�գ�",48,"ע�⣡"
						form2.password.focus()
						exit sub
					end if
					if len(form2.password.value )<5 or len(form2.password.value )>27 then
						msgbox "���볤��Ҫ����5С��27",48,"ע�⣡"
						form2.password.focus()
						exit sub
					end if
					if form2.apassword.value<>form2.password.value then
						msgbox "ȷ�����벻��Ϊ�գ�����"&vbcrlf&"��������������벻һ����",48,"ע�⣡"
						form2.apassword.focus()
						exit sub
					end if
					form2.submit()
				end sub
			-->
		</script>
		<script language="javascript" type="text/javascript">
			function GetRadioValue(RadioName){
				var obj;    
				obj=document.getElementsByName(RadioName);
				if(obj!=null){
					var i;
					for(i=0;i<obj.length;i++){
						if(obj[i].checked){
							return obj[i].value;            
						}
					}
				}
				return null;
			}
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
		<center>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">�û�������ҳ</td>
				</tr>
				<tr>
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						&nbsp;<b>&nbsp;��������</b>
						&nbsp;<a href="Admin_Add_User.asp?flag=man">�����û���ҳ</a>
						 | <a href="Admin_Add_User.asp">����µĲ���Ա</a>
						  | <a href="javascript:history.back()">����</a>
					</td>
				</tr>
			</table>
			<%
				If flag = "man" Then
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr>
					<td height="30">��ǰλ�ã��û�����</td>
				</tr>
				<tr>
					<td bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0" valign="top"> 
						<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
							<tr height="25">
								<td width="19%" background="admin_top_bg.jpg">&nbsp;�û�����</td>
								<td width="27%" background="admin_top_bg.jpg">&nbsp;�û�Ȩ��</td>
								<td width="15%" background="admin_top_bg.jpg">&nbsp;���һ�ε���IP</td>
								<td width="18%" background="admin_top_bg.jpg">&nbsp;���һ�ε���ʱ��</td>
								<td width="21%" background="admin_top_bg.jpg">&nbsp;����ѡ��</td>
							</tr>
							<%
								Set rs = rsAdmin2()
								While Not rs.BOF And Not rs.EOF
							%>
							<tr height="25"> 
								<td width="19%">&nbsp;<img src="images/tree_folder4.gif" valign="abvmiddle"><a href="Admin_Add_User.asp?flag=Edit&id=<%=rs("id")%>"><%=rs("Username")%></a></td>
								<td width="27%">&nbsp;<%If rs("Purview") = 1 Then%>��������Ա<%Else%>��ͨ����Ա<%End If%> </td>
								<td width="15%">&nbsp;<%=rs("LastloginIP")%></td>
								<td width="18%">&nbsp;<%=rs("LastloginTime")%></td>
								<td width="21%">&nbsp;<a href="Admin_Add_User.asp?flag=Edit&id=<%=rs("id")%>">�޸�</a> <a href="Admin_Add_User.asp?flag=man&zt=del&id=<%=rs("id")%>">ɾ��</a></td>
							</tr>
							<%
									rs.MoveNext
								Wend
								rs.Close
								Set rs = Nothing
							%>
						</table>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>
			<%
				ElseIf flag = "Edit" Then
					Set rs3 = rsAdmin3(id)
			%>
			<form method="POST" action="Admin_Add_User.asp?flag=Edit&zt=save&id=<%=id%>" name="form2">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr>
						<td>
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#E3E3E3" bordercolordark="#FFFFFF">
								<tr height="25">
									<td align="center" colspan="2" background="admin_top_bg.jpg"><b>�༭����Ա</b></td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;�� �� ����</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="text" name="Username" size="27" value="<%=rs3("Username")%>"></td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;��&nbsp;&nbsp;&nbsp; �룺</td>
								  <td width="75%">&nbsp;<input class="INPUT1" type="password" name="Password" size="27" value="<%=rs3("Password")%>">
&nbsp;5-16λ����</td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;ȷ�����룺</td>
								  <td width="75%">&nbsp;<input class="INPUT1" type="password" name="aPassword" size="27" value="<%=rs3("Password")%>">
&nbsp;5-16λ����</td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;Ȩ�����ã�</td>
									<td width="75%">
										<input type="radio" value="1" name="Purview" onClick="PurviewDetail.style.display='none'" <%If rs3("Purview") = 1 Then%>checked<%End If%>>
										<font color="#800000">��������Ա</font>�����Զ����е���Ŀ���й���<br>
										<input type="radio" value="2" name="Purview" onClick="PurviewDetail.style.display=''" <%If rs3("Purview") = 2 Then%>checked<%End If%>>
										<font color="#800000">��ͨ����Ա</font>��ֻ�Բ�����Ŀ���й���
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" id="PurviewDetail" style="border: 1 solid #C0C0C0;">
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25">
									<td width="100%" align="center" background="admin_top_bg.jpg" colspan="4"><b>ѡ��ɲ�����Ȩ��</b></td>
								</tr>
								<tr height="25">
									<td width="30%">&nbsp;��Ŀ����</td>
									<td width="30%" align="center">��Ŀ���</td>
									<td width="20%" align="center">�Ƿ����¼��</td>
									<td width="20%" align="center">�Ƿ�������</td>
								</tr>
								<%
									AdminScArray = Split(rs3("AdminSc_Purview_ClassID"), ",")
									AdminShArray = Split(rs3("AdminSh_Purview_ClassID"), ",")
									Set rs = rsClass1(0)
									While Not rs.EOF And Not rs.BOF
								%>
								<tr height="25">
									<td width="30%">&nbsp;<img src="images/tree_folder4.gif" valign="abvmiddle"><%=rs("ClassName")%></td>
									<td width="30%">&nbsp;
										<%
											If rs("Class_num") = 0 Then
												If rs("ClassNameD") = 1 Then
													Response.Write("�ⲿ")
												Else
													Response.Write("�ڲ�")
												End If
											End If
										%>
									</td>
            						<td width="20%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs("XcClassid")%>" 
										<%
											For i = 0 To UBound(AdminScArray)
												If CLng(AdminScArray(i)) = CLng(rs("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td>
									<td width="20%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs("XcClassid")%>"
										<%
											For i = 0 To UBound(AdminShArray)
												If CLng(AdminShArray(i)) = CLng(rs("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td></tr>
								<%
										If CInt(rs("Class_num")) <> 0 Then
											Set rs1 = rsClass1(rs("XcClassid"))
											While Not rs1.EOF And Not rs1.BOF
								%>
								<tr height="25">
									<td width="30%">&nbsp; <font color="#FF0000">|--</font><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"><%=rs1("ClassName")%></td>
            						<td width="30%" align="center">
										<%
											If rs1("Class_num") = 0 Then
												If rs1("ClassNameD") = 1 Then
													Response.Write("�ⲿ")
												Else
													Response.Write("�ڲ�")
												End If
											End If
										%>
									</td>
									<td width="20%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs1("XcClassid")%>"
										<%
											For i = 0 To UBound(AdminScArray)
												If CLng(AdminScArray(i)) = Clng(rs1("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td>
									<td width="20%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs1("XcClassid")%>"
										<%
											For i = 0 To UBound(AdminShArray)
												If CLng(AdminShArray(i)) = CLng(rs1("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td>
							  </tr>
								<%
												If rs1("Class_num") <> 0 Then
													Set rs2 = rsClass1(rs1("XcClassid"))
													While Not rs2.EOF And Not rs2.BOF
								%>
								<tr height="25"> 
									<td width="30%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">|--</font><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"><%=rs2("ClassName")%></td>
									<td width="30%" align="center">
										<%
											If rs2("ClassNameD") = 1 Then
												Response.Write("�ⲿ")
											Else
												Response.Write("�ڲ�")
											End If
										%>
									</td>
									<td width="20%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs2("XcClassid")%>"
										<%
											For i = 0 To UBound(AdminScArray)
												If CLng(AdminScArray(i)) = CLng(rs2("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td>
									<td width="20%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs2("XcClassid")%>"
										<%
											For i = 0 To UBound(AdminShArray)
												If CLng(AdminShArray(i)) = CLng(rs2("XcClassid")) Then
													Response.Write(" CHECKED")
													Exit For
												End If
											Next
										%>>
									</td>
								</tr>
								<%
														rs2.MoveNext
													Wend
													rs2.Close
													Set rs2 = Nothing
												End If
								%>
								<%
												rs1.MoveNext
											Wend
											rs1.Close
											Set rs1 = Nothing
										End If
								%>
								<%
										rs.MoveNext
									Wend
									rs.Close
									Set rs = Nothing
								%>
								<tr>
									<td colspan="4" height="25">
										&nbsp;<input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'AdminSc_Purview_ClassID'); else SelectAllItem(false, 'AdminSc_Purview_ClassID');" type="checkbox" value="1" name="select_all">¼��Ȩ��ȫѡ
										 | <input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'AdminSh_Purview_ClassID'); else SelectAllItem(false, 'AdminSh_Purview_ClassID');" type="checkbox" value="1" name="select_all0">���Ȩ��ȫѡ
									</td>
								</tr>
							</table>
						</td>
					</tr>
      				<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="center"><input type="button" value=" ȷ���޸�" name="B1" onClick="chat1()"> <input type="reset" value=" ���·���" name="B2"></td>
					</tr>
      				<tr>
						<td>&nbsp;</td>
					</tr>
				</table>
			</form>
			<script language="javascript" type="text/javascript">
				temp = GetRadioValue("Purview")
				if(temp == 2){
					document.getElementById("PurviewDetail").style.display = "";
				}
				else{
					document.getElementById("PurviewDetail").style.display = "none";
				}
			</script>
			<%
					rs3.Close
					Set rs3 = Nothing
			%>
			<%
				Else
			%>
			<form method="POST" action="Admin_Add_User.asp?zt=save" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr>
						<td>
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#E3E3E3" bordercolordark="#FFFFFF">
								<tr height="25">
									<td align="center" colspan="2" background="admin_top_bg.jpg"><b>��ӹ���Ա</b></td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;�� �� ����</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="text" name="Username" size="27"></td>
								</tr>
								<tr height="25">
									<td width="25%">&nbsp;��&nbsp;&nbsp;&nbsp; �룺</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="password" name="Password" size="27">
									&nbsp;5-16λ����</td>
								</tr>
								<tr height="25">

									<td width="25%">&nbsp;ȷ�����룺</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="password" name="aPassword" size="27">
									&nbsp;5-16λ����</td>
								</tr>
								<tr  height="25">
									<td width="25%">&nbsp;Ȩ�����ã�</td>
									<td width="75%">
										<input type="radio" value="1" name="Purview" checked onClick="PurviewDetail.style.display='none'">
										<font color="#800000">��������Ա</font>�����Զ����е���Ŀ���й���<br>
										<input type="radio" value="2" name="Purview" onClick="PurviewDetail.style.display=''">
										<font color="#800000">��ͨ����Ա</font>��ֻ�Բ�����Ŀ���й���<br>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td id="PurviewDetail" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;display:none"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25">
									<td background="admin_top_bg.jpg" colspan="4" align="center"><b>ѡ��ɲ�����Ȩ��</b></td>
								</tr>
								<tr height="25">
									<td width="40%">&nbsp;��Ŀ����</td>
									<td width="12%" align="center">��Ŀ���</td>
									<td width="23%" align="center">�Ƿ����¼��</td>
									<td width="25%" align="center">�Ƿ�������</td>
								</tr>
								<%
									Set rs = rsClass1(0)
									While Not rs.EOF And Not rs.BOF
								%>
								<tr height="25">
									<td width="40%">&nbsp;<img src="images/tree_folder4.gif" valign="abvmiddle"><%=rs("ClassName")%></td>
									<td width="12%">&nbsp;
										<%
											If rs("Class_num") = 0 Then
												If rs("ClassNameD") = 1 Then
													Response.Write("�ⲿ")
												Else
													Response.Write("�ڲ�")
												End If
											End If
										%>
									</td>
									<td width="23%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs("XcClassid")%>"></td>
									<td width="25%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs("XcClassid")%>"></td>
								</tr>
								<%
										If CInt(rs("Class_num")) <> 0 Then
											Set rs1 = rsClass1(rs("XcClassid"))
											While Not rs1.EOF And Not rs1.BOF
								%>
								<tr height="25">
									<td width="40%">&nbsp; <font color="#FF0000">|--</font><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"><%=rs1("ClassName")%></td>
									<td width="12%" align="center">
										<%
											If rs1("Class_num") = 0 Then
												If rs1("ClassNameD") = 1 Then
													Response.Write("�ⲿ")
												Else
													Response.Write("�ڲ�")
												End If
											End If
										%>
									</td>
									<td width="23%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs1("XcClassid")%>"></td>
									<td width="25%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs1("XcClassid")%>"></td>
								</tr>
								<%
												If rs1("Class_num") <> 0 Then
													Set rs2 = rsClass1(rs1("XcClassid"))
													While Not rs2.EOF And Not rs2.BOF
								%>
								<tr height="25">
									<td width="40%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">|--</font><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"><%=rs2("ClassName")%></td>
									<td width="12%" align="center">
										<%
											If rs1("ClassNameD") = 1 Then
												Response.Write("�ⲿ")
											Else
												Response.Write("�ڲ�")
											End If
										%>
									</td>
									<td width="23%" align="center"><input type="checkbox" name="AdminSc_Purview_ClassID" value="<%=rs2("XcClassid")%>"></td>
									<td width="25%" align="center"><input type="checkbox" name="AdminSh_Purview_ClassID" value="<%=rs2("XcClassid")%>"></td>
								</tr>
								<%
														rs2.MoveNext
													Wend
													rs2.Close
													Set rs2 = Nothing
												End If
								%>
								<%
												rs1.MoveNext
											Wend
											rs1.Close
											Set rs1 = Nothing
										End If
								%>
								<%
										rs.MoveNext
									Wend
									rs.Close
									Set rs = Nothing
								%>
								<tr>
									<td colspan="4" height="24">
										&nbsp;<input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'AdminSc_Purview_ClassID'); else SelectAllItem(false, 'AdminSc_Purview_ClassID');" type="checkbox" value="1" name="select_all">¼��Ȩ��ȫѡ
										 | <input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'AdminSh_Purview_ClassID'); else SelectAllItem(false, 'AdminSh_Purview_ClassID');" type="checkbox" value="1" name="select_all0">���Ȩ��ȫѡ
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="center"><input type="button" value=" ȷ�����" name="B1" onClick="chat()"> <input type="reset" value=" ���·���" name="B2"></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
				</table>
			</form>
			<%
				End If
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>