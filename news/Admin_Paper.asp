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
	If flag = "man" And zt = "del" Then 'ɾ��
		conn.execute("delete from GW_Paper where id=" & id)
		Response.Redirect("Admin_Paper.asp?flag=man")
		Response.End()
	ElseIf flag = "edit" And zt = "save" Then '�༭
		Title = request.Form("Title")
		Profession = request.Form("Profession")
		TotalScore = request.Form("TotalScore")
		ExamTime = request.Form("ExamTime")
		Verific = request.Form("Verific")
		If strlen(Title) < 3 Or strlen(Title) > 30 Then
			Error1 "�Բ����Ծ����Ƴ���Ҫ����3С��30������������!", "Admin_Paper.asp?flag=man"
		End If
		if len(Profession) = 0 then
			Error1 "�Բ�����ѡ��רҵ!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(TotalScore) = false then
			Error1 "�Բ����Ծ��ܷ�������������!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(ExamTime) = false then
			Error1 "�Բ��𣬿���ʱ������������!", "Admin_Paper.asp?flag=man"
		end if
		if len(Verific) = 0 then
			Error1 "�Բ�����ѡ���Ծ�״̬!", "Admin_Paper.asp?flag=man"
		end if
		set rs2 = rsPaper2(id)
		rs2("Title") = Title
		rs2("ProfessionId") = Profession
		rs2("TotalScore") = TotalScore
		rs2("User") = session("AdminName")
		rs2("Verific") = cint(Verific)
		rs2("ExamTime") = ExamTime
		rs2("Date") = now()
		rs2.update
		rs2.close
		set rs2 = nothing
		Response.Redirect("Admin_Paper.asp?flag=man")
		Response.End()
	Elseif zt = "save" then '�½�
		Title = SafeRequest("Title", 0)
		Profession = SafeRequest("Profession", 0)
		TotalScore = SafeRequest("TotalScore", 0)
		ExamTime = SafeRequest("ExamTime", 0)
		Verific = SafeRequest("Verific", 0)
		If strlen(Title) < 3 Or strlen(Title) > 30 Then
			Error1 "�Բ����Ծ����Ƴ���Ҫ����3С��30������������!", "Admin_Paper.asp?flag=man"
		End If
		if len(Profession) = 0 then
			Error1 "�Բ�����ѡ��רҵ!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(TotalScore) = false then
			Error1 "�Բ����Ծ��ܷ�������������!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(ExamTime) = false then
			Error1 "�Բ��𣬿���ʱ������������!", "Admin_Paper.asp?flag=man"
		end if
		if len(Verific) = 0 then
			Error1 "�Բ�����ѡ���Ծ�״̬!", "Admin_Paper.asp?flag=man"
		end if
		
		Set rs = rsPaper()
		rs.AddNew
		rs("Title") = Title
		rs("ProfessionId") = Profession
		rs("TotalScore") = TotalScore
		rs("User") = session("AdminName")
		rs("Verific") = Verific
		rs("ExamTime") = ExamTime
		rs("Date") = now()
		rs("ProfessionName") = ProfessionName
		rs.Update
		rs.Close
		Set rs = Nothing 
		Response.Redirect("Admin_Paper.asp?flag=man")
		Response.End()
	end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>������</title>
		<script language="vbscript">
			sub add()
					if formAdd.Title.value="" then
						msgbox "�������Ծ����ƣ�",48,"ע�⣡"
						formAdd.Title.focus()
						exit sub
					end if
					if len(formAdd.Title.value )<3 or len(formAdd.Title.value )>30 then
						msgbox "�Ծ����Ƴ���Ҫ����3С��30",48,"ע�⣡"
						formAdd.Title.focus()
						exit sub
					end if
				
					if formAdd.TotalScore.value="" then
						msgbox "�Ծ��ֲܷ���Ϊ�գ�",48,"ע�⣡"
						formAdd.TotalScore.focus()
						exit sub
					end if
					formAdd.submit()
				end sub
		</script>
        <script language="vbscript">
				sub edit()
					if formEdit.Title.value="" then
						msgbox "�������Ծ����ƣ�",48,"ע�⣡"
						formEdit.Title.focus()
						exit sub
					end if
					if len(formEdit.Title.value )<3 or len(formEdit.Title.value )>30 then
						msgbox "�Ծ����Ƴ���Ҫ����3С��30",48,"ע�⣡"
						form2.StudnetName.focus()
						exit sub
					end if
				
					if formEdit.TotalScore.value="" then
						msgbox "�Ծ��ֲܷ���Ϊ�գ�",48,"ע�⣡"
						formEdit.TotalScore.focus()
						exit sub
					end if
					formEdit.submit()
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
					<td align="center" height="25" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">�Ծ���� �� ��Ϣ����</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						<b>&nbsp;��������</b> <a href="javascript:history.back()">����</a> <a href="Admin_Paper.asp?flag=new">�½��Ծ�</a>
					</td>
				</tr>
			</table>
			<%
				if flag ="man" then
				Set rs = rsPaper1()
			%>
			<form method="POST" action="Admin_Paper.asp?flag=man&pageno=<%=pageno%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;��Ϣ������<font color="#FF0000">�Ծ����</font> ȫ����Ϣ</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td   background="images/admin_top_bg.jpg" align="center">���</td>
									<td width="30%"  background="images/admin_top_bg.jpg" align="center">�Ծ�����</td>
									<td   background="images/admin_top_bg.jpg" align="center">רҵ</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">��������</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">�����</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">״̬</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">����</td>
								</tr>
								<%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=20
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
								%>
								<tr height="25"> 
									<td  align="center"><%=rs("ID")%></td>
									<td   align="center"><%=rs("Title")%></td>
                                    <td   align="center" >
									<%
								 	dim proName
									proName = rs("ProfessionName")
								 	if isnull(proName) or len(proName) = 0 then
									%>
                                    <font color=#FF0000>
										�Ծ�û����дרҵ
                                        </font>
									<%
                                    else%>
										<%=proName%>
									<%end if%>
                                    </td>
                                    <td   align="center"> <%=rs("Date")%></td>
                                    <td  align="center"> <%=rs("Hits")%></td>
                                    <td  align="center">
                                    	<font color=#FF0000>
											<%
												If rs("Verific") = 1 Then
													Response.Write("�����")
												Else
													Response.Write("δ���")
												End If
											%>
										</font>
                                    </td>
									<td   align="center">
										<a href="Admin_Paper.asp?flag=edit&id=<%=rs("ID")%>">�޸�</a>
										 | <a href="Admin_Paper.asp?flag=man&zt=del&id=<%=rs("ID")%>&pageno=<%=pageno%>">ɾ��</a>| <a href="Admin_Questions_List.asp?flag=man&papaerId=<%=rs("ID")%>">�鿴����</a>| <a href="Admin_Questions_Add.asp?flag=add&papaerId=<%=rs("ID")%>&pageno=<%=pageno%>&tmtype=1">�������</a>
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
									<td colspan="8" align="center">
										<a href="Admin_Paper.asp?pageno=1">��ҳ</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Paper.asp?pageno=<%=pageno-1%>">��һҳ</a> 
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
										<a href="Admin_Paper.asp?pageno=<%=pageno+1%>">��һҳ</a> 
										<%
											else
										%>
										<font color="#C0C0C0">��һҳ</font>
										<%
											end if
										%>
										<a href="Admin_Paper.asp?pageno=<%=rs.pagecount%>">βҳ</a>
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
				rs.close
				set rs = nothing
			%>
			<%
				elseif flag ="edit" then '�༭
				set rs2 =  rsPaper2(id)
			%>
            <form action="Admin_Paper.asp?flag=edit&zt=save&id=<%=id%>" name ="formEdit" method="POST">
            	<table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#F7F7F7">
                <tr>
                    <th colspan="2" align="center">�޸��Ծ�</th>
                    </tr>
                	<tr>
                    	<th align="right">�Ծ����ƣ�</th>
                        <th align="left"><input type="text" name="Title" value="<%=rs2("Title")%>" /></th>
                    </tr>
                    <tr>
                    	<th align="right">רҵ��</th>
                        <th align="left" style="word-break:break-all; word-wrap:break-word;" >
                        <%
										Set rsPro = rsProfesion()
										if not rsPro.bof and not rsPro.eof then
										if rsPro.size >0 then
											for i=1 to rsPro.recordcount
											%>
											<input type="radio" name="Profession"  value="<%=rsPro("ID")%>"  
											<%
											 	if clng(rsPro("ID")) = clng(rs2("ProfessionId")) then
													response.Write("checked")
												end if
											 %>/><%=rsPro("ProfessionName")%>
											<%
                                            rsPro.movenext
                                                    If rsPro.EOF Then
                                                        Exit For
                                                    End If
                                            next
											End if
											Else
											%>
                                            <b>�������רҵ</b>
                                            <%
                                         End If
                                        %></th>
                    </tr>
                    <tr>
                    	<th align="right">�Ծ��ܷ֣�</th>
                        <th align="left"><input type="text" name="TotalScore" value="<%=rs2("TotalScore")%>" /></th>
                    </tr>
                    <tr>
                    	<th align="right">Ĭ�Ͽ���ʱ�䣺</th>
                        <th align="left"><input type="text" name="ExamTime" value="<%=rs2("ExamTime")%>" /><font color="#FF0000"> ˵�����Է��Ӽ���(<strong>��120����) </strong></font></th>
                    </tr>
                     <tr>
                    	<th align="right">״̬��</th>
                        <th align="left">
							
                            ��ˣ�<input type="radio" name="Verific" <%If rs2("Verific") = 1 then response.Write("checked")%> value="1"/>
							
                            δ��ˣ�<input type="radio" name="Verific" <%If rs2("Verific") = 0 then response.Write("checked")%> value="0"/>
							
                        </th>
                    </tr>
                    <tr>
                    	<th colspan="2">
                        	<input type="button" value="ȷ���޸�" name="B1" onClick="edit()">
                        </th>
                    </tr>
                </table>
            
            </form>
			

			<%
				rs2.close
				set rs2 = nothing
				else '�½�
			%>
            	<form method="post" action="Admin_Paper.asp?zt=save" name ="formAdd">
            	<table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#F7F7F7">
                	<tr>
                    <th colspan="2" align="center">�½��Ծ�</th>
                    </tr>
                	<tr>
                    	<th align="right">�Ծ����ƣ�</th>
                        <th align="left"><input type="text" name="Title" value="" /></th>
                    </tr>
                    <tr>
                    	<th align="right">רҵ��</th>
                        <th align="left" style="word-break:break-all; word-wrap:break-word;" >
                        	<%
										Set rsPro = rsProfesion()
										if not rsPro.bof and not rsPro.eof then
										if rsPro.size >0 then
											for i=1 to rsPro.recordcount
											%>
											<input type="radio" name="Profession"  value="<%=rsPro("ID")%>" /><%=rsPro("ProfessionName")%>
											<%
                                            rsPro.movenext
                                                    If rsPro.EOF Then
                                                        Exit For
                                                    End If
                                            next
											End if
											Else
											%>
                                            <b>�������רҵ</b>
                                            <%
                                         End If
                                        %>
                        </th>
                    </tr>
                    <tr>
                    	<th align="right">�Ծ��ܷ֣�</th>
                        <th align="left"><input type="text" name="TotalScore" value="" /></th>
                    </tr>
                    <tr>
                    	<th align="right">Ĭ�Ͽ���ʱ�䣺</th>
                        <th align="left"><input type="text" name="ExamTime" value="120" /><font color="#FF0000"> ˵�����Է��Ӽ���(<strong>��120����) </strong></font></th>
                    </tr>
                     <tr>
                    	<th align="right">״̬��</th>
                        <th align="left">��ˣ�<input type="radio" name="Verific" value="1"/>
                        	δ��ˣ�<input type="radio" name="Verific" value="0"/>
                        </th>
                    </tr>
                    <tr>
                    	<th colspan="2">
                        	<input type="button" value="���" name="B1" onClick="add()">
                        </th>
                    </tr>
                </table>
            
            </form>
            
            <%
			end if
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>