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
		conn.execute("delete from GW_Student where ID=" & id)
		Response.Redirect("Admin_Add_Student.asp?flag=man")
		Response.End()
	ElseIf flag = "Edit" And zt = "save" Then
		StudnetName = SafeRequest("StudnetName", 0)
		IdCard = SafeRequest("IdCard", 0)
		ProfessionId = SafeRequest("Profession", 0)
		If strlen(StudnetName) < 2 Or strlen(StudnetName) > 27 Then
			Error1 "对不起，学员或身份证不合法，请重新输入!", "Admin_Add_Student.asp?flag=man"
		End If
		If strlen(IdCard) < 3 Or strlen(IdCard) > 19 Then
			Error1 "对不起，学员或身份证不合法，请重新输入!", "Admin_Add_Student.asp?flag=man"
		End If
		if len(ProfessionId) = 0 then
			Error1 "对不起，请选择学员专业!", "Admin_Add_Student.asp?flag=man"
		end if
		Set rs = rsStudnet2(id)
		rs("StudnetName") = StudnetName
		rs("IdCard") = IdCard
		rs("ProfessionId") = ProfessionId
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("Admin_Add_Student.asp?flag=man")
		Response.End()
	ElseIf zt = "save" Then
		StudnetName = SafeRequest("StudnetName", 0)
		IdCard = SafeRequest("IdCard", 0)
		ProfessionId = SafeRequest("Profession", 0)
		If strlen(StudnetName) < 2 Or strlen(StudnetName) > 27 Then
			Error1 "对不起，学员或身份证不合法，请重新输入!", "Admin_Add_Student.asp?flag=man"
		End If
		If strlen(IdCard) < 3 Or strlen(IdCard) > 19 Then
			Error1 "对不起，学员或身份证不合法，请重新输入!", "Admin_Add_Student.asp?flag=man"
		End If
		If ProfessionId ="" Then
			Error1 "对不起，请选择专业!", "Admin_Add_Student.asp?flag=man"
		End If
		set rs3 = rsStudnet3(IdCard)
		if rs3.bof  and rs3.eof then 
			Set rs = rsStudnet1()
			rs.AddNew
			rs("StudnetName") = StudnetName
			rs("IdCard") = IdCard
			rs("ProfessionId") = ProfessionId
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Redirect("Admin_Add_Student.asp?flag=man")
			Response.End()
		else
			Error1 "对不起，学员身份证已存在!", "Admin_Add_Student.asp?flag=man"
		end if
		rs3.close
		set rs3 = nothing
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>用户管理</title>
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
				if form1.StudnetName.value="" then
					msgbox "请输入学员名！",48,"注意！"
					form1.StudnetName.focus()
					exit sub
				end if
				if len(form1.StudnetName.value )<2 or len(form1.StudnetName.value )>27 then
					msgbox "学员名长度要大于3小于27",48,"注意！"
					form1.StudnetName.focus()
					exit sub
				end if
				if form1.IdCard.value="" then
					msgbox "身份证不能为空！",48,"注意！"
					form1.IdCard.focus()
					exit sub
				end if
				if len(form1.IdCard.value )<5 or len(form1.IdCard.value )>19 then
					msgbox "身份证长度要大于5小于19",48,"注意！"
					form1.IdCard.focus()
					exit sub
				end if
				form1.submit()
			end sub
		</script>
		<script language="vbscript">
				sub chat1()
					if form2.StudnetName.value="" then
						msgbox "请输入学员名！",48,"注意！"
						form2.StudnetName.focus()
						exit sub
					end if
					if len(form2.StudnetName.value )<3 or len(form2.StudnetName.value )>27 then
						msgbox "学员名长度要大于3小于27",48,"注意！"
						form2.StudnetName.focus()
						exit sub
					end if
				
					if form2.IdCard.value="" then
						msgbox "身份证不能为空！",48,"注意！"
						form2.IdCard.focus()
						exit sub
					end if
					if len(form2.IdCard.value )<5 or len(form2.IdCard.value )>19 then
						msgbox "身份证长度要大于5小于19",48,"注意！"
						form2.IdCard.focus()
						exit sub
					end if
					form2.submit()
				end sub
			
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
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">学员管理首页</td>
				</tr>
				<tr>
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						&nbsp;<b>&nbsp;管理导航：</b>
						&nbsp;<a href="Admin_Add_Student.asp?flag=man">管理学员首页</a>
						 | <a href="Admin_Add_Student.asp">添加新的学员</a>
						  | <a href="javascript:history.back()">返回</a>
					</td>
				</tr>
			</table>
			<%
				If flag = "man" Then
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr>
					<td height="30">当前位置：学员管理</td>
				</tr>
				<tr>
					<td bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0" valign="top"> 
						<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
							<tr height="25">
								<td width="19%" background="admin_top_bg.jpg">&nbsp;学员名称</td>
								<td width="15%" background="admin_top_bg.jpg">&nbsp;学员身份证</td>
                                <td width="27%" background="admin_top_bg.jpg">&nbsp;专业</td>
								<td width="15%" background="admin_top_bg.jpg">&nbsp;最后一次到访IP</td>
								<td width="18%" background="admin_top_bg.jpg">&nbsp;最后一次到访时间</td>
								<td width="21%" background="admin_top_bg.jpg">&nbsp;操作选项</td>
							</tr>
							<%	
							Set rs = rsStudnet1()
								While Not rs.BOF And Not rs.EOF
							%>
							<tr height="25"> 
								<td width="19%">&nbsp;<img src="images/tree_folder4.gif" valign="abvmiddle"><a href="Admin_Add_Student.asp?flag=Edit&id=<%=rs("id")%>"><%=rs("StudnetName")%></a></td>
								<td width="15%"><%=rs("IdCard")%></td>
                                 <td width="27%">
                                 <%
								 	dim proName
									proName = rs("ProfessionName")
								 	if isnull(proName) or len(proName) = 0 then
									%>
										暂无专业
									<%
                                    else%>
										<%=proName%>
									<%end if%>
                                 
                                 </td>
								<td width="15%">&nbsp;<%=rs("LastloginIP")%></td>
								<td width="18%">&nbsp;<%=rs("LastloginTime")%></td>
								<td width="21%">&nbsp;<a href="Admin_Add_Student.asp?flag=Edit&id=<%=rs("id")%>">修改</a> <a href="Admin_Add_Student.asp?flag=man&zt=del&id=<%=rs("id")%>">删除</a></td>
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
					Set rs2 = rsStudnet2(id)
			%>
			<form method="POST" action="Admin_Add_Student.asp?flag=Edit&zt=save&id=<%=id%>" name="form2">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr>
						<td>
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#E3E3E3" bordercolordark="#FFFFFF">
								<tr height="25">
									<td align="center" colspan="2" background="admin_top_bg.jpg"><b>编辑学员</b></td>
								</tr>
								<tr height="25">
									<td width="25%" align="right">&nbsp;学员名：</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="text" name="StudnetName" size="27" value="<%=rs2("StudnetName")%>"></td>
								</tr>
								<tr height="25">
									<td width="25%" align="right">&nbsp;身份证：</td>
								  <td width="75%">&nbsp;<input class="INPUT1" type="text" name="IdCard" size="27" value="<%=rs2("IdCard")%>"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" id="PurviewDetail" style="border: 1 solid #C0C0C0;">
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25">
									<td background="admin_top_bg.jpg" colspan="4" align="center"><b>选择专业</b></td>
								</tr>
								<tr height="25">
                                	<th style="word-break:break-all; word-wrap:break-word;">
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
                                            <b>请先添加专业</b>
                                            <%
                                         End If
                                        %>
                                    </th>
								</tr>
							</table>
						</td>
					</tr>
      				<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="center"><input type="button" value=" 确认修改" name="B1" onClick="chat1()"></td>
					</tr>
      				<tr>
						<td>&nbsp;</td>
					</tr>
				</table>
			</form>
			
			<%
					rs2.Close
					Set rs2 = Nothing
			%>
			<%
				Else
			%>
			<form method="POST" action="Admin_Add_Student.asp?zt=save" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr>
						<td>
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#E3E3E3" bordercolordark="#FFFFFF">
								<tr height="25">
									<td align="center" colspan="2" background="admin_top_bg.jpg"><b>添加学员</b></td>
								</tr>
								<tr height="25">
									<td width="25%" align="right">&nbsp;学员名：</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="text" name="StudnetName" size="27"></td>
								</tr>
								<tr height="25">
									<td width="25%" align="right">&nbsp;身份证：</td>
									<td width="75%">&nbsp;<input class="INPUT1" type="IdCard" name="IdCard" size="27"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td id="PurviewDetail" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25">
									<td background="admin_top_bg.jpg" colspan="4" align="center"><b>选择专业</b></td>
								</tr>
								<tr height="25">
                                	<th style="word-break:break-all; word-wrap:break-word;">
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
                                            <b>请先添加专业</b>
                                            <%
                                         End If
                                        %>
                                    </th>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="center"><input type="button" value=" 确认添加" name="B1" onClick="chat()"> </td>
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