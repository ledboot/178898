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
	If flag = "man" And zt = "del" Then '删除
		conn.execute("delete from GW_Paper where id=" & id)
		Response.Redirect("Admin_Paper.asp?flag=man")
		Response.End()
	ElseIf flag = "edit" And zt = "save" Then '编辑
		Title = request.Form("Title")
		Profession = request.Form("Profession")
		TotalScore = request.Form("TotalScore")
		ExamTime = request.Form("ExamTime")
		Verific = request.Form("Verific")
		If strlen(Title) < 3 Or strlen(Title) > 30 Then
			Error1 "对不起，试卷名称长度要大于3小于30，请重新输入!", "Admin_Paper.asp?flag=man"
		End If
		if len(Profession) = 0 then
			Error1 "对不起，请选择专业!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(TotalScore) = false then
			Error1 "对不起，试卷总分数请输入数字!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(ExamTime) = false then
			Error1 "对不起，考试时间请输入数字!", "Admin_Paper.asp?flag=man"
		end if
		if len(Verific) = 0 then
			Error1 "对不起，请选择试卷状态!", "Admin_Paper.asp?flag=man"
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
	Elseif zt = "save" then '新建
		Title = SafeRequest("Title", 0)
		Profession = SafeRequest("Profession", 0)
		TotalScore = SafeRequest("TotalScore", 0)
		ExamTime = SafeRequest("ExamTime", 0)
		Verific = SafeRequest("Verific", 0)
		If strlen(Title) < 3 Or strlen(Title) > 30 Then
			Error1 "对不起，试卷名称长度要大于3小于30，请重新输入!", "Admin_Paper.asp?flag=man"
		End If
		if len(Profession) = 0 then
			Error1 "对不起，请选择专业!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(TotalScore) = false then
			Error1 "对不起，试卷总分数请输入数字!", "Admin_Paper.asp?flag=man"
		end if
		if IsNumeric(ExamTime) = false then
			Error1 "对不起，考试时间请输入数字!", "Admin_Paper.asp?flag=man"
		end if
		if len(Verific) = 0 then
			Error1 "对不起，请选择试卷状态!", "Admin_Paper.asp?flag=man"
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
		<title>内容区</title>
		<script language="vbscript">
			sub add()
					if formAdd.Title.value="" then
						msgbox "请输入试卷名称！",48,"注意！"
						formAdd.Title.focus()
						exit sub
					end if
					if len(formAdd.Title.value )<3 or len(formAdd.Title.value )>30 then
						msgbox "试卷名称长度要大于3小于30",48,"注意！"
						formAdd.Title.focus()
						exit sub
					end if
				
					if formAdd.TotalScore.value="" then
						msgbox "试卷总分不能为空！",48,"注意！"
						formAdd.TotalScore.focus()
						exit sub
					end if
					formAdd.submit()
				end sub
		</script>
        <script language="vbscript">
				sub edit()
					if formEdit.Title.value="" then
						msgbox "请输入试卷名称！",48,"注意！"
						formEdit.Title.focus()
						exit sub
					end if
					if len(formEdit.Title.value )<3 or len(formEdit.Title.value )>30 then
						msgbox "试卷名称长度要大于3小于30",48,"注意！"
						form2.StudnetName.focus()
						exit sub
					end if
				
					if formEdit.TotalScore.value="" then
						msgbox "试卷总分不能为空！",48,"注意！"
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
					<td align="center" height="25" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">试卷管理 — 信息管理</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						<b>&nbsp;管理导航：</b> <a href="javascript:history.back()">返回</a> <a href="Admin_Paper.asp?flag=new">新建试卷</a>
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
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;信息导航：<font color="#FF0000">试卷管理</font> 全部信息</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td   background="images/admin_top_bg.jpg" align="center">编号</td>
									<td width="30%"  background="images/admin_top_bg.jpg" align="center">试卷名称</td>
									<td   background="images/admin_top_bg.jpg" align="center">专业</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">更新日期</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">点击数</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">状态</td>
                                    <td   background="images/admin_top_bg.jpg" align="center">操作</td>
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
										试卷没有填写专业
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
													Response.Write("已审核")
												Else
													Response.Write("未审核")
												End If
											%>
										</font>
                                    </td>
									<td   align="center">
										<a href="Admin_Paper.asp?flag=edit&id=<%=rs("ID")%>">修改</a>
										 | <a href="Admin_Paper.asp?flag=man&zt=del&id=<%=rs("ID")%>&pageno=<%=pageno%>">删除</a>| <a href="Admin_Questions_List.asp?flag=man&papaerId=<%=rs("ID")%>">查看试题</a>| <a href="Admin_Questions_Add.asp?flag=add&papaerId=<%=rs("ID")%>&pageno=<%=pageno%>&tmtype=1">添加试题</a>
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
										<a href="Admin_Paper.asp?pageno=1">首页</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Paper.asp?pageno=<%=pageno-1%>">上一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">上一页</font>
										<%
											end if
										%>
										<%
											if cint(rs.pagecount)>cint(pageno) then
										%>
										<a href="Admin_Paper.asp?pageno=<%=pageno+1%>">下一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">下一页</font>
										<%
											end if
										%>
										<a href="Admin_Paper.asp?pageno=<%=rs.pagecount%>">尾页</a>
										&nbsp;<font class="font_navigation">页次：</font>
										<font color="#FF0000"><%=pageno%></font>
										<font class="font_navigation">/</font>
										<font color="#FF0000"><%=rs.pagecount%></font>
										<font class="font_navigation">页</font>
										&nbsp;<font color="#FF0000"><%=rs.pagesize %></font>
										<font class="font_navigation">条/页 共</font>
									  <font color="#FF0000"><%=rs.recordcount%></font>
										<font class="font_navigation">条</font>
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
				elseif flag ="edit" then '编辑
				set rs2 =  rsPaper2(id)
			%>
            <form action="Admin_Paper.asp?flag=edit&zt=save&id=<%=id%>" name ="formEdit" method="POST">
            	<table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#F7F7F7">
                <tr>
                    <th colspan="2" align="center">修改试卷</th>
                    </tr>
                	<tr>
                    	<th align="right">试卷名称：</th>
                        <th align="left"><input type="text" name="Title" value="<%=rs2("Title")%>" /></th>
                    </tr>
                    <tr>
                    	<th align="right">专业：</th>
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
                                            <b>请先添加专业</b>
                                            <%
                                         End If
                                        %></th>
                    </tr>
                    <tr>
                    	<th align="right">试卷总分：</th>
                        <th align="left"><input type="text" name="TotalScore" value="<%=rs2("TotalScore")%>" /></th>
                    </tr>
                    <tr>
                    	<th align="right">默认考试时间：</th>
                        <th align="left"><input type="text" name="ExamTime" value="<%=rs2("ExamTime")%>" /><font color="#FF0000"> 说明：以分钟计算(<strong>如120分钟) </strong></font></th>
                    </tr>
                     <tr>
                    	<th align="right">状态：</th>
                        <th align="left">
							
                            审核：<input type="radio" name="Verific" <%If rs2("Verific") = 1 then response.Write("checked")%> value="1"/>
							
                            未审核：<input type="radio" name="Verific" <%If rs2("Verific") = 0 then response.Write("checked")%> value="0"/>
							
                        </th>
                    </tr>
                    <tr>
                    	<th colspan="2">
                        	<input type="button" value="确认修改" name="B1" onClick="edit()">
                        </th>
                    </tr>
                </table>
            
            </form>
			

			<%
				rs2.close
				set rs2 = nothing
				else '新建
			%>
            	<form method="post" action="Admin_Paper.asp?zt=save" name ="formAdd">
            	<table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#F7F7F7">
                	<tr>
                    <th colspan="2" align="center">新建试卷</th>
                    </tr>
                	<tr>
                    	<th align="right">试卷名称：</th>
                        <th align="left"><input type="text" name="Title" value="" /></th>
                    </tr>
                    <tr>
                    	<th align="right">专业：</th>
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
                                            <b>请先添加专业</b>
                                            <%
                                         End If
                                        %>
                        </th>
                    </tr>
                    <tr>
                    	<th align="right">试卷总分：</th>
                        <th align="left"><input type="text" name="TotalScore" value="" /></th>
                    </tr>
                    <tr>
                    	<th align="right">默认考试时间：</th>
                        <th align="left"><input type="text" name="ExamTime" value="120" /><font color="#FF0000"> 说明：以分钟计算(<strong>如120分钟) </strong></font></th>
                    </tr>
                     <tr>
                    	<th align="right">状态：</th>
                        <th align="left">审核：<input type="radio" name="Verific" value="1"/>
                        	未审核：<input type="radio" name="Verific" value="0"/>
                        </th>
                    </tr>
                    <tr>
                    	<th colspan="2">
                        	<input type="button" value="添加" name="B1" onClick="add()">
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