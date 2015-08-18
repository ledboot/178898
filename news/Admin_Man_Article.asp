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
		<title>内容区</title>
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
					<td align="center" height="25" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;"><%=classname%> — 信息管理</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;">
						<b>&nbsp;管理导航：</b>
						<a href="Admin_Add_Article.asp?classid=<%=classid%>">添加该栏目下的信息</a>
						 |&nbsp;<a href="Admin_Man_Article.asp?classid=<%=classid%>">管理该栏目下的信息</a>
						  | <a href="Admin_HszM_Article.asp?classid=<%=classid%>">信息回收站管理</a>
						   | <a href="javascript:history.back()">返回</a>
					</td>
				</tr>
			</table>
			<form method="POST" action="Admin_Man_Article.asp?Flag=del&zt=1&pageno=<%=pageno%>&classid=<%=classid%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;信息导航：<font color="#FF0000"><%=classname%></font> 全部信息</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td width="4%" background="images/admin_top_bg.jpg">&nbsp;</td>
									<td width="6%" background="images/admin_top_bg.jpg" align="center">编号</td>
									<td width="44%" background="images/admin_top_bg.jpg">&nbsp;信息标题</td>
									<td width="8%" background="images/admin_top_bg.jpg" align="center">点击次数</td>
									<td width="10%" background="images/admin_top_bg.jpg" align="center">信息属性</td>
									<td width="5%" background="images/admin_top_bg.jpg" align="center">是否<br>审核</td>
									<td width="23%" background="images/admin_top_bg.jpg" align="center">操作选项</td>
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
													Response.Write("固顶")
												End If
												If rs("Hot") = 0 Then
													Response.Write("&nbsp;")
												Else
													Response.Write("&nbsp;推荐")
												End If
											%>
										</font>
									</td>
									<td width="5%" align="center">
										<font color=#FF0000>
											<%
												If rs("Passed") = 1 Then
													Response.Write("是")
												Else
													Response.Write("否")
												End If
											%>
										</font>
									</td>
									<td width="23%">
										&nbsp;<a href="Admin_Edit_Article.asp?Articleid=<%=rs("ArticleID")%>&classid=<%=classid%>">修改</a>
										 | <a href="Admin_Man_Article.asp?Flag=del&zt=2&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>">删除</a>
										  | <a href="Admin_Man_Article.asp?Flag=del&zt=4&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>"><%If rs("OnTop") = 0 Then%>固顶<%Else%>解固<%End If%></a>
										   | <a href="Admin_Man_Article.asp?Flag=del&zt=6&Articleid=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>"><%If rs("Hot") = 0 Then%>推荐<%Else%>不推荐<%End If%></a> 
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
										<a href="Admin_Man_Article.asp?pageno=1&Classid=<%=classid%>">首页</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_Man_Article.asp?pageno=<%=pageno-1%>&Classid=<%=classid%>">上一页</a> 
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
										<a href="Admin_Man_Article.asp?pageno=<%=pageno+1%>&Classid=<%=classid%>">下一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">下一页</font>
										<%
											end if
										%>
										<a href="Admin_Man_Article.asp?pageno=<%=rs.pagecount%>&Classid=<%=classid%>">尾页</a>
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
								<tr height="25"> 
									<td colspan="7">
										&nbsp;<input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'ArticleID_1'); else SelectAllItem(false, 'ArticleID_1');" type="checkbox" value="1" name="select_all" > 全选 
										<input type="button" value="删除选定的信息" name="B1" onClick="javaScript:if (confirm('确定要删除选中的信息吗。本操作将选中的信息暂时存放到回收站中，必要时您可以从回收站中恢复！ ')) form1.submit();" > 
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