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
%>
<%
	If flag = "del" Then
		If zt = 1 Then
			conn.execute("delete from Article where Articleid in(" & ArticleID_1 & ")")
			Response.Redirect("Admin_Hszm_Article.asp?classid=" & classid)
			Response.End()
		End If
		If zt = 2 Then
			conn.execute("delete from Article where Articleid=" & Articleid & "")
			Response.Redirect("Admin_Hszm_Article.asp?classid=" & classid)
			Response.End()
		End If
		If zt = 3 Then
			Set rs = rsArticle4(Articleid)
			rs("Deleted") = 0
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Redirect("Admin_Hszm_Article.asp?classid=" & classid)
			Response.End()
		End If
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>删除信息</title>
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
			<%
				Set rs = rsClass7(classid)
				classname = rs("classname")
				rs.Close
				Set rs = Nothing
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;"><%=classname%> — 信息删除</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;管理导航：</b><a href="Admin_Add_Article.asp?classid=<%=classid%>">添加该栏目下的信息</a> |&nbsp;<a href="Admin_Man_Article.asp?classid=<%=classid%>">管理该栏目下的信息</a> | <a href="Admin_HszM_Article.asp?classid=<%=classid%>">信息回收站管理</a> | <a href="javascript:history.back()">返回</a></td>
				</tr>
			</table>
			<form method="POST" action="?Flag=del&zt=1&pageno=<%=pageno%>&classid=<%=classid%>" name="form1">
				<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" height="25" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0">&nbsp;信息导航：<font color="#FF0000"><%=classname%></font> 全部信息</td>
					</tr>
					<tr> 
						<td bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0">
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
								<tr height="25"> 
									<td width="4%" background="images/admin_top_bg.jpg">　</td>
									<td width="5%" background="images/admin_top_bg.jpg" align="center">编号</td>
									<td width="80%" background="images/admin_top_bg.jpg">&nbsp;信息标题</td>
									<td width="11%" background="images/admin_top_bg.jpg" align="center">操作选项</td>
								</tr>
								<%
									Set rs = rsArticle6(classid)
								%>
								<%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=20
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
								%>
								<tr height="25" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
									<td width="4%" align="center"><input type="checkbox" name="ArticleID_1" value="<%=rs("ArticleID")%>"></td>
									<td width="5%" align="center"><%=rs("ArticleID")%></td>
									<td width="80%">&nbsp;<%=rs("Title")%></td>
									<td width="11%" align="center">
										<a href="Admin_HszM_Article.asp?Flag=del&zt=3&ArticleID=<%=rs("ArticleID")%>&pageno=<%=pageno%>&classid=<%=classid%>">还原</a> 
										<a href="Admin_HszM_Article.asp?Flag=del&zt=2&ArticleID=<%=rs("ArticleID")%>&pageno=<%=pageno%>&Classid=<%=classid%>">彻底删除</a>
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
										<a href="Admin_HszM_Article.asp?pageno=1&Classid=<%=classid%>">首页</a> 
										<%
											if pageno>1 then
										%>
										<a href="Admin_HszM_Article.asp?pageno=<%=pageno-1%>&Classid=<%=classid%>">上一页</a> 
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
										<a href="Admin_HszM_Article.asp?pageno=<%=pageno+1%>&Classid=<%=classid%>">下一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">下一页</font>
										<%
											end if
										%>
										<a href="Admin_HszM_Article.asp?pageno=<%=rs.pagecount%>&Classid=<%=classid%>">尾页</a>
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
									<td colspan="4">
										&nbsp;<input onClick="javascript:if(this.checked==true) SelectAllItem(true, 'ArticleID_1'); else SelectAllItem(false, 'ArticleID_1');" type="checkbox" value="1" name="select_all"> 全选 
										<input type="button" value=" 删除选定的信息" name="B1" onclick="javaScript:if (confirm('确定要删除选中的信息吗。本操作将来不能恢复！ ')) form1.submit();">
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