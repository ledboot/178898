<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	classid = Request("classid")
	Articleid = Request("Articleid")
	flag = Request("flag")
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>编辑信息</title>
		<script language="vbscript">
				sub chat()
					if form1.Title.value="" then
						msgbox "请输入信息标题！",48,"注意！"
						form1.Title.focus()
						exit sub
					end if
					if len(form1.Title.value )<1 or len(form1.Title.value )>80 then
						msgbox "信息标题长度要大于1小于80",48,"注意！"
						form1.Title.focus()
						exit sub
					end if
					form1.submit()
				end sub
			-->
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
		<center>
			<%
				Set rs = rsClass7(classid)
				classname = rs("classname")
				Titlename_1 = rs("Titlename_1")
				Subheadname_1 = rs("Subheadname_1")
				Authorname_1 = rs("Authorname_1")
				CopyFromname_1 = rs("CopyFromname_1")
				Gjzname_1 = rs("Gjzname_1")
				Titlename = rs("Titlename")
				Subheadname = rs("Subheadname")
				Authorname = rs("Authorname")
				CopyFromname = rs("CopyFromname")
				Gjzname = rs("Gjzname")
				rs.Close
				Set rs = Nothing
			%>
			<%
				If flag = "save" Then
					Set rs = rsArticle4(Articleid)      
					Title = Request.Form("Title")
					Subhead = Request.Form("Subhead")
					Author = Request.Form("Author")
					CopyFrom = Request.Form("CopyFrom")
					Gjz = Request.Form("Gjz")
					UpdateTime = Request.Form("UpdateTime")
					content = Request.Form("content")
					IncludePic = Request.Form("IncludePic")
					If IncludePic = "" Then
						IncludePic = 0
					End If
					IncludePic1 = Request.Form("IncludePic1")
					If IncludePic1 = "" Then
						IncludePic1 = 0
					End If
					Editor = Request.Form("Editor")
					indeximg = Request.Form("indeximg")
					indeximg1 = Request.Form("indeximg1")
					Passed = Request.Form("Passed")
					rs("ClassID") = classid
					rs("Title") = Title
					rs("Subhead") = Subhead
					rs("Author") = Author
					rs("CopyFrom") = CopyFrom
					rs("Gjz") = Gjz
					rs("Editor") = Editor
					rs("UpdateTime") = UpdateTime
					rs("Passed") = Passed
					rs("Content") = content
					rs("IncludePic") =IncludePic
					rs("IncludePic1") =IncludePic1
					rs("Indeximg") = indeximg
					rs("Indeximg1") = Indeximg1
					rs.Update
					rs.Close
					Set rs = Nothing
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="65%" align="center">
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td align="center" height="25" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;" background="images/admin_top_bg.jpg"><b>信息修改成功</b></td>
				</tr>
				<tr>
					<%
						Set rs = rsArticle4(Articleid)
					%>
					<td align="center" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0"> 
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<%
								If Titlename_1 = 1 Then
							%>
							<tr height="25"> 
								<td width="20%" align="right"><b><%=Titlename%>：</b></td>
								<td width="80%"><font color="#800000">&nbsp;<%=rs("Title")%></font></td>
							</tr>
							<%
								End If
							%>
							<%
								If Subheadname_1 = 1 Then
							%>
							<tr height="25"> 
								<td width="20%" align="right"><b><%=Subheadname%>：</b></td>
								<td width="80%"><font color="#800000">&nbsp;<%=rs("Subhead")%></font></td>
							</tr>
							<%
								End If
							%>
							<%
								If Authorname_1 = 1 Then
							%>
							<tr height="25"> 
								<td width="20%" align="right"><b><%=Authorname%>：</b></td>
								<td width="80%"><font color="#800000">&nbsp;<%=rs("Author")%></font></td>
							</tr>
							<%
								End If
							%>
							<%
								If CopyFromname_1 =1 Then
							%>
							<tr height="25"> 
								<td width="20%" align="right"><b><%=CopyFromname%>：</b></td>
								<td width="80%"><font color="#800000">&nbsp;<%=rs("CopyFrom")%></font></td>
							</tr>
							<%
								End If
							%>
							<%
								If Gjzname_1 =1 Then
							%>
							<tr height="25"> 
								<td width="20%" align="right"><b><%=Gjzname%>：</b></td>
								<td width="80%"><font color="#800000">&nbsp;<%=rs("Gjz")%></font></td>
							</tr>
							<%
								End If
							%>
							<tr>
								<td align="center" colspan="2" height="30"><p><font color="#800000">◇</font> <a href="Admin_Edit_Article.asp?ArticleID=<%=ArticleID%>&ClassID=<%=classid%>">修改本文</a> <font color="#800000">◇</font> <a href="Admin_Add_Article.asp?ClassID=<%=classid%>">继续添加</a> <font color="#800000">◇</font> <a href="Admin_Man_Article.asp?ClassID=<%=classid%>">管理文章</a></p></td>
							</tr>
						</table>
					</td>
					<%
						rs.Close
						Set rs = Nothing
					%>
				</tr>
				<tr> 
					<td>&nbsp;</td>
				</tr>
			</table>
			<%
				Else
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;"><%=classname%> — 信息修改</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;管理导航：</b><a href="Admin_Add_Article.asp?classid=<%=classid%>">添加该栏目下的信息</a> |&nbsp;<a href="Admin_Man_Article.asp?classid=<%=classid%>">管理该栏目下的信息</a> | <a href="Admin_HszM_Article.asp?classid=<%=classid%>">信息回收站管理</a> | <a href="javascript:history.back()">返回</a></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td height="25" style="border: 1 solid #C0C0C0" background="images/admin_top_bg.jpg">&nbsp;当前状态：<%=classname%>&nbsp;<font color="#FF0000">修改信息</font></td>
				</tr>
				<tr> 
					<td align="center" bgcolor="#F7F7F7" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0"> 
						<%
							Set rs = rsArticle4(Articleid)
						%>
						<form method="POST" action="?Flag=save&Articleid=<%=Articleid%>&classid=<%=classid%>" name="form1">
							<table border="0" cellpadding="0" cellspacing="0">
								<tr height="25" width="80">
									<td>所属栏目：</td>
									<td><strong><font color="#800000"> <%=classname%></font></strong></td>
								</tr>
								<%
									If Titlename_1 = 1 Then
								%>
								<tr height="25"> 
									<td><%=Titlename%>：</td>
									<td><input class="INPUT1"type="text" name="Title" style="width: 420;" maxlength="80" value="<%=rs("Title")%>"> <font color="#FF0000">*</font></td>
								</tr>
								<%
									End If
								%>
								<%
									If Subheadname_1 = 1 Then
								%>
								<tr height="25"> 
									<td><%=Subheadname%>：</td>
									<td><input class="INPUT1" type="text" name="Subhead" style="width: 350;" maxlength="50" value="<%=rs("Subhead")%>"></td>
								</tr>
								<%
									End If
								%>
								<%
									If Authorname_1 = 1 Then
								%>
								<tr height="25"> 
									<td><%=Authorname%>：</td>
									<td><input class="INPUT1" type="text" name="Author" style="width: 350;" maxlength="50" value="<%=rs("Author")%>"></td>
								</tr>
								<%
									End If
								%>
								<%
									If CopyFromname_1 = 1 Then
								%>
								<tr height="25"> 
									<td><%=CopyFromname%>：</td>
									<td><input class="INPUT1" name="CopyFrom" type="text" style="width: 350;" maxlength="50" value="<%=rs("CopyFrom")%>"></td>
								</tr>
								<%
									End If
								%>
								<%
									If Gjzname_1 = 1 Then
								%>
								<tr height="25"> 
									<td><%=Gjzname%>：</td>
									<td><textarea name="Gjz" rows="5" style="width: 350;"><%=rs("Gjz")%></textarea></td>
								</tr>
								<%
									End If
								%>
								<tr height="25"> 
									<td>发布时间：</td>
									<td><input class="INPUT1" name="UpdateTime" type="text" style="width: 120;" value="<%=rs("UpdateTime")%>" maxlength="29"></td>
								</tr>
								<tr height="25"> 
									<td>立即审核：</td>
									<td><input type="checkbox" name="Passed" value="1" <%If rs("Passed") = 1 Then%>checked<%End If%>>是 <font color="#0000ff">（选中后此信息将前台将直接显示。）</font></td>
								</tr>
								<tr> 
									<td colspan="2"><textarea name="content" style="display:none"><%=Server.HtmlEncode(rs("Content"))%></textarea><IFRAME ID="eWebEditor1" SRC="Zzz4eweb/ewebeditor.htm?id=content&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME></td>
								</tr>
								<SCRIPT LANGUAGE="JavaScript">
									<!--
										function showadv(){
											if (document.form1.IncludePic.checked == true) {
												document.getElementById("SuolueTable").style.display = "";
												document.getElementById("SizeTable").style.display = "";
												document.form1.indeximg1.value='';
												window.frames["indeximg_1"].location = 'photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=1&ThumbnailSize=0'
											}else{
												document.getElementById("SuolueTable").style.display = "none";
												document.getElementById("SizeTable").style.display = "none";
												document.form1.indeximg1.value='';
												window.frames["indeximg_1"].location = 'photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=0&ThumbnailSize=0'
											}
										}
										function showadv1(){
											if (document.form1.IncludePic1.checked == true) {
												window.frames["indeximg_1"].location = 'photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=1&ThumbnailSize=1'
											}else{
												window.frames["indeximg_1"].location = 'photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=1&ThumbnailSize=0'
											}
										}
									//-->
								</SCRIPT>
								<tr height="25"> 
									<td>缩略图：</td>
									<td>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr> 
												<td><input type="checkbox" name="IncludePic" value="1" <%If rs("IncludePic") = 1 Then%>checked<%End If%> onClick="showadv()"> 是</td>
												<td>
													<table width="100%" border="0" cellspacing="0" cellpadding="0" id="SizeTable" style='display:none'>
														<tr> 
															<td><input type="checkbox" name="IncludePic1" value="1" <%If rs("IncludePic1") = 1 Then%>checked<%End If%> onClick="showadv1()"> 图片尺寸是否为横图（默认为竖图）</td>
														</tr>
													</table>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="25"> 
									<td>信息录入： </td>
									<td><input class="INPUT1" type="text" name="editor" size="20" readonly value="<%=rs("Editor")%>"></td>
								</tr>
								<tr height="25"> 
									<td>文档上传：</td>
									<td><iframe name="indeximg_1" src="photo_add.asp?up_name=indeximg&action=&ThumbnailComponent=0&ThumbnailSize=0" width="94%" height=25 scrolling="no" marginWidth=1 frameSpacing=0 marginHeight=0 frameBorder=0></iframe></td>
								</tr>
								<tr height="25"> 
									<td>上传地址：</td>
									<td><input class="INPUT1" type="text" name="indeximg" size="45" value="<%=rs("indeximg")%>"></td>
								</tr>
								<tr id="SuolueTable" height="25" style='display:none'> 
									<td>缩略图地址：</td>
									<td><input class="INPUT1" type="text" name="indeximg1" size="45" value="<%=rs("indeximg1")%>" readonly></td>
								</tr>
								<tr> 
									<td></td>
									<td></td>
								</tr>
								<tr height="25"> 
									<td>&nbsp; <input type="button" value=" 确 定" name="B1" onClick="chat()"></td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</form>
						<%
							rs.Close
							Set rs = Nothing
						%>
					</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
				</tr>
			</table>
			<%
				End If
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>