<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type="text/css" rel="stylesheet">
		<title>������</title>
	</head>
	<body topmargin="0" background="images/kabg.jpg" leftmargin="0">
		<center>
			<table border="0" cellpadding="5" cellspacing="0" width="169" style="border: 1 solid #C0C0C0;">
				<tr align="center">
					<td>��վ��̨����</td>
				</tr>
				<tr align="center">
					<td style="border-bottom: 1 solid #C0C0C0">
						<a href="Admin_Index_Main.asp" target="main"><b>������ҳ</b></a> | <b><a target="_top" href="logout.asp">��ȫ�˳�</a></b>
					</td>
				</tr>
				<tr>
					<td bgcolor="#F7F7F7">
						<table cellpadding="0" cellspacing="0" align="center">
							<tr>
								<td>�û�����<font color="#FF0000"><%=Session("AdminName")%></font></td>
							</tr>
							<tr>
								<td>Ȩ&nbsp;&nbsp;�ޣ�<font color="#FF0000"><%If CInt(Session("Purview")) = 1 Then%>��������Ա<%Else%>��ͨ����Ա<%End If%></font></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table border=0 cellpadding=5 cellspacing=0 width=169 style="border-left: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">
			  <tr>
					<td>
						<script language="JavaScript" src="MzTree/MzTreeView10.js"></script>
						<script language="JavaScript">
							window.tree = new MzTreeView("tree");
							tree.icons["property"] = "property.gif";
							tree.icons["css"] = "collection.gif";
							tree.icons["book"]  = "book.gif";
							tree.iconsExpand["book"] = "bookopen.gif"; //չ��ʱ��Ӧ��ͼƬ
							tree.setIconPath("./MzTree/"); //�������·��
							tree.nodes["-1_0"] = "text:��վ��Ŀ;";
						</script>
						<%
							Set rs = rsClass()
							sql = "select * from Articleclass Order by Class_sx"
							rs.Open sql, conn, 1, 3
							Dim quanxian, QX, QX_1
							If CInt(Session("Purview")) = 2 And ((session("AdminSc") <> "") Or (Session("AdminSh") <> "")) Then
								quanxian = Session("AdminSc") & "," & Session("AdminSh")
								QX = Split(quanxian, ",")
								QX_1 = F(QX)
							End If
							While Not rs.EOF And Not rs.BOF
								If CInt(Session("Purview")) = 1 Then
									If CInt(rs("ClassNameD")) = 1 Then
										Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";url:" & rs("ClassLink") & ";target:main;"";</script>")
									Else
										If CInt(rs("Class_num")) = 0 Then
											Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";url:Admin_Man_Article.asp?classid=" & rs("XcClassid") & ";target:main;"";</script>")
										Else
											Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";"";</script>")
										End If
									End If
								Else
									If CInt(Session("Purview")) = 2 And ((session("AdminSc") <> "") Or (Session("AdminSh") <> "")) Then
										For i = 0 To UBound(QX_1)
											If CLng(QX_1(i)) = rs("XcClassid") Then
												If CInt(rs("ClassNameD")) = 1 Then
													Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";url:" & rs("ClassLink") & ";target:main;"";</script>")
												Else
													If CInt(rs("Class_num")) = 0 Then
														Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";url:Admin_Man_Article.asp?classid=" & rs("XcClassid") & ";target:main;"";</script>")
													Else
														Response.Write("<script language=""JavaScript"">tree.nodes[""" & rs("ScClassid") & "_" & rs("XcClassid") & """] = ""text:" & rs("ClassName") & ";"";</script>")
													End If
												End If
												Exit For
											End If
										Next
									End If
								End If
								rs.movenext
							Wend
							rs.Close
							Set rs = Nothing
						%>
						<script language="javascript">
							<!--
								document.write(tree.toString());    //����� obj.innerHTML = tree.toString();
							//-->
						</SCRIPT>
					</td>
				</tr>
			</table>
			<table cellpadding="0" cellspacing="0" width="169" style="border: 1 solid #C0C0C0" bgcolor="#F7F7F7">
			  <tr> 
					<td height="25" background="images/admin_top_bg.jpg" style="border-bottom: 1 solid #C0C0C0"><font color="#800000"><b>������Ϣ</b></font></td>
				</tr>
				<tr>
					<td><BR>&nbsp;������������³�<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����Ԫ<BR>&nbsp;����֧�֣�<a href="http://Www.599dns.Com/" target="_blank">��ͨ����</a><BR>&nbsp;</td>
				</tr>
			</table>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>