<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	flag = Request("flag")
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>�����Ϣ</title>
			<script language="vbscript">
				sub chat()
					if form1.ProfessionName.value="" then
						msgbox "������רҵ���ƣ�",48,"ע�⣡"
						form1.Title.focus()
						exit sub
					end if
					if len(form1.ProfessionName.value )<1 or len(form1.ProfessionName.value )>80 then
						msgbox "רҵ���Ƴ���Ҫ����1С��80",48,"ע�⣡"
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
				If flag = "save" Then
					Set rs = rsProfesion()      
					ProfessionName = Request.Form("ProfessionName")
					rs.AddNew
					rs("ProfessionName") = ProfessionName
					rs.Update
					rs.Close
					Set rs = Nothing
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="65%" align="center">
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td align="center" height="25" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;" background="images/admin_top_bg.jpg"><b>רҵ��ӳɹ�</b></td>
				</tr>
				<tr>
					<%
						Set rs = rsProfesion1()
					%>
					<td align="center" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0"> 
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr height="25"> 
								<td width="200px" align="center"><b>רҵ���ƣ�</b></td>
								<td width="200px" align="left"><font color="#800000">&nbsp;<%=rs("ProfessionName")%></font></td>
							</tr>
							<tr>
								<td align="center" colspan="2" height="30"></a> <font color="#800000">��</font> <a href="Admin_Profession_Add.asp">�������</a> <font color="#800000">��</font> <a href="Admin_Profession.asp?flag=man">����רҵ</a></p></td>
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
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">רҵ���� �� ��Ϣ���</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;��������</b><a href="Admin_Profession.asp">�������Ŀ�µ���Ϣ</a> | <a href="javascript:history.back()">����</a></td>
				</tr>
			</table>
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
									<td ><input class="INPUT1" name="ProfessionName" type="text" style="width: 420;" maxlength="80"> <font color="#FF0000">*</font></td>
								</tr>
								<tr height="25" align="center"> 
									<td colspan="2"><input type="button" value="���" name="B1" onClick="chat()"></td>
								</tr>
							</table>
						</form>
					</td>
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