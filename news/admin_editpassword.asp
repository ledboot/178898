<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	Action = Request("Action")
%>
<%
	If Action = "Edit" Then
		n_pwd = Request.Form("n_pwd")
		o_pwd = Request.Form("o_pwd")
		Set rs = rsAdmin(Session("AdminName"), Session("AdminPass"))
		If rs("password") <> o_pwd Then
			Error1 "�������ԭ���벻��ȷ��", "Admin_editpassword.asp"
		End If
		rs("password") = n_pwd
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("Admin_editpassword.asp")
		Response.End()
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>�޸�����</title>
		<script language="JavaScript">
			function EditPassword(){
				if (document.formpsw.o_pwd.value == ""){
					alert("������ԭ����");
					document.formpsw.o_pwd.focus();
					return false;
				}
				if (document.formpsw.n_pwd.value.length<5 || document.formpsw.n_pwd.value.length>16){
					alert("������ 5-16 λ������");
					document.formpsw.n_pwd.focus();
					return false;
				}
				if (document.formpsw.n_pwd.value != document.formpsw.q_pwd.value){
					alert("��������ȷ�����벻һ�£�����������");
					document.formpsw.q_pwd.focus();
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
		<center>
			<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td align="center">
						<table border="0" cellpadding="0" cellspacing="0">
							<form method="POST" action="Admin_Editpassword.asp?Action=Edit" onsubmit="return EditPassword();" name="formpsw">
								<tr> 
									<td height="30" align="center"><font color="#FF0000">ע���޸�������´�����Ҫʹ���������½</font></td>
								</tr>
								<tr> 
									<td align="center">----------------- ��ǰ�������޸����� -----------------</td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
								<tr> 
									<td>ԭ���룺 <input class="input1" name="o_pwd" type="text" style="width: 150;" maxlength="16">									  <font color="#FF0000">*</font> 5-16���ַ�</td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
								<tr> 
									<td>�����룺 <input class="input1" type="password" name="n_pwd" style="width: 150;" maxlength="16"> <font color="#FF0000">*</font> 5-16���ַ�</td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
								<tr> 
									<td>ȷ&nbsp; �ϣ� <input class="input1" type="password" name="q_pwd" style="width: 150;" maxlength="16"> <font color="#FF0000">*</font> ����������������ͬ���ַ�</td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
								<tr> 
									<td align="center">----------------------------------------------------</td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
								<tr> 
									<td align="center"><input class="btn1" type="submit" value=" ȷ �� " name="B1">&nbsp;<input class="btn1" type="reset" value=" �� �� " name="B2"></td>
								</tr>
								<tr> 
									<td>&nbsp;</td>
								</tr>
							</form>
						</table>
					</td>
				</tr>
			</table>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>