<!--#include file="inc/conn.asp" -->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<title>����Ա��¼</title>
		<link rel="stylesheet" href="css/index.css" type="text/css">
	</head>
	<body class="loginBody">
		<center>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0">
            	<tr>
                	<td class="loginBodyHead">&nbsp;</td>
                </tr>
            </table>
            <table width="380" border="0" cellpadding="0" cellspacing="0">
            	<tr>
                	<td class="loginBodyHeadTitle">��ӭ��¼<%=website%>��վ��̨����ϵͳ</td>
                </tr>
            </table>
			<form method="post" name="loginFrom" id="loginForm" action="chklogin.asp" target="_top">
                <table width="380" height="160" border="0" cellspacing="0" cellpadding="1" align="center">
                    <tr>
                        <td>
                            <table width="380" border="0" cellspacing="1" cellpadding="1">
                                <tr> 
                                    <td class="loginFormTitle">�û�����</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput" type="text" name="username" id="username" maxlength="27"></td>
                                </tr>
                                <tr> 
                                    <td class="loginFormTitle">��&nbsp; �룺</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput" type="password" name="password" id="password" maxlength="27"></td>
                                </tr>
                                <tr> 
                                    <td class="loginFormTitle">��֤�룺</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput1" name="Passcode" id="Passcode" type="text" size="8" maxlength="4"> <img src="inc/VerifyCode.asp" ></td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr align="center">
                        <td height="40"><input type="submit" name="submitButton" id="submitButton" class="submitButton" value="">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="resetButton" id="resetButton" class="resetButton" value=""></td>
                    </tr>
                </table>
            </form>
			<table width="380" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="loginLinkBox"><a href="<%=webAddress%>" class="loginLink">��վ��ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="lxwm.asp" class="loginLink">��ϵ����</a></td>
				</tr>
			</table>
		</center>
	</body>
</html>
<%
	Call CloseDatabase()
%>