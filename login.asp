<!--#include file="inc/conn.asp" -->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<title>管理员登录</title>
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
                	<td class="loginBodyHeadTitle">欢迎登录<%=website%>网站后台管理系统</td>
                </tr>
            </table>
			<form method="post" name="loginFrom" id="loginForm" action="chklogin.asp" target="_top">
                <table width="380" height="160" border="0" cellspacing="0" cellpadding="1" align="center">
                    <tr>
                        <td>
                            <table width="380" border="0" cellspacing="1" cellpadding="1">
                                <tr> 
                                    <td class="loginFormTitle">用户名：</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput" type="text" name="username" id="username" maxlength="27"></td>
                                </tr>
                                <tr> 
                                    <td class="loginFormTitle">密&nbsp; 码：</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput" type="password" name="password" id="password" maxlength="27"></td>
                                </tr>
                                <tr> 
                                    <td class="loginFormTitle">验证码：</td>
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
					<td class="loginLinkBox"><a href="<%=webAddress%>" class="loginLink">网站首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="lxwm.asp" class="loginLink">联系我们</a></td>
				</tr>
			</table>
		</center>
	</body>
</html>
<%
	Call CloseDatabase()
%>