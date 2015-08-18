<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/Conn.asp" -->
<!--#include file="../news/Admin_Config.asp" -->
<%
	lmname = "考试中心"
	NavPage = "kszx.asp"
%>
<%

	lmname_1 = lmname
	NavPage_1 = NavPage
	meta_key_title = lmname
	meta_key_title = lmname
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<%=meta_des&CHR(10)%>
		<%=meta_key&CHR(10)%>
		<%=meta_cop&CHR(10)%>
		<%=meta_aut&CHR(10)%>
		<%=meta_ner&CHR(10)%>
		<%=meta_pub&CHR(10)%>
		<%=meta_gen%>
		<title><%If Not checkIsEmpty(titleRs) Then%><%=titleRs%>-<%End If%><%If Not checkIsEmpty(SubheadRs) Then%><%=SubheadRs%>-<%End If%><%If Not checkIsEmpty(GjzStr) Then%><%=GjzStr%>-<%End If%><%If Not checkIsEmpty(meta_key_title) Then%><%=meta_key_title%>-<%End If%><%=lmname%>-<%=website%></title>
		<script language="javascript" src="../js/Common_Function.js"></script>
		<script language="javascript" src="../js/MSClass.js"></script>
		<link href="../css/index.css" rel="stylesheet" type="text/css" />
		<script language="Javascript"> 
			document.oncontextmenu=new Function("event.returnValue=false"); 
			document.onselectstart=new Function("event.returnValue=false"); 
		</script> 
	</head>
	<body <%=bgcolor%><%=selectmouse%><%=rightmouse%> style="margin:0px; padding:0px; height:100%">
		<noscript>
		<iframe style="display:none;" src="*.htm"></iframe> 
		</noscript>
		<!--#include file="../head.asp" -->
		<div class="mainA_Center">
			<div class="mainA_CenterBG">
				<div class="mainB_left">
					<!--#include file="../leftMenu.asp" -->
					<!--#include file="../leftKCJJ.asp" -->
					<!--#include file="../leftKSDT.asp" -->
    			</div>
    			<div class="marginB_A"><!----></div>
    			<div class="main_BRight" id="main_BRight">
					<!--#include file="../breadCrumbsNavigation.asp" -->        

			<form method="post" name="loginFrom" style="height:500px;" id="loginForm" action="kszx_chklogin.asp" target="_top">
                <table width="380" height="160" border="0" cellspacing="0" cellpadding="1" align="center">
                    <tr>
                        <td>
                            <table width="380" border="0" cellspacing="1" cellpadding="1">
                                <tr> 
                                    <td class="loginFormTitle">身份证：</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput" type="text" name="username" id="username" maxlength="27"></td>
                                </tr>
                                <tr> 
                                    <td class="loginFormTitle">验证码：</td>
                                    <td class="loginFormInputBox"><input class="loginFormInput1" name="Passcode" id="Passcode" type="text" size="8" maxlength="4"> <img src="../inc/VerifyCode.asp" ></td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr align="center">
                        <td height="40"><input type="submit" name="submitButton" id="submitButton" class="submitButton" value="">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="resetButton" id="resetButton" class="resetButton" value=""></td>
                    </tr>
                </table>
            </form>
        
    			</div>
			</div>
		</div>
		<!--#include file="../foot.asp" -->
	</body>
</html>