<!--#include file="../inc/conn.asp" -->
<!--#include file="../news/Admin_Config.asp"-->
<!--#include file="../inc/md5.asp" -->
<%
	'Response.Write(Session("VerifyCode") & "=====<br>")
	'Response.End()
	username=trim(SafeRequest("username", 0))
	if checkIsEmpty(username) then
		Error1 "对不起，身份证号码不合法，请重新输入!", "kszx.asp"
	end if
	If strlen(username) > 27 Then
		Error1 "对不起，身份证号码不合法，请重新输入!", "kszx.asp"
	End If
	i = 1
	len1 = 0
	charinner = mid(username,1,1)
	while not charinner = ""
		if (charinner="'" or charinner="-" or charinner="%" or charinner="<" or charinner=">" or charinner="&") then
			Error1 "对不起，身份证号码不合法，请重新输入!", "kszx.asp"
		else
			len1 = len1 + 1
		end if  
		i = i + 1
		charinner = mid(username,i,1)
	wend
	if not isnumeric(SafeRequest("passcode", 1)) then
		Error1 "登录失败！验证码必须是数字，请正确填写！", "kszx.asp"
	end if
	passcode=Cint(SafeRequest("passcode", 1))

	if CInt(passcode)<>CInt(Session("VerifyCode")) then
		Error1 "登录失败！验证码错误！", "kszx.asp"
	end if
	set rs=rsStudnet3(username)
	
	If Rs.EOF AND Rs.BOF then
		Session.Abandon
		Error1 "您输入的身份证号码不正确，\n请返回重新输入,谢谢！", "kszx.asp"
	Else
		GetIp()	
		conn.execute("update GW_Student set LastloginIP='"&getip&"' LastloginTime=now() where IdCard='"&username& "'")
		Session("StudentId") = Rs("id")
		Session("StudnetName")=Rs("StudnetName")
		Session("IdCard")=Rs("IdCard")
		Session("ProfessionId")=Rs("ProfessionId")
		response.redirect "kszx_sel.asp?flag=man"
		response.end
	End if
	Rs.close
	Set rs = Nothing
	session.timeout=6000
%>
<%
	call CloseDatabase()
%>