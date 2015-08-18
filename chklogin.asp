<!--#include file="inc/conn.asp" -->
<!--#include file="news/Admin_Config.asp"-->
<!--#include file="inc/md5.asp" -->
<%
	'Response.Write(Session("VerifyCode") & "=====<br>")
	'Response.End()
	username=trim(SafeRequest("username", 0))
	if checkIsEmpty(username) then
		Error1 "~{6T2;Fp#,SC;'C{;rC\Bk2;:O7(#,GkVXPBJdHk~}!", "login.asp"
	end if
	If strlen(username) > 27 Then
		Error1 "~{6T2;Fp#,SC;'C{;rC\Bk2;:O7(#,GkVXPBJdHk~}!", "login.asp"
	End If
	i = 1
	len1 = 0
	charinner = mid(username,1,1)
	while not charinner = ""
		if (charinner="'" or charinner="-" or charinner="%" or charinner="<" or charinner=">" or charinner="&") then
			Error1 "~{6T2;Fp#,SC;'C{@8SP2;:O7(5DWV7{#,GkVXPBJdHk~}!", "login.asp"
		else
			len1 = len1 + 1
		end if  
		i = i + 1
		charinner = mid(username,i,1)
	wend
	password=trim(SafeRequest("password", 0))
	if checkIsEmpty(password) then
		Error1 "~{6T2;Fp#,SC;'C{;rC\Bk2;:O7(#,GkVXPBJdHk~}!", "login.asp"
	end if
	If strlen(username) > 27 Then
		Error1 "~{6T2;Fp#,SC;'C{;rC\Bk2;:O7(#,GkVXPBJdHk~}!", "login.asp"
	End If
	i = 1
	len1 = 0
	charinner = mid(password,1,1)
	while not charinner = ""
		if (charinner="'" or charinner="-" or charinner="%" or charinner="<" or charinner=">") then
			Error1 "~{6T2;Fp#,C\Bk@8SP2;:O7(5DWV7{#,GkVXPBJdHk~}!", "login.asp"
		else
			len1 = len1 + 1
		end if  
		i = i + 1
		charinner = mid(password,i,1)
	wend
	'Response.Write(username & "<Br>")
	'Response.Write(password & "<br>")
	'Response.End()
	if not isnumeric(SafeRequest("passcode", 1)) then
		Error1 "~{5GB<J'0\#!QiV$Bk1XPkJGJ}WV#,GkU}H7LnP4#!~}", "login.asp"
	end if
	passcode=Cint(SafeRequest("passcode", 1))
	'Response.Write(Session("VerifyCode") & "=====<br>")
	'Response.Write(passcode & "=====<br>")
	'Response.End()
	if CInt(passcode)<>CInt(Session("VerifyCode")) then
		Error1 "~{5GB<J'0\#!QiV$Bk4mNs#!~}", "login.asp"
	end if
	password = MD5(password)
	set rs=rsAdmin(username, password)
	
	If Rs.EOF AND Rs.BOF then
		Session.Abandon
		Error1 "~{DzJdHk5DSC;'C{;rC\Bk2;U}H7#,~}\n~{Gk75;XVXPBJdHk~},~{P;P;#!~}", "login.asp"
	Else
		GetIp()	
		conn.execute("update admin set LastloginIP='"&getip&"' where username='" & replace(username, "'", "") & "'")
		Session("AdminName")=Rs("username")
		Session("AdminPass")=Rs("Password")
		Session("Purview") = rs("Purview")
		session("AdminSc") = rs("AdminSc_Purview_ClassID")
		session("AdminSh") = rs("AdminSh_Purview_ClassID")
		response.redirect "news/Admin_Index.asp"
		response.end
	End if
	Rs.close
	Set rs = Nothing
	session.timeout=6000
%>
<%
	call CloseDatabase()
%>