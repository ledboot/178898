<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
	
	stuId = Request("stuId")
	
	conn.Execute("delete from registrationInfo where id=" & stuId)
	
	Call Error1("ɾ���ɹ�", "Online_RegistrationManage.asp")
%>
<%
	Call CloseDatabase()
%>