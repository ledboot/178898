<!--#include file="../inc/conn.asp" -->
<%
	if session("AdminName") = "" then
		Response.Write(session("AdminName") & "=============>")
		Response.End()
		response.redirect "../login.asp"
	end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<title><%=website%> -- 后台管理</title>
	</head>
	<frameset cols="181,*" framespacing="0" frameborder="1" border="false" id="frame" scrolling="yes">
		<frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="Admin_Index_Left.asp">
		<frameset framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="no">
			<frame name="top" scrolling="no" src="Admin_Index_Top.asp" target="main">
			<frame name="main" scrolling="auto" src="Admin_Index_Main.asp">
		</frameset>
		<noframes>
			<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
				<p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
				<p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
			</body>
		</noframes>
	</frameset>
</html>
<%
	call CloseDatabase()
%>