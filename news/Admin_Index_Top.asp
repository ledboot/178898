<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<html>
	<head>
		<title>��̨����ҳ��</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link href="css/index.css" type=text/css rel=stylesheet>
		<style type="text/css">
			td {FILTER: dropshadow(color=#FFFFFF,offx=1,offy=1);}
		</style>
		<SCRIPT language=javascript1.2>
			function preloadImg(src){
				var img=new Image();
				img.src=src
			}
			preloadImg("images/admin_top_open.gif");
			var displayBar=true;
			function switchBar(obj){
				if (displayBar)
				{
					parent.frame.cols="0,*";
					displayBar=false;
					obj.src="images/admin_top_open.gif";
					obj.title="����߹���˵�";
				}
				else{
					parent.frame.cols="180,*";
					displayBar=true;
					obj.src="images/admin_top_close.gif";
					obj.title="�ر���߹���˵�";
				}
			}
		</script>
	</head>
	<body background="images/admin_top_bg.jpg" leftmargin="0" topmargin="0">
		<table width="100%" border=0 cellpadding=0 cellspacing=0>
			<tr>
				<td width=40><img onClick="switchBar(this)" src="images/admin_top_close.gif" title="�ر���߹���˵�" style="cursor:hand"></td>
				<td width=40><img src="images/admin_top_icon_1.gif"></td>
				<td width=100><a href="Admin_Editpassword.asp" target="main">�޸�����</a></td>
				<td width=40><img src="images/admin_top_icon_5.gif" width="30" height="30"></td>
				<td width=100><a href="../login.asp" target="main">�ص�½</a></td>
				<td align="right"><%=website%>��̨����ϵͳ&nbsp; </td>
			</tr>
		</table>
	</body>
</html>
<%
	call CloseDatabase()
%>