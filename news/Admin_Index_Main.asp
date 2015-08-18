<!--#include file = "../inc/conn.asp"-->
<%
	session.Contents.Remove("GetCode")
	Dim theInstalledObjects(9)
	theInstalledObjects(0) = "Scripting.FileSystemObject"
	theInstalledObjects(1) = "Adodb.Connection"
	theInstalledObjects(2) = "WScript.Shell"
	theInstalledObjects(3) = "Microsoft.XMLHTTP"
	theInstalledObjects(4) = "Persits.Jpeg"
	theInstalledObjects(5) = "JMail.SmtpMail"
	theInstalledObjects(6) = "CDONTS.NewMail"
	theInstalledObjects(7) = "Adodb.Stream"
	theInstalledObjects(8) = "Scripting.Dictionary"
	theInstalledObjects(9) = "Persits.Upload"
%>
<html>
	<head>
		<title>管理首页</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
	</head>
	<SCRIPT language="JavaScript" runat="server">
		function getEngVerJs(){
			try{
				return ScriptEngineMajorVersion() +"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion() + " ";
			}
			catch(e){
				return "服务器不支持此项检测";
			}
		}
	</SCRIPT>
	<SCRIPT language="VBScript" runat="server">
		Function getEngVerVBS()
			getEngVerVBS=ScriptEngineMajorVersion() &"."&ScriptEngineMinorVersion() &"." & ScriptEngineBuildVersion() & " "
		End Function
	</SCRIPT>
	<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
		<br>
		<center> 
			<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" style="border: 1px solid #C0C0C0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
				<tr> 
					<td height="25" colspan="2" align="center" class="topbg" style="border-bottom:1px solid #C0C0C0; background-image: url('images/admin_top_bg.jpg')">
						<font color="#000000">
							<strong>服 务 器 信 息</strong>
						</font>
					</td>
				</tr> 
				<tr>
					<td width="50%" height="23">　</td>
					<td width="50%">　</td>
				</tr>
				<tr> 
					<td height="23">&nbsp;服务器名：<%=Request.ServerVariables("SERVER_NAME")%>（<%=request.ServerVariables("APPL_PHYSICAL_PATH")%>）</td>
					<td>&nbsp;服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
				</tr>
				<tr> 
					<td width="50%" height="23">&nbsp;服务器端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
					<td width="50%">&nbsp;服务器时间：<%=now()%></td>
				</tr>
				<tr> 
					<td height="23">&nbsp;IIS版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
					<td>&nbsp;脚本超时时间：<%=Server.ScriptTimeout%> 秒</td>
				</tr>
				<tr> 
					<td height="23">&nbsp;Application变量：<%=Application.Contents.Count & "个 "%></td>
					<td>&nbsp;Session变量：<%=Session.Contents.Count&"个 "%></td>
				</tr>
				<tr> 
					<td height="23">&nbsp;JScript解释引擎：<%= getEngVerJs() %></td>
					<td>&nbsp;VBScript解释引擎：<%=getEngVerVBS()%></td>
				</tr>
				<tr> 
					<td width="50%" height="23">&nbsp;AspUpload组件：
						<%
							If Not IsObjInstalled(theInstalledObjects(9)) Then
						%>
						<font color="red"><b>×</b></font>
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
					<td width="50%">&nbsp;ASPJpeg组件： 
						<%
							If Not IsObjInstalled(theInstalledObjects(4)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
				</tr>
				<tr> 
					<td width="50%" height="23">&nbsp;FSO文本读写： 
						<%
							If Not IsObjInstalled(theInstalledObjects(0)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
					<td width="50%">&nbsp;ADO 数据对象： 
						<%
							If Not IsObjInstalled(theInstalledObjects(1)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
				</tr>
				<tr> 
					<td height="23">&nbsp;Adodb.Stream： 
						<%
							If Not IsObjInstalled(theInstalledObjects(7)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
						<td>&nbsp;Scripting.Dictionary： 
						<%
							If Not IsObjInstalled(theInstalledObjects(8)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
				</tr>
				<tr> 
					<td height="23">&nbsp;WScript.Shell： 
						<%
							If Not IsObjInstalled(theInstalledObjects(2)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
					<td>&nbsp;Microsoft.XMLHTTP： 
						<%
							If Not IsObjInstalled(theInstalledObjects(3)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
				</tr>
				<tr> 
					<td width="50%" height="23">&nbsp;Jmail组件支持： 
						<%
							If Not IsObjInstalled(theInstalledObjects(5)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
					<td width="50%">&nbsp;CDONTS组件支持： 
						<%
							If Not IsObjInstalled(theInstalledObjects(6)) Then
						%>
						<font color="red"><b>×</b></font> 
						<%
							else
						%>
						<b>√</b> 
						<%
							end if
						%>
					</td>
				</tr>
				<tr> 
					<td height="23">　</td>
					<td>　</td>
				</tr>
			</table>
			<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" style="border: 1px solid #C0C0C0">
				<tr> 
					<td width="100%" height="23">　</td>
				</tr>
				<tr> 
					<td width="100%" height="23"> 
						<p align="center">
							<b>
								<font color="#FF0000">请选择对应的栏目进行操作</font>
							</b>
						</p>
					</td>
				</tr>
				<tr> 
					<td width="100%" height="23">　</td>
				</tr>
			</table>
			<br>
			Copyright (c) 2008-2010 <%=website%> All Rights Reserved.<BR>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>