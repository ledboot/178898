<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	lmname = "报名信息"
%>
<html>
	<head>
		<%=meta_des&CHR(10)%>
		<%=meta_key&CHR(10)%>
		<%=meta_cop&CHR(10)%>
		<%=meta_aut&CHR(10)%>
		<%=meta_pub&CHR(10)%>
		<%=meta_gen%>
		<title><%=website%>-<%=lmname%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link href="../css/css.css" rel="stylesheet" type="text/css">
		<script language="JavaScript">
		<!--
		function check(id){
			//v2.0
				if (confirm('请确定是否要删除这条记录'))
					location.href='Online_Registration_del.asp?id=' + id;
				}
			//-->
		</script>
	</head>
	<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
		<center>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td height="58" align="center" valign="middle"><font color="#7F7F7F" size="+1"><b><%=lmname%></b></font></td>
                </tr>
            </table>
            <table width="100%" border="0" cellpadding="3" cellspacing="1">
                <tr> 
                    <td width="150" height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">姓名</font></td>
                    <td width="75" height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">电话</font></td>
                    <td height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">地址</font></td>
                    <td width="150" height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">注册时间</font></td>
                    <td height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">说 &nbsp;&nbsp;&nbsp;明</font></td>
                    <td width="75" height="25" align="center" background="images/admin_top_bg.jpg"><font style="font-size:12px">管理操作</font></td>
                </tr>
            </table>
            <table width="100%" border="0">
                <tr><td>&nbsp;</td></tr>
            </table>
            <table width="100%" border="0" cellpadding="3" cellspacing="1">
                <%
                    Dim countForm
                    countForm = 0
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT id, stuName, stuSex, stuBirthday, stuI_D_Card, stuCensus_Register_Provinces_ID, stuCensus_Register_City_ID, stuGraduation_School, stuCultural_Degree_ID, stuHome_Address, stuTelephone, stuZip, stuHeight, stuEyesight_Left, stuEyesight_Right, stuDeclare_Class_Type_ID, stuResume_Text, RegisterDatatime, stuAddress, stuQQNumber, stuEMail FROM registrationInfo"
                    rs.Open sql, conn, 1, 3
                    If not rs.bof and not rs.eof Then                      
                        While Not rs.BOF And Not rs.EOF
                            stuNameRs = rs("stuName")
                            stuTelephoneRs = rs("stuTelephone")
                            RegisterDatatimeRs = rs("RegisterDatatime")
                            stuAddressRs = rs("stuAddress")
							stuResume_TextRs = rs("stuResume_Text")
							idRs = rs("id")
							
							If Not checkIsEmpty(stuResume_TextRs) And stuResume_TextRs <> "未填写" Then
								stuResume_TextRsStr = stuResume_TextRs
							Else
								stuResume_TextRsStr = "&nbsp;"
							End If
                %> 
                <tr>
                    <td width="150" align="center"><font style="font-size:12px"><%=stuNameRs%>&nbsp;</font></td>
                    <td width="75" align="center"><font style="font-size:12px"><%=stuTelephoneRs%>&nbsp;</font></td>
                    <td align="center"><font style="font-size:12px"><%=stuAddressRs%>&nbsp;</font></td>
                    <td width="150" align="center"><font style="font-size:12px"><%=RegisterDatatimeRs%>&nbsp;</font></td>
                    <td align="left"><font style="font-size:12px"><%=stuResume_TextRsStr%></font></td>
                    <td align="center" width="75"><a style="font-size:12px; color:#06F; text-decoration:none" href="delOnline_Registration.asp?stuId=<%=idRs%>">删除</a></td>
                </tr>
                <%
                            countForm = countForm + 1
                            rs.movenext
                        Wend
                    else
                        response.write ""
						'response.write "内容正在整理中..."
                    end if
                    rs.Close
                    Set rs = Nothing
                %>
            </table>
            <table width="100%" border="0" cellpadding="1" cellspacing="1">
                <tr><td colspan="5">&nbsp;</td></tr>
            </table>
            <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                <tr><td>&nbsp;</td>
                </tr>
            </table>
		</center>
	</body>
</html>
<%
	Call CloseDatabase()
%>