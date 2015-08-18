<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
<%
	lmname = "在线报名"
	NavPage = "zxbm.asp?lx=1"
%>
<%
	lx = Cint(Request.QueryString("lx"))
	Select Case lx
		Case 1
			lmname_1 = "在线报名"
			lx = 1
			NavPage_1 = "zxbm.asp?lx=1"
		Case Else
			lmname_1 = "在线报名"
			lx = 1
			NavPage_1 = "zxbm.asp?lx=1"
	End Select
	
	ArticleId = Request("ArticleId")
	If ArticleId = "" Or IsNull(ArticleId) Or IsEmpty(ArticleId) Then
		ArticleId = 0
	End If
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
		<title><%If meta_key_title_2_ny <> "" Then%><%=meta_key_title_2_ny%>-<%End If%><%=lmname%>-<%=website%></title>
		<script language="javascript" src="js/Common_Function.js"></script>
		<script language="javascript" src="js/MSClass.js"></script>
		<link rel="stylesheet" type="text/css" href="css/index.CSS">
	</head>
	<body <%=bgcolor%><%=selectmouse%><%=rightmouse%> style="margin:0px; padding:0px; height:100%">
		<!--#include file="head.asp" -->
		<div class="mainA_Center">
			<div class="mainA_CenterBG">
				<div class="mainB_left">
					<!--#include file="leftMenu.asp" -->
					<!--#include file="leftKCJJ.asp" -->
					<!--#include file="leftKSDT.asp" -->
    			</div>
    			<div class="marginB_A"><!----></div>
    			<div class="main_BRight" id="main_BRight">
					<!--#include file="breadCrumbsNavigation.asp" -->        
        			<div style="height:15px; width:745px; clear:both"></div>
					<!--#include file="ZXBMForm.asp" -->
    			</div>
			</div>
		</div>
		<!--#include file="foot.asp" -->
	</body>
</html>