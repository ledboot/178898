<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
<%
	lmname = "考试动态"
	NavPage = "ksdt.asp"
%>
<%
	lx = Cint(Request.QueryString("lx"))
	Select Case lx
		Case 1
			lmname_1 = "考试动态"
			lx = 1
			NavPage_1 = "ksdt.asp?lx=1"
		Case 2
			lmname_1 = "最新招考资讯"
			lx = 2
			NavPage_1 = "ksdt.asp?lx=2"
		Case Else
			lmname_1 = "考试动态"
			lx = 1
			NavPage_1 = "ksdt.asp?lx=1"
	End Select
	
	showNavPage = "showksdt.asp"
	meta_key_title = "考试动态"
	meta_des = "<meta name=""Description"" content=""考试动态"">"
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
		<script language="javascript" src="js/Common_Function.js"></script>
		<script language="javascript" src="js/MSClass.js"></script>
		<link href="css/index.css" rel="stylesheet" type="text/css" />
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
					<!--#include file="textList.asp" -->
        			<div style="height:20px; width:748px; clear:both"></div>
					<!--#include file="TextNavigation.asp" -->
    			</div>
			</div>
		</div>
		<!--#include file="foot.asp" -->
	</body>
</html>