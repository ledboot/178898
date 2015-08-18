<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
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
		<div class="mainA_Top">
			<div class="mainA_TopBG">
				<!--#include file="indexLDZGL.asp" -->
				<!--#include file="indexXLZCL.asp" -->
				<!--#include file="indexJZGCL.asp" -->
			</div>
		</div>
		<div class="index_Bg2"><span><!--背景线条--></span></div>
		<div class="mainA_Center">
			<!--#include file="indexKCJJ.asp" -->
			<div class="mainA_CenterMargin_A"></div>
			<!--#include file="indexKSDT.asp" -->
			<div class="mainA_CenterMargin_A"></div>
			<!--#include file="indexPXXX.asp" -->
		</div>
		<div class="index_Bg3"><span><!--背景线条--></span></div>
		<div class="mainA_Bottom">
			<!--#include file="indexZXZKZX.asp" -->
			<div class="mainA_BottomMargin_A"><!--边距--></div>
			<!--#include file="indexZGKS.asp" -->
			<div class="mainA_BottomMargin_A"><!--边距--></div>
			<!--#include file="indexZLXZ.asp" -->
		</div>
		<!--#include file="foot.asp" -->
	</body>
</html>