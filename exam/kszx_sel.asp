<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/Conn.asp" -->
<!--#include file="../news/Admin_Config.asp" -->
<%
	if session("IdCard") = "" then
		response.redirect "kszx.asp"
	end if
%>
<%
	flag = SafeRequest("flag", 0)
	typs = SafeRequest("typs", 0)
	lmname = "考试中心"
	NavPage = "kszx.asp"
%>
<%

	lmname_1 = lmname
	NavPage_1 = NavPage
	meta_key_title = lmname
	meta_key_title = lmname
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
		<script language="javascript" src="../js/Common_Function.js"></script>
		<script language="javascript" src="../js/MSClass.js"></script>
		<link href="../css/index.css" rel="stylesheet" type="text/css" />
<script language="Javascript"> 
			document.oncontextmenu=new Function("event.returnValue=false"); 
			document.onselectstart=new Function("event.returnValue=false"); 
		</script> 
	</head>
	<body <%=bgcolor%><%=selectmouse%><%=rightmouse%> style="margin:0px; padding:0px; height:100%">
<noscript>
		<iframe style="display:none;" src="*.htm"></iframe> 
		</noscript>
		<!--#include file="../head.asp" -->
		<div class="mainA_Center">
			<div class="mainA_CenterBG">
				<div class="mainB_left">
					<!--#include file="../leftMenu.asp" -->
					<!--#include file="../leftKCJJ.asp" -->
					<!--#include file="../leftKSDT.asp" -->
    			</div>
    			<div class="marginB_A"><!----></div>
    			<div class="main_BRight" id="main_BRight">
					<!--#include file="../breadCrumbsNavigation.asp" -->    
                    <%
						if flag = "man" and typs = 0 then
					%>
                    <div><a href="kszx_moni.asp">模拟考试</a><br /><a href="kszx_training.asp">复习平台</a></div>
                    
                    <%
						end if
					%>
         		</div>
			</div>
		</div>
		<!--#include file="../foot.asp" -->
	</body>
</html>