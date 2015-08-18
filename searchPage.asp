<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
<%
	NavPage = "searchPage.asp"
	defaultSearchClassIdStr = "102100,103100,103101,104100,104101,105100,105101,105102,105103,106100,107100,113100,113101,113102"
	searchWordStr = SafeRequest("searchWordStr", 0)
	searchClassId = SafeRequest("searchClassId", 1)
	
	If Not IsSafeStr(searchWordStr) Then
		Error2("请输入有效的关键词")
	End If
	
	If searchClassId <> 0 Then
		If searchClassId < 102100 Or searchClassId > 113102 Then
			Error2("请选择有效的搜索类别")
		End If
	End If
	
	If searchClassId > 0 Then
		If searchClassId = 102100 Or searchClassId = 103100 Or searchClassId = 104100 Or searchClassId = 105100 Or searchClassId = 106100 Or searchClassId = 107100 Or searchClassId = 113100 Then
			lx = 1
		ElseIf searchClassId = 103101 Or searchClassId = 104101 Or searchClassId = 105101 Or searchClassId = 113101 Then
			lx = 2
		ElseIf searchClassId = 105102 Or searchClassId = 113102 Then
			lx = 3
		ElseIf searchClassId = 105103 Then
			lx = 4
		End If
		
		If searchClassId = 102100 Then
			lmname = "课程简介"
			Select Case lx
				Case 1
					lmname_1 = "课程简介"
					lx = 1
					NavPage_1 = "kcjj.asp?lx=1"
				Case Else
					lmname_1 = "课程简介"
					lx = 1
					NavPage_1 = "kcjj.asp?lx=1"
			End Select
			
			showNavPage = "showkcjj.asp"
		ElseIf searchClassId = 103100 Or searchClassId = 103101 Then
			lmname = "培训信息"
			Select Case lx
				Case 1
					lmname_1 = "培训信息"
					lx = 1
					NavPage_1 = "pxxx.asp?lx=1"
				Case 2
					lmname_1 = "资料下载"
					lx = 2
					NavPage_1 = "pxxx.asp?lx=2"
				Case Else
					lmname_1 = "培训信息"
					lx = 1
					NavPage_1 = "pxxx.asp?lx=1"
			End Select
	
			showNavPage = "showpxxx.asp"
		ElseIf searchClassId = 104100 Or searchClassId = 104101 Then
			lmname = "考试动态"
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
		ElseIf searchClassId = 105100 Or searchClassId = 105101 Or searchClassId = 105102 Or searchClassId = 105103 Then
			lmname = "资格认证"
			Select Case lx
				Case 1
					lmname_1 = "建筑工程考试"
					lx = 1
					NavPage_1 = "zgrz.asp?lx=1"
				Case 2
					lmname_1 = "国家职业资格考试"
					lx = 2
					NavPage_1 = "zgrz.asp?lx=2"
				Case 3
					lmname_1 = "成人高等教育考试"
					lx = 3
					NavPage_1 = "zgrz.asp?lx=3"
				Case 4
					lmname_1 = "自学考试"
					lx = 4
					NavPage_1 = "zgrz.asp?lx=4"
				Case Else
					lmname_1 = "建筑工程考试"
					lx = 1
					NavPage_1 = "zgrz.asp?lx=1"
			End Select
			
			showNavPage = "showzgrz.asp"
	End If
%>
<%
	meta_key_title = searchWordStr
	meta_des = "<meta name=""Description"" content=""" & searchWordStr & """>"
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