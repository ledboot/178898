<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Conn.asp" -->
<!--#include file="news/Admin_Config.asp" -->
<%
	lmname = "分类信息"
	NavPage = "flxx.asp"
%>
<%
	If IsNumeric(request.QueryString("id")) = False Then
		response.write "请通过页面上的链接进行操作，不要试图破坏系统。"
		response.end
	End If
	id=cint(request.QueryString("id"))
	Set rs=Server.createobject("adodb.recordset")
	Sql="select Title,Subhead,Author,CopyFrom,Content,Indeximg,UpdateTime,Hits,ClassID,ArticleID,Gjz from Article where ArticleID="&id&" and Deleted=0 and Passed=1"
	rs.open sql,conn,1,1
	if (rs.bof and rs.eof) then
		response.write "请通过页面上的链接进行操作，不要试图破坏系统。"
		response.end
	end if
	select case rs("ClassID")
		case 113100
			lx = 1
			lmname_1="建筑工程类"
			NavPage_1 = "flxx.asp?lx=1"
		case 113102
			lx = 2
			lmname_1="学历职称类"
			NavPage_1 = "flxx.asp?lx=2"
		case 113101
			lx = 3
			lmname_1="劳动资格类"
			NavPage_1 = "flxx.asp?lx=3"
		case else
			response.write "请通过页面上的链接进行操作，不要试图破坏系统。"
			response.end
	end select
	Title=rs("Title")
	Subhead=rs("Subhead")
	Author=rs("Author")
	CopyFrom=rs("CopyFrom")
	Content=rs("Content")
	Indeximg=rs("Indeximg")
	UpdateTime=rs("UpdateTime")
	ArticleID = rs("ArticleID")
	Gjz = rs("Gjz")
	Hits=rs("Hits")+1
	rs.close
	set rs=nothing
	sql="update Article set hits="&Hits&" where articleID="&id
	conn.execute(sql)
	
	breadCrumbsNavigationTitle = Title
	breadCrumbsNavigationArticleID = ArticleID
	GjzMeta_key_title = ""
	If Not checkIsEmpty(Gjz) Then
		GjzArray = Split(Gjz, ",")
		
		For iCount = 0 To UBound(GjzArray)
			If iCount = 0 Then
				GjzMeta_key_title = GjzArray(iCount)
			Else
				GjzMeta_key_title = GjzMeta_key_title & "|" & GjzArray(iCount)
			End If
		Next
	End If
	
	showNavPage = "showzgrz.asp"
	meta_key_title = "资格认证"
	If Not checkIsEmpty(GjzMeta_key_title) Then
		meta_key_title = GjzMeta_key_title & "-" & meta_key_title
	End If
		
	If Not checkIsEmpty(Content) Then
		meta_des = "<meta name=""Description"" content=""" & RemoveHTML(Content, "") & """>"
	End If
	
	rsPreviousStrTemp = rsPreviousStr(ArticleID)
	rsNextArticleStrTemp = rsNextArticleStr(ArticleID)
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
		<title><%If Not checkIsEmpty(title) Then%><%=title%>-<%End If%><%If Not checkIsEmpty(GjzStr) Then%><%=GjzStr%>-<%End If%><%If Not checkIsEmpty(Subhead) Then%><%=Subhead%>-<%End If%><%If Not checkIsEmpty(meta_key_title) Then%><%=meta_key_title%>-<%End If%><%=lmname%>-<%=website%></title>
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
					<!--#include file="FinallyShows.asp" -->
    			</div>
			</div>
		</div>
		<!--#include file="foot.asp" -->
	</body>
</html>