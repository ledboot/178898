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
	answerId = SafeRequest("answerId",0)
	lmname = "查看答题"
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
<%=meta_des&CHR(10)%><%=meta_key&CHR(10)%><%=meta_cop&CHR(10)%><%=meta_aut&CHR(10)%><%=meta_ner&CHR(10)%><%=meta_pub&CHR(10)%><%=meta_gen%>
<title>
<%If Not checkIsEmpty(titleRs) Then%>
<%=titleRs%>-
<%End If%>
<%If Not checkIsEmpty(SubheadRs) Then%>
<%=SubheadRs%>-
<%End If%>
<%If Not checkIsEmpty(GjzStr) Then%>
<%=GjzStr%>-
<%End If%>
<%If Not checkIsEmpty(meta_key_title) Then%>
<%=meta_key_title%>-
<%End If%>
<%=lmname%>-<%=website%></title>
<script language="javascript" src="../js/jquery-1.4.2.min.js"></script>
<script language="javascript" src="../js/Common_Function.js"></script>
<script language="javascript" src="../js/MSClass.js"></script>
<script language="javascript" src="../js/jquery.cookie.js"></script>
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
    
    <div class="ksmaindiv" style="overflow-y:scroll;">
      <%
		set rsA = rsAnswer2(answerId)
		idAnswer = split(rsA("UserAnswer"),"@")
		For i = 0 To UBound(idAnswer)
			sid = split(idAnswer(i),"_")
			set rsSub = rsSubjectLibrary4(CInt(sid(0))) 
			%>
      <div class="examone" id="main_M">
        <div class="ajshow" id="frm_main" style="display: block; height:auto;">
          <div class="title"> <span id="t"><%=rsSub("Title")%></span> </div>
          <div class="item">
            <div class="left"> <%=rsSub("Options")%><span id="rightanswer" style="display: none;"></span> </div>
          </div>
        </div>
        <div class="tieba">
          <div style="width:612px; margin:0 auto;line-height:20px; font-size:14px;font-family:'微软雅黑';<%if rsSub("Answer") = sid(1) then response.Write("background:#0F0;") else response.Write("background:#F00;") end if%>">
          	答案:
			<%if rsSub("Answer") = "1" then 
				response.Write("正确")
			elseif rsSub("Answer") = "0" then
				response.Write("错误")
			else 
				response.Write(rsSub("Answer"))
			end if%></br>
            您的答案：
			<%if sid(1) = "1" then 
				response.Write("正确")
			elseif sid(1) = "0" then
				response.Write("错误")
			else 
				response.Write(sid(1))
			end if%>
          </div>
        </div>
      </div>
      <%
		next
	%>
    </div>
  </div>
</div>
<!--#include file="../foot.asp" -->
</body>
</html>