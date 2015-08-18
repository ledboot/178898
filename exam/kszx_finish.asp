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
	lmname = "考试成绩"
	NavPage = "kszx.asp"
%>
<%

	lmname_1 = lmname
	NavPage_1 = NavPage
	meta_key_title = lmname
	meta_key_title = lmname

%>

<%
	set rsA = rsAnswer2(answerId)
	studentId = session("StudentId")
	score = 0
	paperId = rsA("PaperId")
	set rsP = rsPaper2(paperId)
	TotalScore = rsP("TotalScore")
	ExamTime = rsP("ExamTime")
	idAnswer = split(rsA("UserAnswer"),"@")
	For i = 0 To UBound(idAnswer)
		sid = split(idAnswer(i),"_")
		set rsSub = rsSubjectLibrary4(CInt(sid(0))) 
		if sid(1) = rsSub("Answer") then
			score = score+CInt(rsSub("Score"))
		end if
		rsSub.close
	set rsSub = nothing
	Next
	rsP.close
	set rsP = nothing
	rsA.close
	set rsA = nothing
	set rsS =rsScore1()
	rsS.AddNew
	rsS("StudentId") = studentId
	rsS("PaperId") = paperId
	rsS("AnswerId") = answerId
	rsS("Score") = score
	rsS.update
	rsS.close
	set rsS = nothing
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
<!--#include file="../head.asp" -->
<noscript>
		<iframe style="display:none;" src="*.htm"></iframe> 
		</noscript>
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
    
    <table width="720" height="121" align="center" cellpadding="0" cellspacing="0" bgcolor="#f1f1f1" style="font-size:14px">
      <tbody>
        <tr bgcolor="#ffffff" height="40">
          <td height="30" colspan="4" align="center" bgcolor="#fef3de" style="font-weight:bold;font-size:15px;"><strong><%=session("StudnetName")%></strong>考试情况表</td>
        </tr>
        <tr bgcolor="#ffffff" height="30">
          <td width="102" align="center" bgcolor="#FFFFFF">试卷总分：</td>
          <td width="254" align="left" bgcolor="#FFFFFF"><%=TotalScore%>分</td>
          <td width="102" align="center" bgcolor="#FFFFFF">考生得分：</td>
          <td width="302" align="left" bgcolor="#FFFFFF"><font color="#FF0000"><strong><%=score%></strong> 分</font></td>
        </tr>
        <tr bgcolor="#ffffff" height="30">
          <td align="center" bgcolor="#f3f3f3">考试时间：</td>
          <td align="left" bgcolor="#f3f3f3"><%=ExamTime%>分钟</td>
          <td align="center" bgcolor="#f3f3f3"></td>
          <td align="left" bgcolor="#f3f3f3"></td>
        </tr>
      </tbody>
    </table>
    <div>
    	<p align="center"><a class="inp" href="kszx_explained.asp?answerId=<%=answerId%>">查看试卷答案</a></p>
    </div>
        
  </div>
</div>
<!--#include file="../foot.asp" -->
</body>
</html>