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
	papaerId = SafeRequest("papaerId",0)
	lmname = "��������"
	NavPage = "kszx.asp"
%>
<%

	lmname_1 = lmname
	NavPage_1 = NavPage
	meta_key_title = lmname
	meta_key_title = lmname
	
	if flag = "start" then
		studentId = Session("StudentId")
		set rs1 = rsAnswer1()
		rs1.AddNew
		rs1("StudentId") = studentId
		rs1("PaperId") = papaerId
		rs1("UserAnswer") = ""
		rs1.update
		id = rs1("id")
		rs1.close
		set rs1 = nothing
	elseif flag = "next" then
		id = SafeRequest("id",0)
		sid = SafeRequest("sid",0)
		preAnswer = SafeRequest("as",0)
		sIdAnswer = sid+"_"+preAnswer
		set rs2 = rsAnswer2(id)
		oldAnswer = rs2("UserAnswer")
		newAnswer=""
		if oldAnswer = "" then
			newAnswer = sIdAnswer
		else
			newAnswer =oldAnswer+"@"+sIdAnswer
		end if
		rs2("UserAnswer") = newAnswer
		rs2.update
		rs2.close
		set rs2 = nothing
	elseif flag ="submit" then
		id = SafeRequest("id",0)
		sid = SafeRequest("sid",0)
		preAnswer = SafeRequest("as",0)
		sIdAnswer = sid+"_"+preAnswer
		set rs2 = rsAnswer2(id)
		oldAnswer = rs2("UserAnswer")
		newAnswer=""
		if oldAnswer = "" then
			newAnswer = sIdAnswer
		else
			newAnswer =oldAnswer+"@"+sIdAnswer
		end if
		rs2("UserAnswer") = newAnswer
		rs2.update
		rs2.close
		set rs2 = nothing
		response.redirect "kszx_finish.asp?answerId="+id
	end if
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
<script type="text/javascript">

	document.oncontextmenu=new Function("event.returnValue=false"); 
document.onselectstart=new Function("event.returnValue=false"); 
	function next(pageno,papaerId){
		var types = $("#types").val();
		var id = $("#id").val();
		var sid = $("#sid").val();
		var url  = "kszx_start_moni.asp?flag=next&id="+id+"&pageno="+pageno+"&papaerId="+papaerId+"&sid="+sid;
		var cAnswer = getAnswer(types);
		if(cAnswer == ""){
			$("#nochoose").css("display","inline-block");
			return;
		}
		url +="&as="+cAnswer
		location.href = url;
	}
	function toSubmit(papaerId){
		var id = $("#id").val();
		var types = $("#types").val();
		var sid = $("#sid").val();
		var url  = "kszx_start_moni.asp?flag=submit&id="+id+"&papaerId="+papaerId+"&sid="+sid;
		var cAnswer = getAnswer(types);
		if(cAnswer == ""){
			$("#nochoose").css("display","inline-block");
			return;
		}
		url +="&as="+cAnswer
		location.href = url;
	}
	
	function getAnswer(types){
		var cAnswer =""
		if(types == 1){
			var values = $('input[name="answer11"]:checked').val();
			if(values != undefined){
				cAnswer = values;
			}
		}else if (types == 2){
			var values = "";
			$("input[name='answer22']:checked").each(function(){ 
				values+=$(this).val()+","; 
			});
			if(values != ""){
				values = values.substring(0,values.length-1);
				cAnswer = values;
			}
		}else if (types == 3){
			var values = "";
			$("input[name='answer33']:checked").each(function(){ 
				values+=$(this).val()+","; 
			});
			if(values != ""){
				values = values.substring(0,values.length-1);
				cAnswer = values;
			}
		}else if (types == 4){
			var values = "";
			$("input[name='answer44']:checked").each(function(){ 
				values+=$(this).val()+","; 
			});
			if(values != ""){
				values = values.substring(0,values.length-1);
				cAnswer = values;
			}
		}else if (types == 5){
			var values = $('input[name="answer55"]:checked').val();
			if(values != undefined){
				cAnswer = values;
			}
		}
		return cAnswer;
	}
</script>
<link href="../css/index.css" rel="stylesheet" type="text/css" />
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
    <%
						Set rs = rsSubjectLibrary5(papaerId)
					%>
    <%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=1
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
										mTypes =rs("mTypes")
										mAnswer = rs("Answer")
								%>
    <div class="ksmaindiv">
      <div class="examone" id="main_M">
        <div class="ajshow" id="frm_main" style="display: block;">
          <input type="hidden" id="sid" value="<%=rs("GW_SubjectLibrary.ID")%>"/>
          <div class="title"> <span id="t"><%=rs("GW_SubjectLibrary.Title")%></span> </div>
          <div class="item">
            <div class="left"> <%=rs("Options")%><span id="rightanswer" style="display: none;"></span> </div>
          </div>
          <%
											rs.movenext
											If rs.EOF Then
												Exit For
											End If
					  					next
									End If
								%>
        </div>
        <div class="tieba">
        	<input type="hidden" id="types" value="<%=mTypes%>" />
            <input type="hidden" id="answer" value="<%=mAnswer%>" />
            <input type="hidden" id="id" value="<%=id%>" />
            <input type="hidden" id="newAnswer" value="<%=oldAnswer%>" />
        	<div style="width:612px; margin:0 auto;">
          <%
			if mTypes = 1 then
			%>
          <div name="answer1" style="margin-top:5px;" id="answer1"> ѡ�
            <label>
              <input name="answer11" type="radio"  value="A">
              A </label>
            <label>
              <input type="radio" name="answer11"  value="B">
              B</label>
            <label>
              <input type="radio" name="answer11"  value="C">
              C</label>
            <label>
              <input type="radio" name="answer11"  value="D">
              D</label>
          </div>
          <%
			elseif mTypes = 2 then
			%>
          <div name="answer2" style="margin-top:5px;"  id="answer2"> ѡ�
            <label>
              <input type="checkbox" name="answer22"  value="A">
              A</label>
            <label>
              <input type="checkbox" name="answer22" value="B">
              B</label>
            <label>
              <input type="checkbox" name="answer22" value="C">
              C </label>
            <label>
              <input type="checkbox" name="answer22" value="D">
              D</label>
          </div>
          <%
			elseif mTypes = 3 then
			%>
          <div name="answer3" style="margin-top:5px;" id="answer3"> ѡ�
            <label>
              <input type="checkbox" name="answer33"  value="A">
              A </label>
            <label>
              <input type="checkbox" name="answer33" value="B">
              B </label>
            <label>
              <input type="checkbox" name="answer33" value="C">
              C </label>
            <label>
              <input type="checkbox" name="answer33" value="D">
              D</label>
            <label>
              <input type="checkbox" name="answer33" value="E">
              E </label>
          </div>
          <%
			elseif mTypes = 4 then
			%>
          <div name="answer4" style=" margin-top:5px;" id="answer4"> ѡ�
            <label>
              <input type="checkbox" name="answer44"  value="A">
              A</label>
            <label>
              <input type="checkbox" name="answer44" value="B">
              B</label>
            <label>
              <input type="checkbox" name="answer44" value="C">
              C</label>
            <label>
              <input type="checkbox" name="answer44" value="D">
              D </label>
            <label>
              <input type="checkbox" name="answer44" value="E">
              E</label>
            <label>
              <input type="checkbox" name="answer44" value="F">
              F</label>
          </div>
          <%
			elseif mTypes = 5 then
			%>
          <div name="answer5" style="margin-top:5px;"  id="answer5"> ѡ�
            <label>
              <input name="answer55" type="radio" value="1">
              ��</label>
            <label>
              <input type="radio" name="answer55" value="0">
              ��</label>
          </div>
          <%
			end if
		%>
         <div id="panel" style="margin-top:25px;">
          <%
											if cint(rs.pagecount)>cint(pageno) then
										%>
          <a  class="inp" onclick="next(<%=pageno+1%>,<%=papaerId%>);">��һ��</a>
          <%
											else
										%>
          <%
											end if
										%>
          <a  class="inp" onclick="toSubmit(<%=papaerId%>);">����</a> 
         	<span id="nochoose" style="width:auto; display:inline-block; text-align:center; background:#f00;font-family:'΢���ź�'; font-size:12px; color:wheat;display:none;">��ѡ��𰸣�</span>
            <span id="wrong" style="width:auto; display:inline-block; text-align:center; background:#f00;font-family:'΢���ź�'; font-size:12px; color:wheat;display:none;"></span>
         </div>
          <div style="text-align: -webkit-right;line-height:20px; font-size:12px;"> <font>��ǰ��</font> <font color="#FF0000"><%=pageno%></font> <font class="font_navigation">��</font> <font class="font_navigation">��</font> <font color="#FF0000"><%=rs.recordcount%></font> <font class="font_navigation">��</font> 
          </div>
          </div>
       </div>
        
      </div>
    </div>
  </div>
</div>
<!--#include file="../foot.asp" -->
</body>
</html>