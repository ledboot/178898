<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
	flag = SafeRequest("flag", 0)
	papaerId  = SafeRequest("papaerId", 0)
	quesId = SafeRequest("quesId", 0)
	tmtype = SafeRequest("tmtype", 1)
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>添加信息</title>
        <script src="../js/jquery-1.4.2.min.js"></script>
         
        <script type="text/javascript">
			$(document).ready(function(){
				$("#tmtype").change(function(){
					var typeId = $("#tmtype option:selected").val();
					for(var i =1;i<6;i++){
						if(i == typeId){
							$("#answer"+typeId).show();
						}else{
							$("#answer"+i).hide();
						}
					}
					
				});
			});
        </script>

	</head>
	<body >
		<center>
			<%
				If flag = "save" Then
					quesId = Request.Form("quesId")
					PaperId = Request.Form("paperId")
					Title = Request.Form("contentTitle")
					Options = Request.Form("content")
					Score = Request.Form("Score")
					Answer =""
					Types = Request.Form("tmtype")
					if Types = 1 then
						Answer = replace(Request.Form("answer11")," ","")
					elseif Types = 2 then
						Answer = replace(Request.Form("answer22")," ","")
					elseif Types = 3 then
						Answer = replace(Request.Form("answer33")," ","")
					elseif Types = 4 then
						Answer = replace(Request.Form("answer44")," ","")
					elseif Types = 5 then
						Answer = replace(Request.Form("answer55")," ","")
					end if
					if Answer = "" then
						Error1 "对不起，请选择题目答案!", "Admin_Questions_List.asp?flag=add&papaerId="&PaperId
					end if
					Set rs1 = rsSubjectLibrary4(quesId)
					rs1("PaperId") = PaperId
					rs1("Title") = Title
					rs1("Options") = Options
					rs1("Answer") = Answer
					rs1("Types") = Types
					rs1("Score") = Score
					rs1.Update
					
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="65%" align="center" bordercolorlight="#000000" bordercolordark="#F7F7F7">
				<tr> 
					<td align="center" height="25" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;" background="images/admin_top_bg.jpg"><b>试题修改成功</b></td>
				</tr>
				<tr>
					<td align="center" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0"> 
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr height="25"> 
								<td width="20%" align="right"><b>试题名称：</b></td>
								<td width="80%" align="left"><font color="#800000"><%=Title%></font></td>
							</tr>
                            <tr height="25"> 
								<td width="20%" align="right"><b>选项：</b></td>
								<td width="80%" align="left"><font color="#800000"><%=Options%></font></td>
							</tr>
                            <tr height="25"> 
								<td width="20%" align="right"><b>答案：</b></td>
								<td width="80%" align="left"><font color="#800000"><%=Answer%></font></td>
							</tr>
                            <tr height="25"> 
								<td width="20%" align="right"><b>题目分值：</b></td>
								<td width="80%" align="left"><font color="#800000"><%=Score%></font></td>
							</tr>
							<tr>
								<td align="center" colspan="2" height="30"></a> <font color="#800000"></font>  <font color="#800000">◇</font> <a href="Admin_Questions_List.asp?flag=man&papaerId=<%=PaperId%>">返回</a></p></td>
							</tr>
						</table>
					</td>
					<%
						rs1.Close
						Set rs1 = Nothing
					%>
				</tr>
				<tr> 
					<td>&nbsp;</td>
				</tr>
			</table>
            
			<%
				
				elseif flag = "update" then
				set rs3 = rsSubjectLibrary2(quesId)
			%>
            <script language="vbscript">
			sub update()
				if form1.tmtype.value = "" then
					msgbox "请选择题目类型！",48,"注意！"
					form1.tmtype.focus()
					exit sub
				end if
				if form1.Score.value = "" then
					msgbox "请输入题目分数！",48,"注意！"
					form1.Score.focus()
					exit sub
				end if
				if not isnumeric(form1.Score.value) then
					msgbox "题目分数请输入数字！",48,"注意！"
					form1.Score.focus()
					exit sub
				end if
				form1.submit()
			end sub
        	</script>
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr> 
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;">试卷管理 — 信息添加</td>
				</tr>
				<tr> 
					<td height="30" bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;管理导航：</b><a href="Admin_Questions_List.asp?flag=man&papaerId=<%=papaerId%>">管理该栏目下的信息</a> | <a href="Admin_Questions_List.asp?flag=man&papaerId=<%=papaerId%>">返回</a></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td height="25" style="border: 1 solid #C0C0C0" background="images/admin_top_bg.jpg">&nbsp;当前状态：试题管理&nbsp;<font color="#FF0000">添加信息</font></td>
				</tr>
				<tr> 
					<td align="center" bgcolor="#F7F7F7" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0"> 
						<form method="POST" action="?flag=save" name="form1">
							<table width="800px" border="1" align="center" cellpadding="0" cellspacing="0">
                            	<tr height="25">
                                    <th colspan="4">
                                    	<input type="hidden" name="paperId" value="<%=papaerId%>"/>
                                    	<input type="hidden" name="quesId" value="<%=quesId%>"/>
                                    </th>
                                </tr>
                            	<tr height="25">
                                	<td>试卷名称：</td>
                                    <th height="20" align="left" colspan="3"><%=rs3("GW_Paper.Title")%></th>
                                </tr>
								<tr height="25"> 
									<td>题目类型：</td>
                                    <td height="20" align="left" nowrap> <select name="tmtype" id="tmtype">
                    <option value="1"  selected>单项选择题</option>
                    <option value="2" <%if tmtype=2 then%> selected<%end if%>>多项选择题(A-D)</option>
                    <option value="3" <%if tmtype=3 then%> selected<%end if%>>多项选择题(A-E)</option>
                    <option value="4" <%if tmtype=4 then%> selected<%end if%>>多项选择题(A-F)</option>
                    <option value="5" <%if tmtype=5 then%> selected<%end if%>>判断题</option>
                  </select>
                  </td>
                   <td>题目分数：</td>
                                  <td><input type="text" name="Score" value="<%=rs3("Score")%>"/></td>
									
								</tr>
								<tr height="25"> 
									<td>题目标题：</td>
									<td colspan="3" align="left"><textarea name="contentTitle" style="display:none"><%=Server.HtmlEncode(rs3("GW_SubjectLibrary.Title"))%></textarea><IFRAME ID="eWebEditor1" SRC="Zzz4eweb/ewebeditor.htm?id=contentTitle&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="200"></IFRAME> <font color="#FF0000">*</font></td>
								</tr>
                                <tr height="25">
                                	<td>题目选项：</td>
                                    <th colspan="3" align="left">
                                    <textarea name="content" style="display:none"><%=Server.HtmlEncode(rs3("Options"))%></textarea>
                                    <IFRAME ID="eWebEditor2" SRC="Zzz4eweb/ewebeditor.htm?id=content&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME> <font color="#FF0000">*</font>
                                    </th>
                                </tr>
                                <tr>
                                	<td>参考答案：</td>
                                    <th colspan="3" align="left">
                                    	<div name="answer1" id="answer1" <%if tmtype<>1 then response.write " style='display:none'"%>>
                                    <label><input name="answer11" type="radio"  value="A" <%if instr(rs3("Answer"),"A")>0 then response.Write("checked")%>>
                                    A </label>
                                   <label><input type="radio" name="answer11"  value="B" <%if instr(rs3("Answer"),"B")>0 then response.Write("checked")%>>
                                    B</label> 
                                    <label><input type="radio" name="answer11"  value="C" <%if instr(rs3("Answer"),"C")>0 then response.Write("checked")%>>
                                    C</label> 
                                    <label><input type="radio" name="answer11"  value="D" <%if instr(rs3("Answer"),"D")>0 then response.Write("checked")%>>
                                    D</label> 
                    </div>
                  <div name="answer2" id="answer2" <%if tmtype<>2 then response.write " style='display:none'"%>> 
                    <label><input type="checkbox" name="answer22"  value="A" <%if instr(rs3("Answer"),"A")>0 then response.Write("checked")%>>
                    A</label> 
                    <label><input type="checkbox" name="answer22" value="B" <%if instr(rs3("Answer"),"B")>0 then response.Write("checked")%> >
                    B</label> 
                    <label><input type="checkbox" name="answer22" value="C" <%if instr(rs3("Answer"),"C")>0 then response.Write("checked")%>>
                    C </label>
                    <label><input type="checkbox" name="answer22" value="D" <%if instr(rs3("Answer"),"D")>0 then response.Write("checked")%>>
                    D</label> 
					
				  </div>
                  
                  <div name="answer3" id="answer3"<%if tmtype<>3 then response.write " style='display:none'"%>> 
                   <label> <input type="checkbox" name="answer33"  value="A" <%if instr(rs3("Answer"),"A")>0 then response.Write("checked")%>>
                    A </label>
                    <label><input type="checkbox" name="answer33" value="B" <%if instr(rs3("Answer"),"B")>0 then response.Write("checked")%>>
                    B </label>
                    <label><input type="checkbox" name="answer33" value="C" <%if instr(rs3("Answer"),"C")>0 then response.Write("checked")%>>
                    C </label>
                   <label><input type="checkbox" name="answer33" value="D" <%if instr(rs3("Answer"),"D")>0 then response.Write("checked")%>>
                    D</label> 
                   <label> <input type="checkbox" name="answer33" value="E" <%if instr(rs3("Answer"),"E")>0 then response.Write("checked")%>>
                    E </label>
					</div>
                    
					<div name="answer4" id="answer4"<%if tmtype<>4 then response.write " style='display:none'"%>> 
                    <label><input type="checkbox" name="answer44"  value="A" <%if instr(rs3("Answer"),"A")>0 then response.Write("checked")%>>
                    A</label> 
                    <label><input type="checkbox" name="answer44" value="B" <%if instr(rs3("Answer"),"B")>0 then response.Write("checked")%>>
                    B</label> 
                    <label><input type="checkbox" name="answer44" value="C" <%if instr(rs3("Answer"),"C")>0 then response.Write("checked")%>>
                    C</label> 
                    <label><input type="checkbox" name="answer44" value="D" <%if instr(rs3("Answer"),"D")>0 then response.Write("checked")%>>
                    D </label>
                    <label><input type="checkbox" name="answer44" value="E" <%if instr(rs3("Answer"),"E")>0 then response.Write("checked")%>>
                    E</label> 
                   <label> <input type="checkbox" name="answer44" value="F" <%if instr(rs3("Answer"),"F")>0 then response.Write("checked")%>>
                    F</label> 
					</div>
					
                    <div name="answer5" id="answer5"<%if tmtype<>5 then response.write " style='display:none'"%>> 
                   <label> <input name="answer55" type="radio" value="1" <%if rs3("Answer") = 1 then response.Write("checked")%>>
                    对</label> 
                    <label><input type="radio" name="answer55" value="0" <%if rs3("Answer") = 0 then response.Write("checked")%>>
                    错</label> </div>
                                    </th>
                                </tr>
								<tr height="25" align="center"> 
									<td colspan="4"><input type="button" value="修改" name="B1" onClick="update()"></td>
								</tr>
							</table>
						</form>
					</td>
				</tr>
			</table>
			<%
				rs3.close
				set rs3 = nothing
				End If
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>