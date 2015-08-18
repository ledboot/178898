<!--#include file="../inc/conn.asp"-->
<!--#include file="Admin_Config.asp"--> 
<%
	if session("AdminName") = "" then
		response.redirect "logout.asp"
	end if
%>
<%
	flag = Request("flag")
	zt = Request("zt")
	yjsx = Request.Form("yjsx")
%>
<%
	classid = Request("classid")
	classid1 = Request("Class_ID")
	sclass = Request("sclass")
	xclass = Request("xclass")
%>
<%
	If zt = "UpOrder" Then
		If isInteger(yjsx) = False Then
			Error1 "对不起，你填入的不是数字！请重新填写。", "Admin_Class_Article.asp"
		End If
		Set rs = rsClass3(classid1)
		rs("Class_sx") = Request.Form("yjsx")
		rs.Update
		rs.Close
		Set rs =Nothing
		Response.Redirect("Admin_Class_Article.asp")
		Response.End()
	End If
	If zt = "del" Then
		Call rsDelClass(classid)
		Response.Redirect("Admin_Class_Article.asp")
		Response.End()
	End If
%>
<script language="javascript" type="text/javascript">
	<!--
		function checkForm(tempFprm){
			if(tempFprm.Classname.value == ""){
				alert("对不起，类别名称不能为空！");
				tempFprm.Classname.focus();
				return false;
			}
			var tempValue = GetRadioValue("linkradio");
			if(tempValue == 1){
				if(tempFprm.Classlink.value == ""){
					alert("外部类的链接地址不能为空！");
					tempFprm.Classlink.focus();
					return false;
				}
			} 
		}
		function checkForm2(tempFprm){
			if(tempFprm.Classname.value == ""){
				alert("对不起，类别名称不能为空！");
				tempFprm.Classname.focus();
				return false;
			}
			var tempValue = GetRadioValue("linkradio");
			if(tempValue == 1){
				if(tempFprm.Classlink.value == ""){
					alert("外部类的链接地址不能为空！");
					tempFprm.Classlink.focus();
					return false;
				}
			}
			else{
				if(tempFprm.Titlename.value == ""){
					alert("此域不能为空！");
					tempFprm.Titlename.focus();
					return false;
				}
				if(tempFprm.Subheadname.value == ""){
					alert("此域不能为空！");
					tempFprm.Subheadname.focus();
					return false;
				}
				if(tempFprm.Authorname.value == ""){
					alert("此域不能为空！");
					tempFprm.Authorname.focus();
					return false;
				}
				if(tempFprm.CopyFromname.value == ""){
					alert("此域不能为空！");
					tempFprm.CopyFromname.focus();
					return false;
				}
				if(tempFprm.Gjzname.value == ""){
					alert("此域不能为空！");
					tempFprm.Gjzname.focus();
					return false;
				}
			}
		}
	-->
</script>
<%
	Dim ScClassid, Classname, ClassNameD, ClassLink,XcClassid
	If zt = "save" Then
		ScClassid = Request.Form("ClassID")
		Classname = Request.Form("Classname")
		If ScClassid <> 0 Then
			ClassNameD = Request.Form("linkradio")
			If ClassNameD = 1 Then
				ClassLink = Request.Form("Classlink")
			End If
		End If
		If ScClassid <> 0 Then
			Set rs = rsClass5(ScClassid)
			If rs("Class_num") <> 0 Then
				Set rs1 = rsClass6(ScClassid)
				XcClassid = CLng(rs1("XcClassid")) + 1
				rs1.Close
				Set rs1 = Nothing
			Else
				XcClassid = ScClassid & "100"
			End If
			rs("Class_num") = CInt(rs("Class_num")) + 1
			rs.Update
			rs.Close
			Set rs = Nothing
		Else
			Set rs1 = rsClass6(0)
			XcClassid = CLng(rs1("XcClassid")) + 1
			rs1.Close
			Set rs1 = Nothing
		End If
		Set rs = rsClass()
		rs.AddNew
		rs("ScClassid") = ScClassid
		rs("XcClassid") = XcClassid
		rs("ClassName") = ClassName
		If ScClassid <> 0 Then
			rs("ClassNameD") = ClassNameD
			If ClassNameD = 1 Then
				rs("Classlink") = Classlink
			End If
		End If
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("Admin_Class_Article.asp")
		Response.End()
	End If
%>
<%
	If zt = "edit" Then
		Dim TitleName, Subheadname, Authorname, CopyFromname, Gjzname,TitleName_1, Subheadname_1, Authorname_1, CopyFromname_1, Gjzname_1
		Classname = Request.Form("Classname")
		ClassNameD = Request.Form("linkradio")
		If ClassNameD = 1 Then
			ClassLink = Request.Form("Classlink")
		Else
			TitleName = Request.Form("TitleName")
			Subheadname = Request.Form("Subheadname")
			Authorname = Request.Form("Authorname")
			CopyFromname = Request.Form("CopyFromname")
			Gjzname = Request.Form("Gjzname")
			TitleName_1 = Request.Form("TitleName_1")
			If IsEmpty(TitleName_1) Then
				TitleName_1 = 0
			End If
			Subheadname_1 = Request.Form("Subheadname_1")
			If IsEmpty(Subheadname_1) Then
				Subheadname_1 = 0
			End If
			Authorname_1 = Request.Form("Authorname_1")
			If IsEmpty(Authorname_1) Then
				Authorname_1 = 0
			End If
			CopyFromname_1 = Request.Form("CopyFromname_1")
			If IsEmpty(CopyFromname_1) Then
				CopyFromname_1 = 0
			End If
			Gjzname_1 = Request.Form("Gjzname_1")
			If IsEmpty(Gjzname_1) Then
				Gjzname_1 = 0
			End If
		End If
		class_id2 = Request.Form("Classid_2")
		Set rs = rsClass2(class_id2)
		Response.Write(rs.RecordCount)
		rs("ClassName") = ClassName
		rs("ClassNameD") = ClassNameD
		If CInt(ClassNameD) = 1 Then
			rs("ClassLink") = ClassLink
		Else
			rs("TitleName") = TitleName
			rs("Subheadname") = Subheadname
			rs("Authorname") = Authorname
			rs("CopyFromname") = CopyFromname
			rs("Gjzname") = Gjzname
			rs("TitleName_1") = TitleName_1
			rs("Subheadname_1") = Subheadname_1
			rs("Authorname_1") = Authorname_1
			rs("CopyFromname_1") = CopyFromname_1
			rs("Gjzname_1") = Gjzname_1
		End If
		rs.Update
		Set rs = Nothing
		Response.Redirect("Admin_Class_Article.asp")
		Response.End()
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="css/index.css" type=text/css rel=stylesheet>
		<title>内容区</title>
		<script language="javascript" type="text/javascript">
			<!--
				function GetRadioValue(RadioName){
					var obj;    
					obj=document.getElementsByName(RadioName);
					if(obj!=null){
						var i;
						for(i=0;i<obj.length;i++){
							if(obj[i].checked){
								return obj[i].value;            
							}
						}
					}
					return null;
				}
				function showForm(){
					var classidObj = document.getElementById("classid");
					if(classidObj.options[classidObj.selectedIndex].value == 0){
						if(GetRadioValue("linkradio") == 1){
							var obj;    
							obj=document.getElementsByName("linkradio");
							obj[1].checked = 'checked';
							showlinkstr();
						}
						if(document.getElementById("shuxingshezhi") != null){
							document.getElementById("shuxingshezhi").style.display = "none";
						}
						document.getElementById("Classnamedstr").style.display = "none";
					}
					else{
						document.getElementById("Classnamedstr").style.display = "";
						if(document.getElementById("shuxingshezhi") != null){
							document.getElementById("shuxingshezhi").style.display = "";
						}
					}
				}
				function showlinkstr(){
					if(GetRadioValue("linkradio") == 1){
						document.getElementById("linkstr").style.display = "";
						if(document.getElementById("shuxingshezhi") != null){
							document.getElementById("shuxingshezhi").style.display = "none";
						}
					}
					else{
						document.getElementById("linkstr").style.display = "none";
						if(document.getElementById("shuxingshezhi") != null){
							document.getElementById("shuxingshezhi").style.display = "";
						}
					}
				}
			-->
		</script>
	</head>
	<body topmargin="0" leftmargin="0">
  		<center>
  			<table border="0" cellpadding="0" cellspacing="0" width="99%">
    			<tr>
					<td>&nbsp;</td>
    			</tr>
    			<tr height="25"> 
      				<td align="center" bgcolor="#E3E3E3" style="border-left: 1 solid #C0C0C0;border-top: 1 solid #C0C0C0;border-right: 1 solid #C0C0C0;" background="images/admin_top_bg.jpg"><b>栏 目 管 理</b></td>
    			</tr>
				<tr height="30"> 
				  <td bgcolor="#F7F7F7" style="border: 1 solid #C0C0C0;"><b>&nbsp;管理导航：</b><a href="Admin_Class_Article.asp">栏目管理</a> | <a href="Admin_Class_Article.asp?Flag=add"> 添加栏目</a> |&nbsp;<a href="javascript:history.back()">返回</a></td>
				</tr>
  			</table>
			<%
				If flag = "add" Then
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr height="25"> 
					<td align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-top: 1 solid #C0C0C0"><b>添 加 栏 目</b></td>
				</tr>
				<form method="POST" action="Admin_Class_Article.asp?zt=save" name="form1" onSubmit="return checkForm(this)">
					<tr> 
						<td bgcolor="#F7F7F7" height="25"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#484848" bordercolordark="#E3E3E3">
								<tr> 
									<td width="35%" height="38">&nbsp;<strong>所属栏目：</strong> </td>
									<td width="75%" height="38">&nbsp;
									  <select size="1" name="ClassID" onChange="showForm()">
											<option value="0">作为一级栏目</option>
											<%
												Set rs = rsClass1(0)
												While Not rs.EOF And Not rs.BOF
													If rs("ClassNameD") = 0 Then
											%>
											<option value="<%=rs("XcClassid")%>" <%If CLng(xclass) = CLng(rs("XcClassid")) Then%>selected<%End If%>><%=rs("ClassName")%></option>
											<%
													End If
													If CInt(rs("Class_num")) <> 0 Then
														Set rs1 = rsClass1(rs("XcClassid"))
														While Not rs1.EOF And Not rs1.BOF
															If rs1("ClassNameD") = 0 Then
											%>
											<option value="<%=rs1("XcClassid")%>" <%If CLng(xclass) = CLng(rs1("XcClassid")) Then%>selected<%End If%>>&nbsp;&nbsp;|--<%=rs1("ClassName")%></option>
											<%
															End If
											%>
											<%
															rs1.MoveNext
														Wend
														rs1.Close
														Set rs1 = Nothing
													End If
											%>
											<%
													rs.MoveNext
												Wend
												rs.Close
												Set rs = Nothing
											%>
									  </select>								  </td>
								</tr>
								<tr> 
									<td width="35%" height="37"><strong>&nbsp;栏目名称：<font color="#FF0000">(*)</font></strong></td>
									<td width="75%" height="37">&nbsp;
								  <input type="text" class="input1" name="Classname" size="20" style="width: 301; height: 18"></td>
								</tr>
								<tr id="Classnamedstr" style='display:none'> 
									<td width="35%" height="30">&nbsp;<strong>栏目类别：</strong></td>
									<td width="75%" height="30">&nbsp;
								  <input type="radio" name="linkradio" value="1" onClick="showlinkstr()">外部 <input name="linkradio" type="radio" value="0" checked onClick="showlinkstr()">内部</td>
								</tr>
								<script language="javascript" type="text/javascript">
									<!--
										var classidObj = document.getElementById("classid");
										if(classidObj.options[classidObj.selectedIndex].value != 0){
											document.getElementById("Classnamedstr").style.display = "";
										}
									-->
								</script>
								<tr id="linkstr" style='display:none'> 
									<td width="35%" height="37">&nbsp;<strong>栏目链接地址：<font color="#FF0000">(*)</font></strong></td>
									<td width="75%" height="37">&nbsp;
								  <input type="text" class="input1" name="Classlink" size="20" style="width: 301; height: 18">&nbsp;&nbsp;填相对地址</td>
								</tr>
								<tr align="center"> 
									<td width="100%" height="37" colspan="2"><input type="submit" value=" 添 加 " name="B1"> <input type="reset" value=" 取 消" name="B2"> </td>
								</tr>
							</table>
						</td>
					</tr>
				</form>
				<tr height="25"> 
					<td align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0">&nbsp;</td>
				</tr>
			</table>
			<%
				ElseIf flag = "edit" Then
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td height="25" align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-top: 1 solid #C0C0C0"><b>&nbsp;修 改 栏 目</b></td>
				</tr>
				<form method="POST" action="Admin_Class_Article.asp?zt=edit" name="form2" onSubmit="return checkForm2(this)">
					<%
						Set rs = rsClass2(classid)
					%>
					<tr> 
						<td bgcolor="#F7F7F7" height="25" valign="top"> 
							<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#484848" bordercolordark="#E3E3E3">
								<tr height="37"> 
									<td width="25%"><strong>&nbsp;栏目名称：<font color="#FF0000">(*)</font></strong></td>
									<td width="75%">&nbsp;
									  <input type="text" class="input1" name="Classname" size="20" style="width: 301; height: 18" value="<%=rs("ClassName")%>">
								  <input name="Classid_2" type="hidden" value="<%=rs("Class_ID")%>"></td>
								</tr>
								<tr height="37" id="Classnamedstr" <%If rs("ScClassid") = 0 Then%>style='display:none'<%End If%>> 
									<td width="25%">&nbsp;<strong>栏目类别：</strong></td>
									<td width="75%">&nbsp;
									  <input name="linkradio" type="radio" onClick="showlinkstr();" value="1" <%If rs("ClassNameD") = 1 Then%>checked<%End If%>>外部
								  <input name="linkradio" type="radio" value="0" <%If rs("ClassNameD") = 0 Then%>checked<%End If%> onClick="showlinkstr();">内部								  </td>
								</tr>
								<tr height="37" id="linkstr" <%If rs("ClassNameD") = 0 Then%>style='display:none'<%End If%>> 
									<td width="25%">&nbsp;<strong>栏目链接地址：<font color="#FF0000">(*)</font></strong></td>
									<td width="75%">&nbsp;&nbsp;
								  <input type="text" class="input1" name="Classlink" size="20" style="width: 301; height: 18" value="<%=rs("Classlink")%>"></td>
								</tr>
								<tr height="37" id="shuxingshezhi" <%If rs("ClassNameD") = 1 Or rs("ScClassid") = 0 Then%>style='display:none'<%End If%>>
									<td width="25%">&nbsp;<strong>属性设置：</strong></td>
									<td width="75%">
										&nbsp;
										<input type="text" class="input1" name="Titlename" size="10" value="<%=rs("TitleName")%>">
										<input type="checkbox" name="Titlename_1" value="1" checked> 
										| &nbsp;<input type="text" class="input1" name="Subheadname" size="10" value="<%=rs("Subheadname")%>">
										<input type="checkbox" name="Subheadname_1" value="1" <%If rs("Subheadname_1") <> 0 Then%>checked<%End If%>> 
										|&nbsp;<input type="text" class="input1" name="Authorname" size="10" value="<%=rs("Authorname")%>">
										<input type="checkbox" name="Authorname_1" value="1" <%If rs("Authorname_1") <> 0 Then%>checked<%End If%>> 
										|&nbsp;<input type="text" class="input1" name="CopyFromname" size="10" value="<%=rs("CopyFromname")%>">
										<input type="checkbox" name="CopyFromname_1" value="1" <%If rs("CopyFromname_1") <> 0 Then%>checked<%End If%>> 
										|&nbsp;<input type="text" class="input1" name="Gjzname" size="10" value="<%=rs("Gjzname")%>">
								  <input type="checkbox" name="Gjzname_1" value="1" <%If rs("Gjzname_1") <> 0 Then%>checked<%End If%>></td>
								</tr>
								<tr align="center"> 
									<td height="37" colspan="2"><input type="submit" value=" 修 改 " name="B1"> <input type="reset" value=" 取 消 " name="B2"></td>
								</tr>
							</table>
						</td>
					</tr>
				</form>
				<tr height="25"> 
					<td align="center" background="images/admin_top_bg.jpg" style="border-left: 1 solid #C0C0C0; border-right: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0">&nbsp;</td>
				</tr>
			</table>
			<%
				Else
			%>
  			<table border="0" cellpadding="0" cellspacing="0" width="99%">
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td align="center" valign="top" bgcolor="#F7F7F7">
						<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#F7F7F7">
							<tr height="25"> 
								<td width="49%" background="images/admin_top_bg.jpg">&nbsp;栏目名称 -- 排序</td>
								<td width="26%" background="images/admin_top_bg.jpg">&nbsp;链接地址</td>
								<td width="25%" background="images/admin_top_bg.jpg">&nbsp;操作选项</td>
							</tr>
							<%
								Set rs = rsClass1(0)
								While Not rs.EOF And Not rs.BOF
							%>
		  					<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
								<form action="Admin_Class_Article.asp?zt=UpOrder" method="post">
									<td width='49%'>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <tr>
											<td width="25">&nbsp;<img src="images/tree_folder4.gif" valign="abvmiddle"></td>
											<td align="left"><font color="red"><%=rs("ClassName")%></font></td>
											<td width="65">&nbsp;
										    <input class="input1" type="text" name="yjsx" size="2"><input type="hidden" name="Class_ID" value="<%=rs("Class_ID")%>"></td>
											<td width="60">&nbsp;
										    <input class="input1" type="submit" name="Submit" value="修改"></td>
											<td width="100">&nbsp;当前排序：<font color="red"><%=rs("Class_sx")%></font></td>
										  </tr>
									  </table>
									</td>
								</form>
								<td width="26%">&nbsp;</td>
								<td width="25%" align="right">
									<%
										If rs("ClassName") <> "常规设置" Then
									%><a href="Admin_Class_Article.asp?Flag=add&xclass=<%=rs("XcClassid")%>">添加子栏目</a> 
									| <a href="Admin_Class_Article.asp?Flag=edit&classid=<%=rs("Class_ID")%>">修改设置</a> 
									| <a href="Admin_Class_Article.asp?zt=del&classid=<%=rs("Class_ID")%>">删除栏目</a><%
										Else
									%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%End If%>
								</td>
							</tr>
							<%
									If CInt(rs("Class_num")) <> 0 Then
										Set rs1 = rsClass1(rs("XcClassid"))
										While Not rs1.EOF And Not rs1.BOF
							%>
							<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
								<form action="Admin_Class_Article.asp?zt=UpOrder" method="post">
									<td width='49%'>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <tr>
											<td width="35">&nbsp; <font color="#FF0000">|--</font></td>
											<td width="25"><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"></td>
											<td align="left"><font color="#3366CC"><%=rs1("ClassName")%></font></td>
											<td width="65">&nbsp;
										    <input class="input1" type="text" name="yjsx" size="2" style='background-color: #F5F7FE'><input type="hidden" name="Class_ID" value="<%=rs1("Class_ID")%>"></td>
											<td width="60">&nbsp;
										    <input class="input1" type="submit" name="Submit" value="修改"></td>
											<td width="100">&nbsp;二级排序：<%=rs1("Class_sx")%></td>
										  </tr>
										</table>
									</td>
								</form>
								<td width="26%">&nbsp;
									<%
										If rs1("ClassNameD") = 1 Then
											Response.Write(rs1("ClassLink"))
										End If
									%>
								</td>
								<td width="25%" align="right">
									<%
										If rs1("ClassNameD") = 0 Then
									%>
									<a href="Admin_Class_Article.asp?Flag=add&xclass=<%=rs1("XcClassid")%>">添加子栏目</a>
									<%
										End If
									%><%
										If rs("ClassName") <> "常规设置" Then
									%> 
									| <a href="Admin_Class_Article.asp?Flag=edit&classid=<%=rs1("Class_ID")%>">修改设置</a> 
									| <a href="Admin_Class_Article.asp?zt=del&classid=<%=rs1("Class_ID")%>">删除栏目</a><%
										Else
									%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%End If%>
								</td>
							</tr>
							<%
											If rs1("Class_num") <> 0 Then
												Set rs2 = rsClass1(rs1("XcClassid"))
												While Not rs2.EOF And Not rs2.BOF
							%>
							<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
								<form action="Admin_Class_Article.asp?zt=UpOrder" method="post">
									<td width='49%'>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <tr>
											<td width="35"></td>
											<td width="15"><font color="#FF0000">|-</font></td>
											<td width="25"><img src="images/tree_folder3.gif" valign="abvmiddle" width="15" height="15"></td>
											<td align="left"><font color="#006600"><%=rs2("ClassName")%></font></td>
											<td width="65">&nbsp;
										    <input class="input1" type="text" name="yjsx" size="4" style='background-color: #F5F7FE'><input type="hidden" name="Class_ID" value="<%=rs2("Class_ID")%>"></td>
											<td width="60">&nbsp;
										    <input class="input1" type="submit" name="Submit" value="修改"></td>
											<td width="100">&nbsp;三级排序：<font color=#009900><%=rs2("Class_sx")%></font></td>
										  </tr>
										</table>
									</td>
								</form>
								<td width="26%">&nbsp;
									<%
										If rs2("ClassNameD") = 1 Then
											Response.Write(rs2("ClassLink"))
										End If
									%>
								</td>
								<td width="25%" align="right">
									<%
										If rs("ClassName") <> "常规设置" Then
									%>&nbsp;| <a href="Admin_Class_Article.asp?Flag=edit&classid=<%=rs2("Class_ID")%>">修改设置</a> 
									| <a href="Admin_Class_Article.asp?zt=del&classid=<%=rs2("Class_ID")%>">删除栏目</a><%
										Else
									%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%End If%>
								</td>
							</tr>
							<%
													rs2.MoveNext
												Wend
												rs2.Close
												Set rs2 = Nothing
											End If
							%>
							<%
											rs1.MoveNext
										Wend
										rs1.Close
										Set rs1 = Nothing
									End If
							%>
							<%
									rs.MoveNext
								Wend
								rs.Close
								Set rs = Nothing
							%>
						</table>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>
			<%
				End If
			%>
		</center>
	</body>
</html>
<%
	call CloseDatabase()
%>