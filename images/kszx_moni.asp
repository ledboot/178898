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
	</head>
	<body <%=bgcolor%><%=selectmouse%><%=rightmouse%> style="margin:0px; padding:0px; height:100%">
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
                    <div class="ksflbox2">
                    	<div class="ksfl_ltit">英语四级所有试卷</div>
                        <div class="ksclass_lbox">
                        	<table border="0" cellpadding="0" cellspacing="0" width="99%">
					<tr> 
						<td bgcolor="#F7F7F7" valign="top"> 
							<table  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
								<tr height="25"> 
									<td width="50%" background="images/admin_top_bg.jpg" class="ksname">试卷名称</td>
                                    <td width="15%" background="images/admin_top_bg.jpg" class="kstitle">发布时间</td>
                                    <td width="15%" background="images/admin_top_bg.jpg" class="kstitle">考试时间</td>
									<td width="23%" background="images/admin_top_bg.jpg" class="kstitle">操作选项</td>
								</tr>
                                <%
								  Set rs = rsPaper3(Session("ProfessionId"))
								%>
								<%
									pageno=cint(request.querystring("pageno"))
									if pageno=0 then pageno=1
									rs.pagesize=20
									rs.absolutepage=pageno
									if not rs.bof and not rs.eof then
										for i=1 to rs.pagesize
								%>
								<tr height="25" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
									<td width="50%" align="left"><img src="/mnkc/images/dotdb.gif" border="0"><%=rs("Title")%></td>
                                    <td width="15%" align="center"><%=Format_Time(rs("Date"), 2)%></td>
									<td width="15%" align="center"><%=rs("ExamTime")%>分钟</td>
									<td width="23%" align="center"><a href="kszx_start_moni.asp?flag=man&papaerId=<%=rs("ID")%>">进入考场</a>
									</td>
								</tr>
								<%
											rs.movenext
											If rs.EOF Then
												Exit For
											End If
					  					next
									End If
								%>
								<tr height="25" > 
									<td colspan="7" align="center">
										<a href="kszx_moni.asp?pageno=1">首页</a> 
										<%
											if pageno>1 then
										%>
										<a href="kszx_moni.asp?pageno=<%=pageno-1%>">上一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">上一页</font>
										<%
											end if
										%>
										<%
											if cint(rs.pagecount)>cint(pageno) then
										%>
										<a href="kszx_moni.asp?pageno=<%=pageno+1%>">下一页</a> 
										<%
											else
										%>
										<font color="#C0C0C0">下一页</font>
										<%
											end if
										%>
										<a href="kszx_moni.asp?pageno=<%=rs.pagecount%>">尾页</a>
										&nbsp;<font class="font_navigation">页次：</font>
										<font color="#FF0000"><%=pageno%></font>
										<font class="font_navigation">/</font>
										<font color="#FF0000"><%=rs.pagecount%></font>
										<font class="font_navigation">页</font>
										&nbsp;<font color="#FF0000"><%=rs.pagesize %></font>
										<font class="font_navigation">条/页 共</font>
									  <font color="#FF0000"><%=rs.recordcount%></font>
										<font class="font_navigation">条</font>
								  </td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
                        </div>
                    </div> 
                    
         		</div>
			</div>
		</div>
		<!--#include file="../foot.asp" -->
	</body>
</html>