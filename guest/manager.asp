<!--#include file="gsconn.asp"--> 
<%
if session("AdminName") = "" then
response.redirect "../login.asp"
end if
%>

<html>
<head>
<title>留言反馈管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="guestbook.css">
</head>

<body bgcolor="#FFFFFF" background="gs_images/bg.gif">
<div align="center"> 
<%
set rs=server.createobject("adodb.recordset")
sql="select * from book order by id desc"
rs.open sql,conngs,1,1
dim thisPageSize
dim thisPageCount
dim thisPageCurrent
dim RecordsNumber
thisPageSize=5
if request("page")="" then
	thisPageCurrent=1
else
	thisPageCurrent=CInt(request("page"))
end if
rs.PageSize=thisPageSize
thisPageCount=rs.PageCount
if thisPageCurrent > thisPageCount then thisPagecurrent=thisPageCount
if thisPageCurrent <1 then thisPageCurrent=1
if thisPageCount=0 then
	response.write"暂时没有留言"
else
	rs.AbsolutePage=thisPageCurrent
%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="25"> 
        <%
response.write "页次：<strong><font color=red>"&thisPageCurrent&"</font>/"&thisPageCount&"</strong>页 "
response.write "&nbsp;共<font color=red><b>"&rs.recordcount&"</b></font>篇留言 <b>"&thisPageSize&"</b>篇留言/页 "
		  %>
      </td>
      <td height="25" width="100" align="right">&nbsp;</td>
    </tr>
  </table>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <%
		RecordsNumber=0
		do while RecordsNumber < thisPageSize and not rs.eof
	   %>
    <tr> 
      <td> 
        <table bordercolor=#ffffff cellspacing=0 cellpadding=0 width=100% 
  bordercolorlight=<%if rs("color")="cfffca" then response.write"80ff80" else if rs("color")="d9ecff" then response.write"acd6ff" else if rs("color")="ffe6f2" then response.write"ffacd6" else if rs("color")="FFFBC1" then response.write"FEED1B" else if rs("color")="EAEAEA" then response.write"D0D0D0" else if rs("color")="ECECFF" then response.write"D2D2FF" end if%> border=1>
          <tbody> 
          <tr> 
            <td width="100%" bgcolor=<%=rs("color")%>> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><font color=<%if rs("color")="cfffca" then response.write"12a402" else if rs("color")="d9ecff" then response.write"0080ff" else if rs("color")="ffe6f2" then response.write"ff80c0" else if rs("color")="FFFBC1" then response.write"FFA042" else if rs("color")="EAEAEA" then response.write"808080" else if rs("color")="ECECFF" then response.write"AD5BFF" end if%>>留言者： 
                    <%=rs("name")%></font><font color=<%if rs("color")="cfffca" then response.write"12a402" else if rs("color")="d9ecff" then response.write"0080ff" else if rs("color")="ffe6f2" then response.write"ff80c0" else if rs("color")="FFFBC1" then response.write"FFA042" else if rs("color")="EAEAEA" then response.write"808080" else if rs("color")="ECECFF" then response.write"AD5BFF" end if%>>&nbsp;&nbsp;时间：<%=rs("time")%></font>
					<font color=<%if rs("color")="cfffca" then response.write"12a402" else if rs("color")="d9ecff" then response.write"0080ff" else if rs("color")="ffe6f2" then response.write"ff80c0" else if rs("color")="FFFBC1" then response.write"FFA042" else if rs("color")="EAEAEA" then response.write"808080" else if rs("color")="ECECFF" then response.write"AD5BFF" end if%>>&nbsp;&nbsp;联系电话： 
                    <%=rs("tel")%></font>
					</td>
                  <td align="right" width="15%"> 
                    <%if rs("oicq")<>"" then %>
                      <img src="gs_images/oicq.gif" alt="<%=rs("oicq")%>" width="16" height="16" border="0"> 
                      <%else%>
                    <img src="gs_images/oicq1.gif" border="0"> 
                    <%end if%>
                    <%if rs("homepage")<>"" and rs("homepage")<>"http://" then %>
                    <a href="<%=rs("homepage")%>" target="_blank"><img src="gs_images/web.gif" alt="<%=rs("homepage")%>" border="0"></a> 
                    <%else%>
                    <img src="gs_images/noweb.gif" border="0"> 
                    <%end if%>
                    <%if rs("email")<>"" then %>
                    <a href="mailto:<%=rs("email")%>"><img src="gs_images/mail.gif" alt="<%=rs("email")%>" border="0"></a> 
                    <%else%>
                    <img src="gs_images/nomail.gif" border="0"> 
                    <%end if%>
                    <img src="gs_images/ip.gif" alt=<%=rs("ip")%> width="13" height="15" border="0"></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr bgcolor="<%=rs("color")%>"> 
            <td width="100%"> 
              <table cellspacing=0 cellpadding=1 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="10%" align="center"> <font color="<%if rs("color")="cfffca" then response.write"12a402" else if rs("color")="d9ecff" then response.write"0080ff" else if rs("color")="ffe6f2" then response.write"ff80c0" else if rs("color")="FFFBC1" then response.write"FFA042" else if rs("color")="EAEAEA" then response.write"808080" else if rs("color")="ECECFF" then response.write"AD5BFF" end if%>">留言内容</font></td>
                  <td width="90%"> 
                    <table bordercolor=#ffffff cellspacing=0 bordercolordark=#c0c0c0 
              cellpadding=2 width="100%" bgcolor=#ffffff border=1>
                      <tbody> 
                      <tr> 
                        <td valign=top width="100%"> 
                          <table width="98%" border="0" cellspacing="0" cellpadding="2" align="center" height="70">
                            <tr align="left" valign="top"> 
                              <td>&nbsp;<img border="0" src="<%=rs("pic")%>">&nbsp;<font color=#ffb312><%=rs("content")%></font></td>
                            </tr>
                            <tr> 
                              <td> 
                                <div align=center> 
                                  <center>
                                  </center>
                                </div>
                                <div align="left"> 
                                  <%if rs("answer")<>"" then%>
                                  <br>
                                  <font color="#ff80c0"><%=rs("webmaster")%>回复<b>:</b></font><font color=0099ff> 
                                  <%=rs("answer")%> 
                                  <%end if%>
                                  </font></div>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      </tbody> 
                    </table>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr bgcolor="<%=rs("color")%>">
              <td width="100%" align="center"> <img src="gs_images/write.gif" width="28" height="28" align=absmiddle border="0"><a href="answer.asp?id=<%=rs("id")%>">回复</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="gs_images/manage.gif" width="28" height="28" align=absmiddle><a href="edit.asp?id=<%=rs("id")%>">修改留言</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="gs_images/bwz.gif" width="28" height="28" align="absmiddle"><a href="delete.asp?id=<%=rs("id")%>">删除留言</a> 
                <%
	if rs("shenhe") then %>
            <a href='shenhe.asp?action=guan&id=<%= cstr(rs("id")) %>&page=<%=request("page")%>'><font color="#009900">已开通</font></a> 
            <%else %>
            <a href='shenhe.asp?action=kai&id=<%= cstr(rs("id")) %>&page=<%=request("page")%>'><font color="#009900">未开通</font></a> 
            <%	end if %></td>
          </tr>
          </tbody> 
        </table>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <%
		RecordsNumber=RecordsNumber+1
		rs.movenext
		loop
		%>
  </table>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="right"> 
        <form method=Post action=manager.asp?>
          <% if thisPageCurrent <>1 then %>
          <a href="manager.asp?page=1">首页</a> 
          <%end if %>
          <% if thisPageCurrent <>1 then %>
          <a href="manager.asp?page=<%=thisPageCurrent-1%>">上一页</a> 
          <%end if %>
          <% if thisPageCurrent < thisPageCount then %>
          <a href="manager.asp?page=<% =thisPageCurrent+1 %>">下一页</a> 
          <%end if %>
          <% if thisPageCurrent < thisPageCount then %>
          <a href="manager.asp?page=<% =thisPageCount %>">末页</a> 
          <%end if   
						%>
          <font color='#000080'>转到：</font> 
          <input type='text' name='page' size=2 value=<%=currentpage%>>
          <input class=buttonface type='submit'  value='Goto'  name='cndok'>
        </form>
        <%end if %>
      </td>
    </tr>
  </table>
  
</div>
</body>
</html>
 