<div class="gsjjText">
	<div class="OnlineMessageContent">
		<%
            set rs=server.createobject("adodb.recordset")
            sql="select * from book where shenhe order by id desc"
            rs.open sql, conn, 1, 1
			
			pageno=CInt(request("pageno"))
			If pageno = 0 Or IsEmpty(pageno) Or IsNull(pageno) Then
				pageno = 0
			End If
			RCount = 5
			if pageno=0 then
				pageno=1
			else
				pageno=CInt(request("pageno"))
			end if
			rs.PageSize=RCount
			RPC=CInt(rs.PageCount)
			if pageno > RPC then pageno=RPC
			if pageno < 1 then pageno=1
			rs.absolutepage = pageno
			RRC = CInt(rs.RecordCount)
        %>
        <div class="OnlineMessageNavigation"><span class="NavigationSpan">页次：<%=pageno%> / <%=RPC%> 页 共 <%=RRC%> 篇留言 <%=RCount%> 篇留言/页</span><span class="buttonSpan"><a href="zxly.asp?lx=1" target="_blank">我 要 留 言</a></span></div>
		<%
            if pageno=0 then
        %>
        <div class="OnlineMessageListNo">暂时没有留言</div>
		<%
            Else
                RecordsNumber=0
				Do While RecordsNumber < RCount and not rs.eof
					colorRs = rs("color")
					nameRs = rs("name")
					oicqRs = rs("oicq")
					homepageRs = rs("homepage")
					emailRs = rs("email")
					ipRs = rs("ip")
					picRs = rs("pic")
					contentRs = rs("content")
					timeRs = rs("time")
					answerRs = rs("answer")
					
					borderColorStr = ""
					if colorRs = "cfffca" then
						borderColorStr = "#80ff80"
					ElseIf colorRs = "d9ecff" then
						borderColorStr = "#acd6ff"
					ElseIf colorRs = "ffe6f2" then
						borderColorStr = "#ffacd6"
					ElseIf colorRs = "FFFBC1" then
						borderColorStr = "#FEED1B"
					ElseIf colorRs = "EAEAEA" then
						borderColorStr = "#D0D0D0"
					ElseIf colorRs = "ECECFF" then
						borderColorStr = "#D2D2FF"
					End If
					
        %>
        <div class="OnlineMessageClearance"></div>
        <div class="OnlineMessageListOuter" style="border-color:<%=borderColorStr%>; background-color:#<%=colorRs%>">
        	<%
					if colorRs = "cfffca" then
						FontColorStr = "#12a402"
					elseif colorRs = "d9ecff" then
						FontColorStr = "#0080ff"
					elseif colorRs = "ffe6f2" then
						FontColorStr = "#ff80c0"
					elseif colorRs = "FFFBC1" then
						FontColorStr = "#FFA042"
					elseif colorRs = "EAEAEA" then
						FontColorStr = "#808080"
					elseif colorRs = "ECECFF" then
						FontColorStr = "#AD5BFF"
					end if
			%>
        	<div class="OnlineMessageInformationDescription" style="border-bottom-color:<%=borderColorStr%>"><span class="textSpan" style="color:<%=FontColorStr%>">留言者：<%=nameRs%> &nbsp;时间：<%=Format_Time(timeRs,2)%></span><span class="imgSpan">
            	<%
					if Not checkIsEmpty(oicqRs) then
						oicqImgStr = "<img src=""gs_images/oicq.gif"" alt=""" & oicqRs & """ border=""0"">"
					else
						oicqImgStr = "<img src=""gs_images/oicq1.gif"" border=""0"">"
					end if
					Response.Write(oicqImgStr & "&nbsp;")
				%>
				<%
					if Not checkIsEmpty(homepageRs) and homepageRs <> "http://" then
						homepageAStr = "<a href=""" & homepageRs & """ target=""_blank""><img src=""gs_images/web.gif"" alt=""" & homepageRs & """ border=""0""></a>"
					Else
						homepageAStr = "<img src=""gs_images/noweb.gif"" border=""0"">"
					end if
					Response.Write(homepageAStr & "&nbsp;")
				%>
				<%
					if Not checkIsEmpty(emailRs) then
						emailAStr = "<a href=""mailto:" & emailRs & """><img src=""gs_images/mail.gif"" alt=""" & emailRs & """ border=""0""></a>"
					Else
						emailAStr = "<img src=""gs_images/nomail.gif"" border=""0"">"
					End If
					Response.Write(emailAStr & "&nbsp;")
				%><img src="gs_images/ip.gif" alt=<%=ipRs%> width="13" height="15" border="0">&nbsp;
            </span></div>
            <div class="OnlineMessageInformationContent">
            	<div class="InformationTitle" style="color:<%=FontColorStr%>">留言内容</div>
                <div class="InformationContent" style="border-left-color:<%=borderColorStr%>">
                	<div class="InformationContentBody" style="color:<%=FontColorStr%>"><img border="0" src="<%=picRs%>">&nbsp;<%=contentRs%></div>
                    <%
						if Not checkIsEmpty(answerRs) then
					%>
                    <div class="InformationContentReply" style="border-top-color:<%=borderColorStr%>"><span class="ReplyTitle">回复：</span><span class="ReplyContent"><%=answerRs%></span></div>
                    <%
						End If
					%>
                </div>
                <div style="clear:both"></div>
            </div>
        </div>
        <div class="OnlineMessageClearance"></div>
        <%
					RecordsNumber = RecordsNumber + 1
					rs.movenext
				Loop
            End If
			rs.close
			set rs=nothing
		%>
        <div class="SelectionZXLY">
			<%
				if pageno > 1 then
			%>
			<span><a href="<%=NavPage_1%>&pageno=1"><img src="images/Selection_03.jpg" /></a>&nbsp;</span>
			<span><a href="<%=NavPage_1%>&pageno=<%=pageno-1%>"><img src="images/Selection_05.jpg" /></a>&nbsp;</span>
			<%
				Else
			%>
			<span><img src="images/Selection_03.jpg" />&nbsp;</span>
			<span><img src="images/Selection_05.jpg" />&nbsp;</span>
			<%
				End If
			%>
			<span>
				<%
					For i = 1 To RPC
				%>
				<a href="<%=NavPage_1%>&pageno=<%=i%>"><%=i%></a>&nbsp;
				<%
					Next
				%>
			</span>
			<%
				if RPC > pageno then
			%>
			<span><a href="<%=NavPage_1%>&pageno=<%=pageno+1%>"><img src="images/Selection_07.jpg" /></a>&nbsp;</span>
			<span><a href="<%=NavPage_1%>&pageno=<%=RPC%>"><img src="images/Selection_09.jpg" /></a></span>
			<%
				Else
			%>
			<span><img src="images/Selection_07.jpg" />&nbsp;</span>
			<span><img src="images/Selection_09.jpg" /></span>
			<%
				End If
			%>
		</div>
	</div>
</div>