<div class="zdhd_C">
	<ul>
		<%
			pageno = CInt(Request("pageno"))
			If pageno = 0 Or IsEmpty(pageno) Or IsNull(pageno) Then
				pageno = 0
			End If
			set rs = rslm2(lmname_1)
			RCount = 16
			RRC = CInt(rs.RecordCount)
			If RRC > 0 Then
				if CInt(pageno) = 0 then pageno = 1
				rs.pagesize = RCount
				rs.absolutepage = pageno
				RPC = CInt(rs.pagecount)
				if not rs.bof and not rs.eof then
					for i = 1 to RCount						
						ArticleIDRs = rs("ArticleID")
						TitleRs = rs("Title")
						indeximgRs = rs("indeximg")
						updatetimeRs = rs("updatetime")
						
						If Not checkIsEmpty(indeximgRs) Then
							haveImgMsg = "<img src=""images/hoveImg.gif"" width=""16"" height=""16"" />"
						Else
							haveImgMsg = ""
						End If
		%>
		<li><span div class="rq">[<%=Format_Time(updatetimeRs, 2)%>]</span><a href="<%=showNavPage%>?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><span class="li_liBG"></span><%=LeftTrue(TitleRs, 80)%></a></li>
		<%
						rs.MoveNext
						If rs.EOF Then
							Exit For
						End If
					Next
				End If
			End If
			rs.Close
			Set rs = Nothing
		%>
	</ul>
</div>