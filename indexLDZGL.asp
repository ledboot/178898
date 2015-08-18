<div class="jzgcl">
	<div class="ldzgl_Ay"><a href="flxx.asp?lx=3"></a></div>
	<div class="jzgcl_A">
		<%
			rsCount = 0
			Set rs = rslm(22, "劳动资格类")
			While Not rs.BOF And Not rs.EOF
				rsCount = rsCount + 1
				TitleRs = rs("Title")
				AuthorRs = rs("Author")
				CopyFromRs = rs("CopyFrom")
				GjzRs = rs("Gjz")
				Indeximg1Rs = rs("Indeximg1")
				ArticleIDRs = rs("ArticleID")
				ontopRs = rs("ontop")
				hotRs = rs("hot")
				If rsCount = 1 Then
					ArticleIDTempRs = ArticleIDRs
				Else
					ArticleIDTempRs = ArticleIDTempRs & "," & ArticleIDRs
				End If
				
				linkClassName = ""
				If ontopRs Then
					linkClassName = "redLinkClass"
				ElseIf hotRs Then
					linkClassName = "greenLinkClass"
				Else
					linkClassName = ""
				End If
		%>
		<a href="showflxx.asp?id=<%=ArticleIDRs%>" title="<%=AuthorRs%>" target="_blank" class="<%=linkClassName%>"><%=TitleRs%></a>
		<%
				rs.MoveNext
			Wend
			rs.Close
			Set rs = Nothing
		%>
		<a href="flxx.asp?lx=3" class="moreLinkClass">>>更多...</a>
	</div>
</div>