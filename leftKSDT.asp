<div class="bxxx_B">
	<div class="title">
		<div class="title_Bg"><div class="titlt_L"><a href="ksdt.asp">���Զ�̬</a></div><div class="more"><a href="ksdt.asp"><img src="../images/more.gif" /></a></div></div>
	</div>
	<ul>
		<%
			rsCount = 0
			Set rs = rslm(8, "���Զ�̬")
			While Not rs.BOF And Not rs.EOF
				rsCount = rsCount + 1
				TitleRs = rs("Title")
				AuthorRs = rs("Author")
				CopyFromRs = rs("CopyFrom")
				GjzRs = rs("Gjz")
				Indeximg1Rs = rs("Indeximg1")
				ArticleIDRs = rs("ArticleID")
				If rsCount = 1 Then
					ArticleIDTempRs = ArticleIDRs
				Else
					ArticleIDTempRs = ArticleIDTempRs & "," & ArticleIDRs
				End If
		%>
		<li><a href="showksdt.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><%=LeftTrue(TitleRs, 30)%></a></li>
		<%
				rs.MoveNext
			Wend
			rs.Close
			Set rs = Nothing
		%>
	</ul>
</div>