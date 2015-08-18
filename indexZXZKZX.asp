<div class="mainA_BottomOne">
	<div class="title">
		<div class="title_Bg"><div class="titlt_L"><a href="ksdt.asp?lx=2">最新招考资讯</a></div><div class="more"><a href="ksdt.asp?lx=2"><img src="images/more.gif" /></a></div></div>
	</div>
	<div class="bxxx_AMain">
		<ul>
			<%
				rsCount = 0
				Set rs = rslm(12, "最新招考资讯")
				While Not rs.BOF And Not rs.EOF
					rsCount = rsCount + 1
					TitleRs = rs("Title")
					AuthorRs = rs("Author")
					CopyFromRs = rs("CopyFrom")
					GjzRs = rs("Gjz")
					Indeximg1Rs = rs("Indeximg1")
					ArticleIDRs = rs("ArticleID")
					hotRs = rs("hot")
					If rsCount = 1 Then
						ArticleIDTempRs = ArticleIDRs
					Else
						ArticleIDTempRs = ArticleIDTempRs & "," & ArticleIDRs
					End If
					If hotRs Then
						linkClassName = "hotLink"
					Else
						linkClassName = ""
					End If
			%>
			<li><a href="showksdt.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>" class="<%=linkClassName%>"><%=LeftTrue(TitleRs, 30)%></a></li>
			<%
					rs.MoveNext
				Wend
				rs.Close
				Set rs = Nothing
			%>
		</ul>
	</div>
</div>