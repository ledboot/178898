<div class="mainA_BottomTwo">
	<div class="title">
		<div class="title_Bg"><div class="titlt_L"><a href="zgrz.asp">����ְҵ�ʸ���</a></div><div class="more"><a href="zgrz.asp"><img src="images/more.gif" /></a></div></div>
	</div>
	<div class="kszg_AML">
		<div class="kszg_AML1">
			<div class="kszg_title"><a href="zgrz.asp"><img src="images/kszg_63.gif" /></a><span>�������̿���</span></div>
			<ul>
				<%
					rsCount = 0
					Set rs = rslm(5, "�������̿���")
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
				<li><a href="showzgrz.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><%=LeftTrue(TitleRs, 30)%></a></li>
				<%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</ul>
		</div>
		<div class="kszg_AML2">
			<div class="kszg_title"><a href="zgrz.asp?lx=2"><img src="images/kszg_63.gif" /></a><span>����ְҵ�ʸ���</span></div>
			<ul>
				<%
					rsCount = 0
					Set rs = rslm(5, "����ְҵ�ʸ���")
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
				<li><a href="showzgrz.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><%=LeftTrue(TitleRs, 30)%></a></li>
				<%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</ul>
		</div>
	</div>
	<div class="kszg_AMR">
		<div class="kszg_AMR1">
			<div class="kszg_title"><a href="zgrz.asp?lx=3"><img src="images/kszg_63.gif" /></a><span>���˸ߵȽ�������</span></div>
			<ul>
				<%
					rsCount = 0
					Set rs = rslm(5, "���˸ߵȽ�������")
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
				<li><a href="showzgrz.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><%=LeftTrue(TitleRs, 30)%></a></li>
				<%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</ul>
		</div>
		<div class="kszg_AMR2">
			<div class="kszg_title"><a href="zgrz.asp?lx=4"><img src="images/kszg_63.gif" /></a><span>��ѧ����</span></div>
			<ul>
				<%
					rsCount = 0
					Set rs = rslm(5, "��ѧ����")
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
				<li><a href="showzgrz.asp?id=<%=ArticleIDRs%>" target="_blank" title="<%=TitleRs%>"><%=LeftTrue(TitleRs, 30)%></a></li>
				<%
						rs.MoveNext
					Wend
					rs.Close
					Set rs = Nothing
				%>
			</ul>
		</div>
	</div>
</div>