<div class="mainA_CenterTwo">
	<div class="title">
		<div class="title_Bg"><div class="titlt_L"><a href="ksdt.asp">¿¼ÊÔ¶¯Ì¬</a></div><div class="more"><a href="ksdt.asp"><img src="images/more.gif" /></a></div></div>
	</div>
	<div class="ksdtCenter">
		<div class="ksdtCenterLeft">
			<div class="ksdtCenterLeftTop"></div>
			<div class="ksdtCenterLeftCenter">
				<div class="slideNumDiv">
					<div class="slideDiv" id="slideDiv">
						<ul class="slideUl" id="slideUl">
							<%
								rsCount = 0
								Set rs = rstw4(5, 104100)
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
							<li><a href="showksdt.asp?id=<%=ArticleIDRs%>" title="<%=AuthorRs%>" target="_blank"><img src="<%=(DoMain & Indeximg1Rs)%>" title="<%=CopyFromRs%>" alt="<%=GjzRs%>" border="0" width="268" height="226" /></a></li>
							<%
									rs.MoveNext
								Wend
								rs.Close
								Set rs = Nothing
							%>
						</ul>
					</div>
					<ul class="slideNumUl" id="slideNumUl">
						<%
							For iCount = 1 To rsCount
						%>
						<li><%=iCount%></li>
						<%
							Next
						%>
					</ul>
				</div>
				<script type="text/javascript"> 
					new Marquee(
					{
						MSClassID : "slideDiv",
						ContentID : "slideUl",
						TabID	  : "slideNumUl",
						Direction : 0,
						Step	  : 0.3,
						Width	  : 268,
						Height	  : 226,
						Timer	  : 20,
						DelayTime : 5000,
						WaitTime  : 5000,
						ScrollStep: 226,
						SwitchType: 0,
						AutoStart : 1
					})
				</script>
			</div>
			<div class="ksdtCenterLeftBottom"></div>
		</div>
		<div class="ksdtCenterRight">
			<ul>
				<%
					Set rs = rslm6(10, 104100, ArticleIDTempRs)
					While Not rs.BOF And Not rs.EOF
						TitleRs = rs("Title")
						AuthorRs = rs("Author")
						CopyFromRs = rs("CopyFrom")
						GjzRs = rs("Gjz")
						IndeximgRs = rs("Indeximg")
						ArticleIDRs = rs("ArticleID")
						updatetimeRs = rs("updatetime")
						hotRs = rs("hot")
						
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
</div>