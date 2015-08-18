<div style="height:10px; width:700px; clear:both"></div>
<div class="tzal_C_bott">
	<%
		If Not checkIsEmpty(rsPreviousStrTemp) Then
			rsPreviousStrTempArray = Split(rsPreviousStrTemp, "|")
			previousLinkTitleStr = rsPreviousStrTempArray(0)
			If rsPreviousStrTempArray(1) <> "无副标题" Then
				previousLinkTitleStr = previousLinkTitleStr & rsPreviousStrTempArray(1)
			End If
			rsPreviousLinkStr = "<a href=""" & showNavPage & "?id=" & rsPreviousStrTempArray(2) & """ title=""" & previousLinkTitleStr & """>" & rsPreviousStrTempArray(0) & "</a>"
	%>
	<span>上一篇：<%=rsPreviousLinkStr%></span><br />
	<%
		End If
	%>
	<%
		If Not checkIsEmpty(rsNextArticleStrTemp) Then
			rsNextArticleStrTempArray = Split(rsNextArticleStrTemp, "|")
			nextArticleLinkTitleStr = rsNextArticleStrTempArray(0)
			If rsNextArticleStrTempArray(1) <> "无副标题" Then
				nextArticleLinkTitleStr = nextArticleLinkTitleStr & rsNextArticleStrTempArray(1)
			End If
			rsNextArticleLinkStr = "<a href=""" & showNavPage & "?id=" & rsNextArticleStrTempArray(2) & """ title=""" & nextArticleLinkTitleStr & """>" & rsNextArticleStrTempArray(0) & "</a>"
	%>
	<span>下一篇：<%=rsNextArticleLinkStr%></span>
	<%
		End If
	%>
</div>