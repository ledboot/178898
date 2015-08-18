<%
	Set rs = rslm(1, "联系我们")
	ContactUsAddress = rs("Title")
	ContactUsPhone = rs("subhead")
	ContactUsFax = rs("Author")
	ContactUsEMail = rs("CopyFrom")
	ContactUsgjzRs = rs("gjz")
	ContactUsindeximgRs = rs("indeximg")
	rs.Close
	Set rs = Nothing
	
	ContactUsPhoneArray = Split(ContactUsPhone, "|")
	For iCount = 0 To UBound(ContactUsPhoneArray)
		If iCount = 0 Then
			ContactUsPhoneStr = ContactUsPhoneArray(iCount)
		Else
			ContactUsPhoneStr = ContactUsPhoneStr & " &nbsp;" & ContactUsPhoneArray(iCount)
		End If
	Next
	
	YWQQArray = Split(ContactUsEMail, "|")
	For iCount = 0 To UBound(YWQQArray)
		If iCount = 0 Then
			YWQQStr = "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank"">" & YWQQArray(iCount) & "</a>"
			YWQQForPicStr = "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank""><img SRC=""http://wpa.qq.com/pa?p=1:" & YWQQArray(iCount) & ":6"" alt=""湘京教育全国客服"" title=""湘京教育全国客服"" border=""0"" align=""absmiddle""></a>"
		Else
			YWQQStr = YWQQStr & " &nbsp;" & "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank"">" & YWQQArray(iCount) & "</a>"
			YWQQForPicStr = YWQQForPicStr & " &nbsp;" & "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank""><img SRC=""http://wpa.qq.com/pa?p=1:" & YWQQArray(iCount) & ":6"" alt=""湘京教育全国客服"" title=""湘京教育全国客服"" border=""0"" align=""absmiddle""></a>"
		End If
	Next
%>
<div class="headerA">
	<div class="head_topBg">
        <div class="head_top">
            <div class="hy">欢迎光临<%=("湘京教育网" & "！")%></div>
            <div class="collection"><a href="../zxbm.asp">培训考证报名</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="../kcjj.asp">培训科目</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="../lxwm.asp">联系我们</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" onClick="window.external.AddFavorite('<%=webAddress%>','<%=website%> ');return false">加入收藏</a></div>
        </div>
    </div>
	<div class="headLogo">
		<div class="headLogoLeft"><a href="index.asp"></a></div>
		<div class="headLogoRight">
			<div class="headLogoRightTop">咨询热线：<b><%=ContactUsPhone%></b>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;报名热线：<b><%=ContactUsFax%></b></div>
			<div class="headLogoRightBottom">
				<div class="searhFormLink"><a href="javascript:void(0)" onclick="searchInfo()"></a></div>
				<div class="searchFormText"><input type="text" name="searchWordStr" id="searchWordStr" class="searchInputText" /></div>
				<div class="searchFormPic" onclick="searchInfo('searchWordStr')"></div>
			</div>
		</div>
	</div>
</div>
<div class="navA">
	<ul>
    	<li><a href="../index.asp">首&nbsp;&nbsp;页</a></li>
        <li><a href="../gywm.asp">关于我们</a></li>
        <li><a href="../kcjj.asp">课程简介</a></li>
        <li><a href="../pxxx.asp">培训信息</a></li>
        <li><a href="../ksdt.asp">考试动态</a></li>
        <li><a href="../zgrz.asp">资格认证</a></li>
        <li><a href="../zszs.asp">证书展示</a></li>
        <li><a href="../xlwp.asp">学历文凭</a></li>
        <li><a href="../zxbm.asp">在线报名</a></li>
        <li><a href="../zxly.asp">在线咨询</a></li>
        <li><a href="../lxwm.asp">联系我们</a></li>
    </ul>
</div>
<div class="bannerForBob">
	<div class="bannerForBobTop"></div>
	<div class="bannerForBobCenter">
		<div class="bannerForBobCenterLeft"></div>
		<div class="bannerForBobCenterCenter">
			<div class="bannerForBobSlideDiv" id="bannerForBobSlideDiv">
				<ul class="bannerForBobSlideUl" id="bannerForBobSlideUl">
					<%
						rsCount = 0
						bigImgLinkStr = ""
						smallImgLinkStr = ""
						Set rs = rstw(6, "广告图片")
						While Not rs.BOF And Not rs.EOF
							rsCount = rsCount + 1
							
							TitleRs = rs("Title")
							SubheadRs = rs("Subhead")
							GjzRs = rs("Gjz")
							IndeximgRs = rs("Indeximg")
							Indeximg1Rs = rs("Indeximg1")
							ArticleIDRs = rs("ArticleID")
							
							If Not checkIsEmpty(SubheadRs) Then
								bigImgLinkStr = "<a href=""showggtp.asp?id=" & ArticleIDRs & " title=""" & TitleRs & """ target=""_blank""><img src=""" & (DoMain & IndeximgRs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ class=""bannerForBobSlideImg"" /></a>"
							Else
								bigImgLinkStr = "<img src=""" & (DoMain & IndeximgRs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ class=""bannerForBobSlideImg"" />"
							End If
							
							If rsCount = 1 Then
								If Not checkIsEmpty(SubheadRs) Then
									smallImgLinkStr = "<a href=""showggtp.asp?id=" & ArticleIDRs & " title=""" & TitleRs & """ target=""_blank""><img src=""" & (DoMain & Indeximg1Rs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ /></a>"
								Else
									smallImgLinkStr = "<img src=""" & (DoMain & Indeximg1Rs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ class=""bannerForBobSlideImg"" />"
								End If
							Else
								If Not checkIsEmpty(SubheadRs) Then
									smallImgLinkStr = smallImgLinkStr & "|" & "<a href=""showggtp.asp?id=" & ArticleIDRs & " title=""" & TitleRs & """ target=""_blank""><img src=""" & (DoMain & Indeximg1Rs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ /></a>"
								Else
									smallImgLinkStr = smallImgLinkStr & "|" & "<img src=""" & (DoMain & Indeximg1Rs) & """ title=""" & TitleRs & """ alt=""" & GjzRs & """ class=""bannerForBobSlideImg"" />"
								End If
							End If								
					%>
					<li class="bannerForBobSlideLi"><%=bigImgLinkStr%></li>
					<%
							rs.MoveNext
						Wend
						rs.Close
						Set rs = Nothing
					%>
				</ul>
			</div>
			<ul class="bannerForBobSlideTab" id="bannerForBobSlideTab">
				<%
					If rsCount > 1 Then
						smallImgLinkStrArray = Split(smallImgLinkStr, "|")
						For iTemp = 1 To rsCount
				%>
				<li><%=smallImgLinkStrArray(iTemp - 1)%></li>
				<%
						Next
					End If
				%> 
			</ul>
		</div>
		<SCRIPT type=text/javascript> 
			new Marquee(
			{
				MSClass	  : ["bannerForBobSlideDiv", "bannerForBobSlideUl", "bannerForBobSlideTab"],
				Width	  : 1001,
				Height	  : 310,
				DelayTime : 9,
				WaitTime  : 9,
				ScrollStep: 1,
				SwitchType: 2,
				AutoStart : 1
			})
		</SCRIPT>
		<div class="bannerForBobCenterRight"></div>
	</div>
	<div class="bannerForBobBottom"></div>
</div>