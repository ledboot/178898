<%
	Set rs = rslm(1, "��ϵ����")
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
			YWQQForPicStr = "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank""><img SRC=""http://wpa.qq.com/pa?p=1:" & YWQQArray(iCount) & ":6"" alt=""�澩����ȫ���ͷ�"" title=""�澩����ȫ���ͷ�"" border=""0"" align=""absmiddle""></a>"
		Else
			YWQQStr = YWQQStr & " &nbsp;" & "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank"">" & YWQQArray(iCount) & "</a>"
			YWQQForPicStr = YWQQForPicStr & " &nbsp;" & "<a href=""tencent://message/?uin=" & YWQQArray(iCount) & "&Service=0&sigT=c91aed75f6ce89f7fc95efcb07ae5b66d396e37bb891ba96bf7a544943a82b887bf5289cfc75d7b8406d9df66fafa1ca1074ace33e1a626a"" target=""_blank""><img SRC=""http://wpa.qq.com/pa?p=1:" & YWQQArray(iCount) & ":6"" alt=""�澩����ȫ���ͷ�"" title=""�澩����ȫ���ͷ�"" border=""0"" align=""absmiddle""></a>"
		End If
	Next
%>
<div class="headerA">
	<div class="head_topBg">
        <div class="head_top">
            <div class="hy">��ӭ����<%=("�澩������" & "��")%></div>
            <div class="collection"><a href="../zxbm.asp">��ѵ��֤����</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="../kcjj.asp">��ѵ��Ŀ</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="../lxwm.asp">��ϵ����</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" onClick="window.external.AddFavorite('<%=webAddress%>','<%=website%> ');return false">�����ղ�</a></div>
        </div>
    </div>
	<div class="headLogo">
		<div class="headLogoLeft"><a href="index.asp"></a></div>
		<div class="headLogoRight">
			<div class="headLogoRightTop">��ѯ���ߣ�<b><%=ContactUsPhone%></b>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;�������ߣ�<b><%=ContactUsFax%></b></div>
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
    	<li><a href="../index.asp">��&nbsp;&nbsp;ҳ</a></li>
        <li><a href="../gywm.asp">��������</a></li>
        <li><a href="../kcjj.asp">�γ̼��</a></li>
        <li><a href="../pxxx.asp">��ѵ��Ϣ</a></li>
        <li><a href="../ksdt.asp">���Զ�̬</a></li>
        <li><a href="../zgrz.asp">�ʸ���֤</a></li>
        <li><a href="../zszs.asp">֤��չʾ</a></li>
        <li><a href="../xlwp.asp">ѧ����ƾ</a></li>
        <li><a href="../zxbm.asp">���߱���</a></li>
        <li><a href="../zxly.asp">������ѯ</a></li>
        <li><a href="../lxwm.asp">��ϵ����</a></li>
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
						Set rs = rstw(6, "���ͼƬ")
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