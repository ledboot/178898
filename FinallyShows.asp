<div class="xwdt_cWz">
	<div class="gsxw_DATop"><b><%=Title%></b>
		<div class="xwdt_dkei"><span>���ߣ�<%=trim(Author)%></span><span>ժҪ�ڣ�<%=CopyFrom%></span><span>����ʱ�䣺<%=trim(Format_Time(UpdateTime, 2))%></span><span>���������<%=trim(Hits)%>��</span></div>
	</div>
	<div style="height:15px; width:650px; clear:both"></div>
	<div class="gsjj_F">
		<%
			if checkIsEmpty(indeximg) then
			Else
		%>
		<div class="gsjj_CTp"><img src="<%=Domain&indeximg%>" alt="<%=Title%><%If Not checkIsEmpty(Subhead) Then%>-<%=Subhead%><%End If%>" width="600" /></div>
		<%
			End If
		%>
		<div class="gsjj_CMine"><%=Content%></div>
	</div>
	<!--#include file="FinallyShowShare.asp" -->
	<!--#include file="FinallyShowNavigation.asp" -->
</div>