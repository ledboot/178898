<div class="gywm_B">
	<div class="title">
		<div class="title_Bg"><a href="<%=NavPage%>"><%=lmname%></a></div>
	</div>
	<%
		If lmname = "资格认证" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>建筑工程考试</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>国家职业资格考试</a></li>
		<li<%If lx = 3 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=3"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>成人高等教育考试</a></li>
		<li<%If lx = 4 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=4"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>自学考试</a></li>
	</ul>
	<%
		ElseIf lmname = "在线咨询" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="zxly.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>在线留言</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="lyb.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>留言板</a></li>
	</ul>
	<%
		ElseIf lmname = "考试动态" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="ksdt.asp?lx=1"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>考试动态</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="ksdt.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>最新招考资讯</a></li>
	</ul>
	<%
		ElseIf lmname = "培训信息" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="pxxx.asp?lx=1"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>培训信息</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="pxxx.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>资料下载</a></li>
	</ul>
	<%
		ElseIf lmname = "分类信息" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>建筑工程类</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>学历职称类</a></li>
		<li<%If lx = 3 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp?lx=3"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>劳动资格类</a></li>
	</ul>
	<%
		End If
	%>
</div>