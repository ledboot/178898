<div class="gywm_B">
	<div class="title">
		<div class="title_Bg"><a href="<%=NavPage%>"><%=lmname%></a></div>
	</div>
	<%
		If lmname = "�ʸ���֤" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>�������̿���</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>����ְҵ�ʸ���</a></li>
		<li<%If lx = 3 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=3"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>���˸ߵȽ�������</a></li>
		<li<%If lx = 4 Then%> class="nav_CLI1"<%End If%>><a href="zgrz.asp?lx=4"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>��ѧ����</a></li>
	</ul>
	<%
		ElseIf lmname = "������ѯ" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="zxly.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>��������</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="lyb.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>���԰�</a></li>
	</ul>
	<%
		ElseIf lmname = "���Զ�̬" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="ksdt.asp?lx=1"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>���Զ�̬</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="ksdt.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>�����п���Ѷ</a></li>
	</ul>
	<%
		ElseIf lmname = "��ѵ��Ϣ" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="pxxx.asp?lx=1"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>��ѵ��Ϣ</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="pxxx.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>��������</a></li>
	</ul>
	<%
		ElseIf lmname = "������Ϣ" Then
	%>
	<ul>
		<li<%If lx = 1 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>����������</a></li>
		<li<%If lx = 2 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp?lx=2"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>ѧ��ְ����</a></li>
		<li<%If lx = 3 Then%> class="nav_CLI1"<%End If%>><a href="flxx.asp?lx=3"><span>&nbsp;&nbsp;<img src="images/li_liB_52.gif" /></span>�Ͷ��ʸ���</a></li>
	</ul>
	<%
		End If
	%>
</div>