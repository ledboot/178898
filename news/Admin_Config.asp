<%
	Function rslm(top,classname) '---------- 求栏目类别文章
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rslm=Server.createobject(Cloudy_RS)
		Sql="select top "&top&" Title,Author,CopyFrom,subhead,gjz,Classid,indeximg,indeximg1,hot,updatetime,ArticleID,hits,ontop,IncludePic,Content from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm.open sql,conn,1,1
	End Function
	
	Function rslm2(classname) '---------- 求栏目内所有文章
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rslm2=Server.createobject(Cloudy_RS)
		Sql="select Title,Subhead,content,indeximg,indeximg1,updatetime,ArticleID,hits,IncludePic,Author,CopyFrom,Gjz from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm2.open sql,conn,1,1
	End Function
	
	Function rslm3(top, Classid) '---------- 求各栏目的所有新闻带条数
		Set rslm3=Server.createobject(Cloudy_RS)
		Sql="select top " & top & " Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content,UpdateTime,ClassID from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm3.open sql,conn,1,1
	End Function
	
	Function rslm4(top, Classid, ArticleIDTempStr, neSearchClassIdStr) '---------- 求各栏目的所有新闻带条数
		Set rslm4=Server.createobject(Cloudy_RS)
		Sql="select top " & top & " Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content,UpdateTime from Article where ArticleID Not In(" & ArticleIDTempStr & ") and  ClassID Not in(" & neSearchClassIdStr & ") and ClassID like '" & Classid & "%' and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm4.open sql,conn,1,1
	End Function
	
	Function rslm5(top, ClassidStr) '---------- 求各栏目的所有新闻带条数
		Set rslm5=Server.createobject(Cloudy_RS)
		Sql="select top " & top & " Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content,UpdateTime from Article where  ClassID in( " & ClassidStr & ") and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm5.open sql,conn,1,1
	End Function
	
	Function rslm6(top, Classid, ArticleIDTempStr) '---------- 求各栏目的所有新闻带条数
		Set rslm6=Server.createobject(Cloudy_RS)
		Sql="select top " & top & " Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content,UpdateTime,hot from Article where ArticleID Not In(" & ArticleIDTempStr & ") and ClassID like '" & Classid & "%' and Deleted=0 and Passed=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rslm6.open sql,conn,1,1
	End Function
	
	Function rstw(top, classname) '---------- 求各栏目的图说新闻带条数 
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rstw=Server.createobject(Cloudy_RS)
		Sql="select top "&top&" Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw.open sql,conn,1,1
	End Function
	
	Function rstw2(classname) '---------- 求各栏目的图说新闻 
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rstw2=Server.createobject(Cloudy_RS)
		Sql="select Title,Subhead,Author,CopyFrom,Gjz,Content,Indeximg,Indeximg1,ArticleID from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw2.open sql,conn,1,1
	End Function
	
	Function rstw3(searchKW)'---------- 根据关键字求各栏目的图说新闻
		Set rstw3=Server.createobject(Cloudy_RS)
		Sql="select Title,Subhead,Author,CopyFrom,Content,Indeximg,Indeximg1,ArticleID from Article where Deleted=0 and Passed=1 and indeximg<>'' and Title like '%" & searchKW & "%' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw3.open sql,conn,1,1
	End Function
	
	Function rstw4(top, Classid) '---------- 求各栏目的图说新闻带条数
		Set rstw4=Server.createobject(Cloudy_RS)
		Sql="select top "&top&" Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw4.open sql,conn,1,1
	End Function
	
	Function rstw5(Classid) '---------- 求各栏目的图说新闻带条数
		Set rstw5=Server.createobject(Cloudy_RS)
		Sql="select Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw5.open sql,conn,1,1
	End Function
	
	Function rstw6(top, ClassidStr) '---------- 求各栏目的图说新闻带条数
		Set rstw6 = Server.createobject(Cloudy_RS)
		Sql = "select top " & top & " Title,Subhead,Author,CopyFrom,Gjz,Indeximg,Indeximg1,ArticleID,Content from Article where ClassID in(" & ClassidStr & ") and Deleted=0 and Passed=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstw6.open sql,conn,1,1
	End Function
	
	Function rstj(top,classname) '---------- 求各栏目的图说新闻 
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rstj=Server.createobject(Cloudy_RS)
		Sql="select top "&top&" Title,Subhead,indeximg,Indeximg1,ArticleID from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and indeximg<>'' and hot=1 order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstj.open sql,conn,1,1
	End Function
	
	Function rstj2(classname) '---------- 求各栏目的图说新闻 
		Set rs=Server.createobject(Cloudy_RS)
		sql="select XcClassid from ArticleClass where classname='"&classname&"'"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then Error2("有错误，请与管理员联系！"):response.end
		classid=trim(rs(0))
		rs.close
		set rs=nothing
	
		Set rstj2=Server.createobject(Cloudy_RS)
		Sql="select Title,Subhead,indeximg,Indeximg1,ArticleID from Article where ClassID like '"&Classid&"%' and Deleted=0 and Passed=1 and hot=1 and indeximg<>'' order by ontop desc,hot desc,UpdateTime desc,ArticleID desc"
		rstj2.open sql,conn,1,1
	End Function
	
	Function rsPreviousStr(currentArticleID) '---------- 上一篇
		rsPreviousStr = ""
		
		Set rs = Server.CreateObject(Cloudy_RS)
		sql = "select ClassID from Article where ArticleID=" & currentArticleID
		rs.Open sql, conn, 1, 1
		If rs.BOF And rs.EOF Then Error2("有错误，请与管理员联系！"):Response.end
		classid = Trim(rs(0))
		rs.Close
		Set rs = Nothing
		
		rsCount = 0
		Set rsPrevious = Server.CreateObject(Cloudy_RS)
		Sql = "select Title, Subhead, ArticleID from Article where ClassID = " & classid & " and Passed=1 order by ontop desc, hot desc, UpdateTime desc, ArticleID desc"
		rsPrevious.Open sql, conn, 1, 3
		For iCount = 0 To rsPrevious.RecordCount
			If Not rsPrevious.BOF And Not rsPrevious.EOF Then
				rsCount = rsCount + 1
				If currentArticleID = rsPrevious("ArticleID") Then
					If rsCount = 1 Then
						rsPreviousStr = ""
					Else
						rsPrevious.MovePrevious
						rsPreviousStr = rsPrevious("Title")
						If Not checkIsEmpty(rsPrevious("Subhead")) Then
							rsPreviousStr = rsPreviousStr & "|" & rsPrevious("Subhead")
						Else
							rsPreviousStr = rsPreviousStr & "|" & "无副标题"
						End If
						rsPreviousStr = rsPreviousStr & "|" & rsPrevious("ArticleID")
						Exit For
					End If
				End If	
				rsPrevious.MoveNext
			Else
				rsPreviousStr = ""
			End If
		Next
	End Function
	
	Function rsNextArticleStr(currentArticleID) '---------- 下一篇
		rsNextArticleStr = ""
		
		Set rs = Server.CreateObject(Cloudy_RS)
		sql = "select ClassID from Article where ArticleID=" & currentArticleID
		rs.Open sql, conn, 1, 1
		If rs.BOF And rs.EOF Then Error2("有错误，请与管理员联系！"):Response.end
		classid = Trim(rs(0))
		rs.Close
		Set rs = Nothing
		
		rsCount = 0
		Set rsNextArticle = Server.CreateObject(Cloudy_RS)
		Sql = "select Title, Subhead, ArticleID from Article where ClassID = " & classid & " and Passed=1 order by ontop desc, hot desc, UpdateTime desc, ArticleID desc"
		rsNextArticle.Open sql, conn, 1, 3
		For iCount = 0 To rsNextArticle.RecordCount
			If Not rsNextArticle.BOF And Not rsNextArticle.EOF Then
				rsCount = rsCount + 1
				If currentArticleID = rsNextArticle("ArticleID") Then
					If rsCount = rsNextArticle.RecordCount Then
						rsNextArticleStr = ""
					Else
						rsNextArticle.MoveNext
						rsNextArticleStr = rsNextArticle("Title")
						If Not checkIsEmpty(rsNextArticle("Subhead")) Then
							rsNextArticleStr = rsNextArticleStr & "|" & rsNextArticle("Subhead")
						Else
							rsNextArticleStr = rsNextArticleStr & "|" & "无副标题"
						End If
						rsNextArticleStr = rsNextArticleStr & "|" & rsNextArticle("ArticleID")
						Exit For
					End If
				End If	
				rsNextArticle.MoveNext
			Else
				rsNextArticleStr = ""
			End If
		Next
	End Function

	'合并数组并去除重复项
	function F(x)'这里x是形参   
		dim S,D   
		Set D = CreateObject("Scripting.Dictionary")   
		for each xxx in x   
			if not D.Exists(xxx) then
				D.Add xxx,xxx
			End If
		Next
		for each key in D.Keys   
			S = S & key & ","   
		Next 
		set D=nothing   
		F = split(trim(S))   
	end function
	
	Function rsClass() '---------- 查询栏目表 
		Set rsClass = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass Order by Class_sx"
		rsClass.Open sql, conn, 1, 3
	End Function
	
	Function rsClass1(XcClassID) '---------- 查询栏目表中某一层的栏目并按排序字段排序
		Set rsClass1 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass where ScClassid=" & CLng(XcClassID) & " order by Class_sx"
		rsClass1.Open sql, conn, 1, 3
	End Function
	
	Function rsClass2(ClassID) '---------- 查询栏目表中某一个栏目
		Set rsClass2 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass where Class_ID=" & CLng(ClassID)
		rsClass2.Open sql, conn, 1, 3
	End Function
	
	Function rsClass3(ClassID) '---------- 查询栏目表中的某一个栏目的排序字段
		Set rsClass3 = Server.CreateObject("ADODB.Recordset")
		sql = "select Class_sx from Articleclass where Class_ID=" & CLng(ClassID)
		rsClass3.Open sql, conn, 1, 3
	End Function
	
	Function rsClass4(XcClassID) '---------- 查询某栏目的所有子栏目
		Set rsClass4 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass where ScClassID=" & CLng(XcClassID)
		rsClass4.Open sql, conn, 1, 3
	End Function
	
	Function rsClass5(ScClassID) '---------- 查询某子栏目的上一层栏目
		Set rsClass5 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass where XcClassID=" & CLng(ScClassID)
		rsClass5.Open sql, conn, 1, 3
	End Function
	
	Function rsClass6(XcClassID) '---------- 查询某一栏目下的所有子栏目的XcClassid的最大值
		Set rsClass6 = Server.CreateObject("ADODB.Recordset")
		sql = "select max(XcClassid) as XcClassid from Articleclass where ScClassID=" & CLng(XcClassID)
		rsClass6.Open sql, conn, 1, 3
	End Function
	
	Function rsClass7(XcClassID) '---------- 查询某一个栏目
		Set rsClass7 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Articleclass where XcClassID=" & CLng(XcClassID)
		rsClass7.Open sql, conn, 1, 3
	End Function
	
	Function rsDelClass(ClassID) '---------- 删除栏目
		Set rs = rsClass2(ClassID)
		If rs("ScClassid") = 0 Then
			If rs("Class_num") = 0 Then
				Call rsDelClass2(ClassID)
			Else
				Set rs1 = rsClass4(rs("XcClassid"))
				While Not rs1.BOF And Not rs1.EOF
					If rs1("ClassNameD") <> 1 Then
						If rs1("Class_num") <> 0 Then
							Set rs2 = rsClass4(rs1("XcClassid"))
							While Not rs2.BOF And Not rs2.EOF
								If rs2("ClassNameD") <> 1 Then
									Call rsDelArticle(rs2("XcClassid"))
								End If
								rs2.MoveNext
							Wend
							rs2.Close
							Set rs2 = Nothing
							Call rsDelClass1(rs1("XcClassid"))
						Else
							Call rsDelArticle(rs1("XcClassid"))
						End If
					End If
					rs1.MoveNext
				Wend
				rs1.Close
				Set rs1 = Nothing
				Call rsDelClass1(rs("XcClassid"))
				Call rsDelClass2(ClassID)
			End If
		Else
			Set rs3 = rsClass5(rs("ScClassid"))
			SCI = rs3("ScClassid")
			rs3.Close
			Set rs3 = Nothing
			If CLng(SCI) = 0 Then
				If rs("ClasNameD") <> 1 Then
					If CInt(rs("Class_num")) <> 0 Then
						Set rs1 = rsClass4(rs("XcClassid"))
						While Not rs1.BOF And Not rs1.EOF
							If rs1("ClassNameD") <> 1 Then
								Call rsDelArticle(rs1("XcClassid"))
							End If
							rs1.MoveNext
						Wend
						rs1.Close
						Set rs1 = Nothing
						Call rsDelClass1(rs("XcClassid"))
					Else
						Call rsDelArticle(rs("XcClassid"))
					End If
				End If
				Call rsDelClass2(ClassID)
			Else
				If rs("ClassNameD") <> 1 Then
					Call rsDelArticle(rs("XcClassid"))
				End If
				Call rsDelClass2(ClassID)
			End If
		End If
		rs.Close
		Set rs = Nothing
	End Function
	
	Function rsDelClass1(XcClassID) '---------- 删除某栏目的所有子栏目
		conn.execute("delete from Articleclass where ScClassID=" & CLng(XcClassid))
	End Function
	
	Function rsDelClass2(ClassID) '---------- 删除指定id的栏目
		conn.execute("delete from Articleclass where Class_ID=" & CLng(ClassID))
	End Function
	
	Function rsDelArticle(XcClassid)  '--------------------- 删除某个栏目的所有文章
		conn.execute("delete from Article where ClassID=" & CLng(XcClassid))
	End Function
	
	Function rsArticle(XcClassid)  '--------------------- 查询某个栏目的所有文章
		Set rsArticle = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article where ClassID=" & CLng(XcClassid)
		rsArticle.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle1()  '--------------------- 查询所有栏目的文章
		Set rsArticle1 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article"
		rsArticle1.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle2(classid)  '--------------------- 查询指定classid的所有文章
		Set rsArticle2 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article classid=" & clasid
		rsArticle2.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle3()  '--------------------- 查询文章表里自动编号最大的一条
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "select max(ArticleID) as ArticleID from Article"
		rs.Open sql, conn, 1, 3
		Dim ArticleID
		ArticleID = rs("ArticleID")
		rs.Close
		Set rs = Nothing
		Set rsArticle3 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article where ArticleID=" & ArticleID
		rsArticle3.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle4(Articleid)  '--------------------- 查询指定文章id的文章
		Set rsArticle4 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article Where Articleid=" & Articleid
		rsArticle4.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle5(XcClassid)  '--------------------- 查询某个栏目的所有文章并按是否固顶,是否推荐,和文章id排序
		Set rsArticle5 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article where ClassID=" & CLng(XcClassid) & " and Deleted=0 order by ontop desc,hot desc,Articleid desc" 
		rsArticle5.Open sql, conn, 1, 3
	End Function
	
	Function rsArticle6(XcClassid)  '--------------------- 查询某个栏目的所有已经被标记为删除的文章,并按文章id排序
		Set rsArticle6 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from Article where ClassID=" & CLng(XcClassid) & " and Deleted=1 order by Articleid desc" 
		rsArticle6.Open sql, conn, 1, 3
	End Function
	
	Function rsAdmin(username,password) '---------- 查询用户表指定用户名和密码的用户
		Set rsAdmin = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from admin where username='" & replace(username,"'","") & "' and password='" & replace(password,"'","") & "'"
		rsAdmin.Open sql,Conn,1,3
	End Function
	
	Function rsAdmin2() '---------- 查询用户表所有用户
		Set rsAdmin2 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from admin"
		rsAdmin2.Open sql,Conn,1,3
	End Function
	
	Function rsAdmin3(id) '---------- 查询用户表指定id的用户
		Set rsAdmin3 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from admin where id=" & CInt(id)
		rsAdmin3.Open sql,Conn,1,3
	End Function
	
	Function rsProfesion()  '--------------------- 查询所有专业
		Set rsProfesion = Server.CreateObject("ADODB.Recordset")
		sql = "select * from GW_Profession "
		rsProfesion.Open sql, conn, 1, 3
	End Function
	
	Function rsProfesion1()  '--------------------- 查询所有专业自动编号最大的一条
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "select max(ID) as ProfessionID from GW_Profession"
		rs.Open sql, conn, 1, 3
		Dim ProfessionID
		ProfessionID = rs("ProfessionID")
		rs.Close
		Set rs = Nothing
		Set rsProfesion1 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from GW_Profession where ID=" & ProfessionID
		rsProfesion1.Open sql, conn, 1, 3
	End Function
	
	Function rsProfesion3(id)  '--------------------- 查询id专业
		Set rsProfesion3 = Server.CreateObject("ADODB.Recordset")
		sql = "select * from GW_Profession where ID= " & CInt(id)
		rsProfesion3.Open sql, conn, 1, 3
	End Function
	
	
	Function rsStudnet1() '---------- 查询学员表全部学员关联专业表
		Set rsStudnet1 = Server.CreateObject("ADODB.Recordset")
		Sql = " SELECT GW_Student.*, GW_Profession.ProfessionName FROM GW_Student left  JOIN GW_Profession ON GW_Student.ProfessionId = GW_Profession.ID order by GW_Student.Date desc "
		rsStudnet1.Open sql,Conn,1,3
	End Function
	
	Function rsStudnet2(id) '---------- 查询学员表指定id的学员
		Set rsStudnet2 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_Student where id=" & CInt(id)
		rsStudnet2.Open sql,Conn,1,3
	End Function
	
	Function rsStudnet3(idCard) '---------- 查询学员表指定身份证的学员
		Set rsStudnet3 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_Student where IdCard='" &idCard&"'"
		rsStudnet3.Open sql,Conn,1,3
	End Function
	
	
	Function rsStudnet4() '---------- 查询学员表全部
		Set rsStudnet4 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_Student "
		rsStudnet4.Open sql,Conn,1,3
	End Function

	Function rsPaper() '---------- 查询全部试卷
		Set rsPaper = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT *  FROM GW_Paper "
		rsPaper.Open sql,Conn,1,3
	End Function

	Function rsPaper1() '---------- 查询全部试卷
		Set rsPaper1 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT GW_Paper.*, GW_Profession.ProfessionName FROM GW_Paper left JOIN GW_Profession ON GW_Paper.ProfessionId = GW_Profession.ID order by GW_Paper.Date desc"
		rsPaper1.Open sql,Conn,1,3
	End Function
	
	Function rsPaper2(id) '---------- 根据id查找试卷
		Set rsPaper2 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_Paper where id=" & CInt(id)
		rsPaper2.Open sql,Conn,1,3
	End Function
	
	Function rsPaper3(id) '---------- 根据id查找试卷
		Set rsPaper3 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_Paper where Verific=1 and  ProfessionId=" & CInt(id)
		rsPaper3.Open sql,Conn,1,3
	End Function
	
	Function rsSubjectLibrary1() '---------- 查询题库
		Set rsSubjectLibrary1 = Server.CreateObject("ADODB.Recordset")
		Sql = "Select * from GW_SubjectLibrary "
		rsSubjectLibrary1.Open sql,Conn,1,3
	End Function
	
	Function rsSubjectLibrary2(id) '---------- 根据id查找题库
		Set rsSubjectLibrary2 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT GW_SubjectLibrary.*, GW_Paper.* FROM GW_Paper INNER JOIN GW_SubjectLibrary ON GW_Paper.ID = GW_SubjectLibrary.PaperId where GW_SubjectLibrary.id =" & CInt(id)
		rsSubjectLibrary2.Open sql,Conn,1,3
	End Function
	
	Function rsSubjectLibrary3(id) '---------- 根据试卷id查找题库列表
		Set rsSubjectLibrary3 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT GW_SubjectLibrary.*, GW_Paper.* FROM GW_Paper INNER JOIN GW_SubjectLibrary ON GW_Paper.ID = GW_SubjectLibrary.PaperId where GW_Paper.id =" & CInt(id)
		rsSubjectLibrary3.Open sql,Conn,1,3
	End Function
	
	Function rsSubjectLibrary4(id) '---------- 根据试卷id查找题库列表
		Set rsSubjectLibrary4 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT * from GW_SubjectLibrary where id=" &Cint(id)
		rsSubjectLibrary4.Open sql,Conn,1,3
	End Function
	
	Function rsSubjectLibrary5(id) '---------- 根据试卷id查找题库列表
		Set rsSubjectLibrary5 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT GW_SubjectLibrary.ID, GW_SubjectLibrary.PaperId,GW_SubjectLibrary.Title,GW_SubjectLibrary.Options,GW_SubjectLibrary.Answer,GW_SubjectLibrary.Types as mTypes,GW_SubjectLibrary.Score,GW_Paper.* FROM GW_Paper INNER JOIN GW_SubjectLibrary ON GW_Paper.ID = GW_SubjectLibrary.PaperId where GW_Paper.id =" & CInt(id)
		rsSubjectLibrary5.Open sql,Conn,1,3
	End Function
	
	Function rsAnswer1() '---------- 根据试卷id查找题库列表
		Set rsAnswer1 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT * from GW_Answer"
		rsAnswer1.Open sql,Conn,1,3
	End Function
	
	Function rsAnswer2(id) '---------- 根据试卷id查找题库列表
		Set rsAnswer2 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT * from GW_Answer where ID =" &Cint(id)
		rsAnswer2.Open sql,Conn,1,3
	End Function
	
	Function rsScore1() '---------- 根据试卷id查找题库列表
		Set rsScore1 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT * from GW_Score "
		rsScore1.Open sql,Conn,1,3
	End Function
	
	Function rsScore2(id) '---------- 根据试卷id查找题库列表
		Set rsScore2 = Server.CreateObject("ADODB.Recordset")
		Sql = "SELECT * from GW_Score where ID =" &Cint(id)
		rsScore2.Open sql,Conn,1,3
	End Function
%>