<%
Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName,IsAddWaterMark,up_name,action)
	Dim ErrStr,NoUpFileTF,FileName,FileExtName,FileContent,SameFileExistTF
	NoUpFileTF = True
	ErrStr = ""
	For Each FormName in UpFileObj.Files
		SameFileExistTF = False
		FileName = FormName.OriginalFileName
		If NoIllegalStr(FileName)=False Then
			ErrStr=ErrStr&"文件：上传被禁止！\n"
		End If
		FileExtName = Replace(FormName.Ext,".","")
		if FormName.OriginalSize > 1 then
			NoUpFileTF = False
			ErrStr = ""
			if FormName.OriginalSize > CLng(FileSize)*1024 then
				ErrStr = ErrStr & FileName & "文件:超过了限制，最大只能上传" & FileSize & "K的文件\n"
			end if
			Path=Path&Year(now())&"-"&Month(now())&"/"
			UpFileObj.CreateDirectory Server.MapPath(Path), True
		'是否存在重名文件
			if AutoRename = "0" then
				If UpFileObj.FileExists(Server.MapPath(Path & FileName)) = True  then
					ErrStr = ErrStr & FileName & "文件:存在同名文件\n"
				else
					SameFileExistTF = True
				end if
			else
				SameFileExistTF = True
			End If
			if CheckFileType(AllowExtStr,FileExtName) = False then
				ErrStr = ErrStr & FileName & "文件:不允许上传,上传文件类型有\n" + AllowExtStr + "\n"
			end if
			if ErrStr = "" then
				if SameFileExistTF = True then
					SaveFile Path,FormName,AutoReName,IsAddWaterMark,up_name,action
				else
					SaveFile Path,FormName,"",IsAddWaterMark,up_name,action
				end if
			else
				CheckUpFile = CheckUpFile & ErrStr
			end if
		end if
	Next
	if NoUpFileTF = True then
		CheckUpFile = "没有上传文件"
	end if
End Function

Function SaveFile(FilePath,FormNameItem,AutoNameType,IsAddWaterMark,up_name,action)
	Dim FileName,FileExtName,FileContent,FormName,RandomFigure
	Randomize 
	RandomFigure = CStr(Int((99999 * Rnd) + 1))
	FileName = FormNameItem.OriginalFileName
	FileExtName = LCase(Replace(FormNameItem.Ext,".",""))
	FileExtName=DealExtName(FileExtName)
	If AutoNameType = "1" Then
		FileName = FilePath & "Pic" & FileName
	elseif AutoNameType = "2" Then
		FileName = FilePath & Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
	elseif AutoNameType = "3" Then
		If UpFileObj.FileExists(Server.MapPath(Domain&action)) = True  then
			UpFileObj.DeleteFile Server.MapPath(Domain&action)
		end if
		If UpFileObj.FileExists(Server.MapPath(Domain&left(action,InStrRev(action, ".")-1)&"_1."&LCase(mid(action,InStrRev(action, ".")+1)))) = True  then
			'UpFileObj.DeleteFile Server.MapPath(left(action,InStrRev(action, ".")-1)&"_1."&LCase(mid(action,InStrRev(action, ".")+1)))
			UpFileObj.DeleteFile Server.MapPath(Domain&left(action,InStrRev(action, ".")-1)&"_1."&LCase(mid(action,InStrRev(action, ".")+1)))
		end if
		FileName = FilePath & Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
	Else
		FileName = FilePath&FileName
	End If

	FileName_form=replace(FileName,Domain,"")
	'FileName_form=FileName
	FileName_form_1=left(FileName_form,InStrRev(FileName_form, ".")-1)&"_1."&FileExtName
	FormNameItem.SaveAsVirtual FileName
	if ThumbnailComponent="1" then
		FileName_form_1=left(FileName_form,InStrRev(FileName_form, ".")-1)&"_1."&FileExtName
		'CreateThumbnailEx FileName,left(FileName_form,InStrRev(FileName_form, ".")-1)&"_1."&FileExtName
		CreateThumbnailEx FileName,Domain&left(FileName_form,InStrRev(FileName_form, ".")-1)&"_1."&FileExtName
	else
		FileName_form_1=""
	end if
	If IsAddWaterMark = "1" Then   '在保存好的图片上添加水印
		AddWaterMark FileName
	End if
	response.write "<span style='font-size:12px;BACKGROUND: #F7F7F7;'>上传成功！</span>"
	response.write "<script language=""JavaScript"">"
	response.write "parent.form1."+up_name+".value='"+FileName_form+"';"
	if ThumbnailComponent="1" then
	response.write "parent.form1."+up_name&"1"+".value='"+FileName_form_1+"';"
	end if
	response.write "</script>"
End Function

Function CheckFileType(AllowExtStr,FileExtName)
	Dim i,AllowArray
	AllowArray = Split(AllowExtStr,",")
	FileExtName = LCase(FileExtName)
	CheckFileType = False
	For i = LBound(AllowArray) to UBound(AllowArray)
		if LCase(AllowArray(i)) = LCase(FileExtName) then
			CheckFileType = True
		end if
	Next
	if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" then
		CheckFileType = False
	end if
End Function

Function DealExtName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		DealExtName = Lcase(UpFileExt)
		DealExtName = Replace(DealExtName,Chr(0),"")
		DealExtName = Replace(DealExtName,".","")
		DealExtName = Replace(DealExtName,"'","")
		DealExtName = Replace(DealExtName,"asp","")
		DealExtName = Replace(DealExtName,"asa","")
		DealExtName = Replace(DealExtName,"aspx","")
		DealExtName = Replace(DealExtName,"cer","")
		DealExtName = Replace(DealExtName,"cdx","")
		DealExtName = Replace(DealExtName,"htr","")
End Function

Function NoIllegalStr(Byval FileNameStr)
	Dim Str_Len,Str_Pos
	Str_Len=Len(FileNameStr)
	Str_Pos=InStr(FileNameStr,Chr(0))
	If Str_Pos=0 or Str_Pos=Str_Len then
	 	NoIllegalStr=True
	Else
	 	NoIllegalStr=False
	End If
End function

'为文件添加水印
Function AddWaterMark(FileName)
	Dim strMarkSettingSql,MarkSettingRs,objFileSystem,strFileExtName,objImage
	If InStr(FileName,":") = 0 Then												'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	If FileName <> "" and not IsNull(FileName) Then								'文件名是否不为空,否则退出
		strFileExtName = ""
		If InStr(FileName,".") <> 0 Then
			strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
		End if
		If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'文件不是可用图片则退出
			Exit Function
		End if

		Set objFileSystem = Server.CreateObject(Cloudy_FSO)
		If objFileSystem.FileExists(FileName) Then				'文件存在,否则退出
			If IsObjInstalled(Cloudy_JPEG) Then					'AspJpeg组件已安装,否则退出
				If IsExpired(Cloudy_JPEG) Then
					Response.Write("Persits.Jpeg组件已过期，请选择其他组件或关闭水印功能。")
					Response.End
				End if

				If MarkType = "1" Then				'添加文字水印
					AddTextMark MarkText,MarkFontColor,MarkFontName,MarkFontBond,MarkFontSize,MarkPosition,FileName
				ElseIf MarkType = "0" Then												'添加图片水印
					AddPictureMark MarkWidth,MarkHeight,MarkPicture,MarkOpacity,MarkTranspColor,MarkPosition,FileName
				End if
			End if
		End if
		Set objFileSystem = nothing
	End if
End Function

'为图片添加文字水印
Function AddTextMark(MarkText,MarkFontColor,MarkFontName,MarkFontBond,MarkFontSize,MarkPosition,FileName)
	Dim objImage,X,Y,Text,TextWidth,FontColor,FontName,FondBond,FontSize,OriginalWidth,OriginalHeight
	If InStr(FileName,":") = 0 Then			'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	Text = Trim(MarkText)
	If Text = "" Then
		Exit Function
	End if
	FontColor = Replace(MarkFontColor,"#","&H")
	FontName = MarkFontName
	If MarkFontBond = "1" Then
		FondBond = True
	Else
		FondBond = False
	End if
	FontSize = Cint(MarkFontSize)

			If Not IsObjInstalled(Cloudy_JPEG) Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(Cloudy_JPEG)
			objImage.Open FileName
			objImage.Canvas.Font.Color = FontColor
			objImage.Canvas.Font.Family = FontName
			objImage.Canvas.Font.Bold = FondBond
			objImage.Canvas.Font.Size = FontSize
			objImage.Canvas.Font.ShadowColor = &H000000 '阴影色彩 
			objImage.Canvas.Font.ShadowYOffset = 1 
			objImage.Canvas.Font.ShadowXOffset = 1 
			objImage.Canvas.Brush.Solid = True 
			objImage.Canvas.Font.Quality = 4  '输出质量

			TextWidth = objImage.Canvas.GetTextExtent(Text)										'计算GB2313编码的字符串所占宽度
			
			If objImage.OriginalWidth < TextWidth Or objImage.OriginalHeight < FontSize Then	'如果图片高度小于字体大小或宽度小于字符串宽度则退出
				Exit Function
			End if
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,TextWidth,FontSize '计算坐标
			aa=objImage.Binary '将原始数据赋给aa
			objImage.Canvas.PrintText X,Y,Text '水印位置及文字 
			bb=objImage.Binary '将文字水印处理后的值赋给bb，这时，文字水印没有不透明度 
			'============调整文字透明度================ 
			Set MyJpeg = Server.CreateObject(Cloudy_JPEG) 
			MyJpeg.OpenBinary aa 
			Set Logo = Server.CreateObject(Cloudy_JPEG) 
			Logo.OpenBinary bb 
			MyJpeg.DrawImage 0,0, Logo, 0.2 '0.3是透明度 
			cc=MyJpeg.Binary '将最终结果赋值给cc,这时也可以生成目标图片了 
			MyJpeg.Save (FileName) 
			set aa=nothing 
			set bb=nothing 
			set cc=nothing 
			objImage.close 
			MyJpeg.Close 
			Logo.Close
			Set objImage = nothing
			Set MyJpeg = nothing
			Set Logo = nothing
End Function

'为图片添加图片水印
Function AddPictureMark(MarkWidth,MarkHeight,MarkPicture,MarkOpacity,MarkTranspColor,MarkPosition,FileName)
	Dim objImage,objMark,X,Y,OriginalWidth,OriginalHeight,Position
	If InStr(FileName,":") = 0 Then																'把文件名转换为实际路径
		FileName = Server.Mappath(FileName)
	End if
	If IsNull(MarkWidth) Or MarkWidth = "" Then
		MarkWidth = 0
	Else
		MarkWidth = Cint(MarkWidth)
	End if
	If IsNull(MarkHeight) Or MarkHeight = "" Then
		MarkHeight = 0
	Else
		MarkHeight = Cint(MarkHeight)
	End if
	If MarkPicture = "" Then
		Exit Function
	End if
	If IsNull(MarkOpacity) Or MarkOpacity = "" Then
		MarkOpacity = 1
	Else
		MarkOpacity = Csng(MarkOpacity)
	End if
	If MarkTranspColor <> "" Then																'转换颜色代码
		MarkTranspColor = Replace(MarkTranspColor,"#","&H")
	Else
	End if
			If Not IsObjInstalled(Cloudy_JPEG) Then
				Exit Function
			End if
			Set objImage = Server.CreateObject(Cloudy_JPEG)
			Set objMark = Server.CreateObject(Cloudy_JPEG)
			objImage.Open FileName
			If objImage.OriginalWidth < MarkWidth Or objImage.OriginalHeight < MarkHeight Then	'如果图片高度小于水印高度或宽度小于字水印宽度则退出
				Exit Function
			End if
			objMark.Open Server.Mappath(MarkPicture)
			GetPostion Cint(MarkPosition),X,Y,objImage.OriginalWidth,objImage.OriginalHeight,MarkWidth,MarkHeight '计算坐标
				objImage.Canvas.Pen.Color  = &H000000
				objImage.Canvas.Pen.Width  = 1
				objImage.Canvas.Brush.Solid = False
			If MarkTranspColor <> "" Then
				objImage.DrawImage X,Y,objMark,MarkOpacity,MarkTranspColor
			else
				objImage.DrawImage X,Y,objMark,MarkOpacity
			End if
			objImage.Save FileName
	Set objImage = nothing
	Set objMark = nothing
End Function

'计算水印相对图片的坐标
Function GetPostion(MarkPosition,X,Y,ImageWidth,ImageHeight,MarkWidth,MarkHeight)
	Select Case Cint(MarkPosition)
		Case 1
			X = 1
			Y = 1
		Case 2
			X = 1
			Y = Int(ImageHeight - MarkHeight - 1)
		Case 3
			X = Int((ImageWidth - MarkWidth)/2)
			Y = Int((ImageHeight - MarkHeight)/2)
		Case 4
			X = Int(ImageWidth - MarkWidth - 1)
			Y = 1
		Case 5
			X = Int(ImageWidth - MarkWidth - 1)
			Y = Int(ImageHeight - MarkHeight - 1)
	End Select						
End Function

'由原图片根据数据里保存的设置生成缩略图
Function CreateThumbnailEx(FileName,ThumbnailFileName)
	Dim strSql,RsThumbnailSetting
	CreateThumbnailEx = False
	If ThumbnailComponent <> "0" and (not IsNull(ThumbnailComponent))Then
		If RateTF = 0 Then
			CreateThumbnailEx = CreateThumbnail(FileName,Cint(ThumbnailWidth),Cint(ThumbnailHeight),0,ThumbnailFileName)
		Else
			CreateThumbnailEx = CreateThumbnail(FileName,0,0,Csng(ThumbnailRate),ThumbnailFileName)
		End if
	End if
End Function
'由原图片生成指定宽度和高度的缩略图
Function CreateThumbnail(FileName,Width,Height,Rate,ThumbnailFileName)
	Dim strSql,RsSetting,objImage,iWidth,iHeight,strFileExtName
	CreateThumbnail = False
	If IsNull(FileName) Then									'如果原图片未指定直接退出
		Exit Function
	Elseif FileName="" Then
		Exit Function
	End if
	If InStr(FileName,".") <> 0 Then
		strFileExtName = Lcase(Trim(Mid(FileName,InStrRev(FileName,".")+1)))
	End if
	If strFileExtName <> "jpg" and strFileExtName <> "gif" and strFileExtName <> "bmp" and strFileExtName <> "png" Then'文件不是可用图片则退出
		Exit Function
	End if
	If IsNull(ThumbnailFileName) Then							'如果缩略图未指定保存路径直接退出
		Exit Function
	Elseif ThumbnailFileName="" Then
		Exit Function
	End if
	If IsNull(Width) Then										'如果缩略图宽度未指定则将其指定为0
		Width = 0
	Elseif Width="" Then
		Width = 0
	End if
	If IsNull(Rate) Then										'如果缩略图缩放比例未指定则将其指定为0
		Rate = 0
	Elseif Rate="" Then
		Rate = 0
	End if
	If IsNull(Height) Then										'如果缩略图高度未指定则将其指定为0
		Height = 0
	Elseif Height="" Then
		Height = 0
	End if
	If InStr(FileName,":") = 0 Then								'原图片路径转换化物理路径
		FileName = Server.Mappath(FileName)
	End if
	If InStr(ThumbnailFileName,":") = 0 Then					'缩略图路径转换化物理路径
		ThumbnailFileName = Server.Mappath(ThumbnailFileName)
	End if
	Width = Cint(Width)
	Height = Cint(Height)
	Rate = CSng(Rate)
	
	Select Case Cint(ThumbnailComponent)
		Case 0													'缩略图功能关闭,退出
			Exit Function
		Case 1
			If Not IsObjInstalled(Cloudy_JPEG) Then			'Persits.Jpeg未安装,退出
				Exit Function
			End if
			If IsExpired(Cloudy_JPEG) Then
				Response.Write("Persits.Jpeg组件已过期，请选择其他组件或关闭缩略图功能。")
				Response.End
			End if
			Set objImage = Server.CreateObject(Cloudy_JPEG)
			objImage.Open FileName
			If Rate = 0 and Width <> 0 and Height<> 0 Then
				If Width < objImage.OriginalWidth Then
					If Height<objImage.OriginalHeight*Width/objImage.OriginalWidth Then
						objImage.Width = objImage.OriginalWidth*Height/objImage.OriginalHeight
						objImage.Height = objImage.OriginalHeight*Height/objImage.OriginalHeight
					Else
						objImage.Width = Width
						objImage.Height = objImage.OriginalHeight*Width/objImage.OriginalWidth
					End if
				Else
					If Height < objImage.OriginalHeight Then
						objImage.Width = objImage.OriginalWidth*Height/objImage.OriginalHeight
						objImage.Height = Height
					Else
						objImage.Width = objImage.OriginalWidth
						objImage.Height = objImage.OriginalHeight
					End if
				End if
			'If Rate = 0 and (Width <> 0 Or Height<> 0) Then
				'If Width < objImage.OriginalWidth And Height < objImage.OriginalHeight Then
				'	If Width = 0 and Height <> 0 Then
				'		objImage.Width = objImage.OriginalWidth/objImage.OriginalHeight*Height
				'		objImage.Height = Height
				'	Elseif Width <> 0 and Height = 0 Then
				'		objImage.Width = Width
				'		objImage.Height = objImage.OriginalHeight/objImage.OriginalWidth*Width
				'	ElseIf Width <> 0 and Height <> 0 Then
				'		objImage.Width = Width
				'		objImage.Height = Height
				'	End if
				'End if
			Elseif  Rate = 1 Then
				objImage.Width = objImage.OriginalWidth*Rate
				objImage.Height = objImage.OriginalHeight*Rate
			End if
			objImage.Save ThumbnailFileName
	End Select
	CreateThumbnail = True	
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function IsExpired(strClassString)
	On Error Resume Next
	IsExpired = True
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)

	If 0 = Err Then
		If xTestObj.Expires > Now Then
			IsExpired = False
		End if
	End if
	Set xTestObj = Nothing
	Err = 0
End Function
%>