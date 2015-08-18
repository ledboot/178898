<!--#include file="../inc/conn.asp" -->
<!--#include file="../inc/ThumbFunction.asp"-->
<%
	if Session("AdminName")=""  then
		Session.Abandon
		response.write"<SCRIPT language=JavaScript>alert('对不起您无权访问本页！');"
		response.write"javascript:location.href='logout.asp';</SCRIPT>"
		response.end
	end if
	
	Set UpFileObj = Server.CreateObject(Cloudy_Upload)
	UploadObj.OverwriteFiles = False		'不能复盖
	UploadObj.IgnoreNoPost = True
	FileCount=UpFileObj.Save
	
	up_name=UpFileObj.Form("up_name")
	select case up_name
		case "indeximg"
			up_name="indeximg"
		case else
			up_name="indeximg"
	end select
	action=UpFileObj.Form("action")
	ThumbnailComponent=UpFileObj.Form("ThumbnailComponent")
	ThumbnailSize=UpFileObj.Form("ThumbnailSize")
	
	MaxFileSize="20971520"	'文件大小KB
	AllowFileExtStr="rar,zip,swf,txt,gif,jpg,doc,flv"	'文件类型
	IsAddWaterMark="1"	'生成是否要添加水印标记
	FilePath=Domain&UpFiles
	if action="add" or action="" or IsNull(action) then
		AutoReName="2" 	'0不重命名，1“Pic”,2时间,3传递值
	else
		AutoReName="3" 	'0不重命名，1“Pic”,2时间,3传递值
	end if
	MarkType="1" '添加水印1文字0图片
	'文字属性
	MarkText="湘京教育"
	MarkFontColor="#FFFFFF"
	MarkFontName="Arial"
	MarkFontBond="1" '字体是否粗体1是0否
	MarkFontSize=48
	MarkPosition="5"'水印位置坐标  1左上2左下3居中4右上5右下
	'图片属性
	MarkWidth=88
	MarkHeight=30
	MarkPicture="../inc/watermark.gif"
	MarkOpacity=0.6 '水印LOGO图片透明度如60%请填写0.6
	MarkTranspColor="#000000" '水印图片去除底色保留为空则水印图片不去除底色
	'缩略图属性
	select case ThumbnailComponent
		case "1"
			ThumbnailComponent="1"
		case else
			ThumbnailComponent="0"
	end select
	RateTF=0 '0按大小范围1按比例,保证不变形
	select case ThumbnailSize
		case "1"
			ThumbnailWidth=356
			ThumbnailHeight=208
		case else
			ThumbnailWidth=356
			ThumbnailHeight=208
	end select
	ThumbnailRate=0.5'缩略图比例如60%请填写0.6
	
	ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName,IsAddWaterMark,up_name,action)
	if ReturnValue <> "" then
%>
<script language="JavaScript">
	alert('<% = "文件上传失败，错误信息：\n" & ReturnValue %>');
	history.go(-1);
</script>
<%
	end if
	Set UpFileObj=Nothing
%>