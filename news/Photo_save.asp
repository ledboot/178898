<!--#include file="../inc/conn.asp" -->
<!--#include file="../inc/ThumbFunction.asp"-->
<%
	if Session("AdminName")=""  then
		Session.Abandon
		response.write"<SCRIPT language=JavaScript>alert('�Բ�������Ȩ���ʱ�ҳ��');"
		response.write"javascript:location.href='logout.asp';</SCRIPT>"
		response.end
	end if
	
	Set UpFileObj = Server.CreateObject(Cloudy_Upload)
	UploadObj.OverwriteFiles = False		'���ܸ���
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
	
	MaxFileSize="20971520"	'�ļ���СKB
	AllowFileExtStr="rar,zip,swf,txt,gif,jpg,doc,flv"	'�ļ�����
	IsAddWaterMark="1"	'�����Ƿ�Ҫ���ˮӡ���
	FilePath=Domain&UpFiles
	if action="add" or action="" or IsNull(action) then
		AutoReName="2" 	'0����������1��Pic��,2ʱ��,3����ֵ
	else
		AutoReName="3" 	'0����������1��Pic��,2ʱ��,3����ֵ
	end if
	MarkType="1" '���ˮӡ1����0ͼƬ
	'��������
	MarkText="�澩����"
	MarkFontColor="#FFFFFF"
	MarkFontName="Arial"
	MarkFontBond="1" '�����Ƿ����1��0��
	MarkFontSize=48
	MarkPosition="5"'ˮӡλ������  1����2����3����4����5����
	'ͼƬ����
	MarkWidth=88
	MarkHeight=30
	MarkPicture="../inc/watermark.gif"
	MarkOpacity=0.6 'ˮӡLOGOͼƬ͸������60%����д0.6
	MarkTranspColor="#000000" 'ˮӡͼƬȥ����ɫ����Ϊ����ˮӡͼƬ��ȥ����ɫ
	'����ͼ����
	select case ThumbnailComponent
		case "1"
			ThumbnailComponent="1"
		case else
			ThumbnailComponent="0"
	end select
	RateTF=0 '0����С��Χ1������,��֤������
	select case ThumbnailSize
		case "1"
			ThumbnailWidth=356
			ThumbnailHeight=208
		case else
			ThumbnailWidth=356
			ThumbnailHeight=208
	end select
	ThumbnailRate=0.5'����ͼ������60%����д0.6
	
	ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName,IsAddWaterMark,up_name,action)
	if ReturnValue <> "" then
%>
<script language="JavaScript">
	alert('<% = "�ļ��ϴ�ʧ�ܣ�������Ϣ��\n" & ReturnValue %>');
	history.go(-1);
</script>
<%
	end if
	Set UpFileObj=Nothing
%>