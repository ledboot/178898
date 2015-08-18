<%
if Session("AdminName")=""  then
	Session.Abandon
	response.write"<SCRIPT language=JavaScript>alert('对不起您无权访问本页！');"
	response.write"javascript:location.href='logout.asp';</SCRIPT>"
	response.end
end if
up_name=request.QueryString("up_name")
select case up_name
	case "indeximg"
		up_name="indeximg"
	case else
		up_name="indeximg"
end select
action=request.QueryString("action")
ThumbnailComponent=request.QueryString("ThumbnailComponent")
select case ThumbnailComponent
	case "1"
		ThumbnailComponent="1"
	case else
		ThumbnailComponent="0"
end select
ThumbnailSize=request.QueryString("ThumbnailSize")
select case ThumbnailSize
	case "1"
		ThumbnailSize="1"
	case else
		ThumbnailSize="0"
end select
%>
<html>
<head>
<title>上传图片信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/index.CSS" type="text/css">
</head>

<body bgcolor="#F7F7F7" text="#000000" leftmargin="0" topmargin="0">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <form name="formphoto" method="post" action="photo_save.asp" enctype="multipart/form-data">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td><input type="file" name="Photo_address" size="35" style="border: 1 solid #C0C0C0">
				<input type="submit" name="Submit" value="上 传"><input name="up_name" type="hidden" value="<%=up_name%>"><input name="ThumbnailComponent" type="hidden" value="<%=ThumbnailComponent%>"><input name="action" type="hidden" value="<%=action%>"><input name="ThumbnailSize" type="hidden" value="<%=ThumbnailSize%>"></td>
            </tr>
          </table>
        </form>
      </td>
    </tr>
  </table>
</body>
</html>
