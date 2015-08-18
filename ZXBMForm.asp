<script language="javascript" type="text/javascript">
	<!--
		function isok(){
			var theform_Obj = document.getElementById("theform");
			var stuName_Obj = document.getElementById("stuName");
			var stuTelephone_Obj = document.getElementById("stuTelephone");
			var stuDeclare_Class_Type_ID_Obj = document.getElementsByName("stuDeclare_Class_Type_ID");
			var stuResume_Text_Obj = document.getElementById("stuResume_Text");
			var stuQQNumber_Obj = document.getElementById("stuQQNumber");
			var stuEMail_Obj = document.getElementById("stuEMail");
			var stuAddress_Obj = document.getElementById("stuAddress");
			
			
			//return false;
			
			//stuName stuSex_<%=i%> birthdayYear birthdayMonth birthdayDay stuBirthday stuI_D_Card stuGraduation_School stuCensus_Register_Provinces_ID stuCensus_Register_City_ID stuTelephone stuZip stuHeight stuEyesight_Left stuEyesight_Right stuDeclare_Class_Type_ID_<%=Class_Type_Id%> stuResume_Text
			if(checkIsEmpty(stuName_Obj)){
				stuName_Obj.value = "";
				alert("对不起，姓名不能为空！");
				stuName_Obj.focus();
				return false;
			}
			if(checkIsEmpty(stuTelephone_Obj)){
				alert("对不起，电话不能为空！");
				stuTelephone_Obj.focus();
				return false;
			}
			if(stuTelephone_Obj.value.isMobile() || stuTelephone_Obj.value.isTel()){
				stuTelephone_Obj.value = stuTelephone_Obj.value.Trim();
			}
			else{
				alert("请输入正确的手机号码或电话号码\n\n例如:13888888888或021-8888888");
				stuTelephone_Obj.focus();
				return false;
			}
			if(checkIsEmpty(stuResume_Text_Obj)){
				stuResume_Text_Obj.value = "未填写";
			}
			if(checkIsEmpty(stuAddress_Obj)){
				stuAddress_Obj.value = "";
				alert("对不起，地址不能为空！");
				stuAddress_Obj.focus();
				return false;
			}
			if(checkIsEmpty(stuQQNumber_Obj)){
				stuQQNumber_Obj.value = "未填写";
			}
			if(checkIsEmpty(stuEMail_Obj)){
				stuEMail_Obj.value = "未填写";
			}
			
			theform_Obj.submit();
		}
	-->
</script>
<div id="ZXBMForm">
	<form name="theform" id="theform" method="post" action="addOnline_Registration.asp">
		<div>
			<div class="FormTitle">姓 &nbsp;&nbsp;&nbsp;名</div>
			<div class="FormContent"><input type="text" name="stuName" id="stuName" value="" class="inputStyle">&nbsp;*&nbsp;</div>
			<div class="FormTitle">电 &nbsp;&nbsp;&nbsp;话</div>
			<div class="FormContent"><input type="text" name="stuTelephone" id="stuTelephone" value="" class="inputStyle">&nbsp;*&nbsp;</div>
		</div>
        <div style="clear:both"></div>
		<div>
			<div class="FormTitle">Q &nbsp;&nbsp;&nbsp;Q</div>
			<div class="FormContent"><input type="text" name="stuQQNumber" id="stuQQNumber" value="" class="inputStyle"></div>
			<div class="FormTitle">邮 &nbsp;&nbsp;&nbsp;箱</div>
			<div class="FormContent"><input type="text" name="stuEMail" id="stuEMail" value="" class="inputStyle"></div>
		</div>
        <div style="clear:both"></div>
		<div>
			<div class="FormTitle">地 &nbsp;&nbsp;&nbsp;址</div>
			<div class="FormContentCol3"><input type="text" name="stuAddress" id="stuAddress" class="inputStyle1"></select>&nbsp;*&nbsp;</div>
		</div>
        <div style="clear:both"></div>
        <div>
			<div class="FormTitle">验 证 码</div>
			<div class="FormContentCol3"><input type="text" name="Passcode" id="Passcode" value="" class="inputStyle"> <img src="inc/VerifyCode.asp" > * </div>
		</div>
        <div style="clear:both"></div>
		<div>
			<div class="FormTitle">说 &nbsp;&nbsp;&nbsp;明</div>
			<div class="FormContentCol3FortextArea"><textarea name="stuResume_Text" id="stuResume_Text" class="textareaStyle"></textarea></div>
		</div>
        <div style="clear:both"></div>
		<div class="FormSubmitDiv"><input type="button" name="Submit" id="Submit" value="立即报名" style="cursor:pointer" onClick="return isok()"> &nbsp;<input type="reset" name="Reset" id="Reset" value="重新填写" style="cursor:pointer"></div>
	</form>
</div>