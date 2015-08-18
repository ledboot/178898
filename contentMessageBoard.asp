<div class="gsjjText">
	<div class="messageBoardForm">
		<form name="frmgstbook" action="save.asp" method="post">
			<div>
				<span class="titleSpan">验 证 码：</span>
				<span class="inputSpan"><input class="inputStyle" name="Passcode" type="text"> <img src="inc/VerifyCode.asp" > * </span>
			</div>
			<div>
				<span class="titleSpan">您的名字：</span>
				<span class="inputSpan"><input class="inputStyle2" type="text" name="txtname"> * </span>
			</div>
			<div>
				<span class="titleSpan">您的电邮：</span>
				<span class="inputSpan"><input class="inputStyle2" type="text" name="txtemail"></span>
			</div>
			<div>
				<span class="titleSpan">主页网址：</span>
				<span class="inputSpan"><input class="inputStyle2" type="text" name="txthomepage" value="http://"></span>
			</div>
			<div>
				<span class="titleSpan">您的OICQ：</span>
				<span class="inputSpan"><input class="inputStyle2" type="text" name="txtoicq"></span>
			</div>
			<div>
				<span class="titleSpan">联系电话：</span>
				<span class="inputSpan"><input class="inputStyle2" type="text" name="tel"> *</span>
			</div>
			<div>
				<span class="titleSpan">表格颜色：</span>
				<span class="inputSpan">
					<select name="color" class="inputStyle3">
						<option class="optionStyle1" value="cfffca"></option>
						<option class="optionStyle2" value="d9ecff" selected></option>
						<option class="optionStyle3" value="ffe6f2"></option>
						<option class="optionStyle4" value="FFFBC1"></option>
						<option class="optionStyle5" value="EAEAEA"></option>
						<option class="optionStyle6" value="ECECFF"></option>
					</select>
				</span>
			</div>
			<div>
				<span class="titleSpan">选择表情：</span>
				<span class="inputSpan">
					<select name="sel_picname" onClick="theme_pic.src=sel_picname.value" class="inputStyle3">
						<option value="gs_images/gif/19.gif">请问</option>
						<option value="gs_images/gif/1.gif">哇，真高兴</option>
						<option value="gs_images/gif/2.gif">恐惧</option>
						<option value="gs_images/gif/3.gif">昏倒</option>
						<option value="gs_images/gif/4.gif">玩酷</option>
						<option value="gs_images/gif/5.gif">有趣</option>
						<option value="gs_images/gif/6.gif">滑稽</option>
						<option value="gs_images/gif/7.gif">不宜</option>
						<option value="gs_images/gif/8.gif">大笑</option>
						<option value="gs_images/gif/9.gif">喜欢</option>
						<option value="gs_images/gif/10.gif">嘻嘻,我也来参加</option>
						<option value="gs_images/gif/11.gif">伤心</option>
						<option value="gs_images/gif/12.gif">嘿，你看</option>
						<option value="gs_images/gif/13.gif">好极了</option>
						<option value="gs_images/gif/14.gif">嗯，还可以</option>
						<option value="gs_images/gif/15.gif">哇！厉害喔</option>
						<option value="gs_images/gif/16.gif">困倦</option>
						<option value="gs_images/gif/17.gif">无聊</option>
						<option value="gs_images/gif/18.gif">内容转贴</option>
						<option value="gs_images/gif/20.gif">新闻，消息</option>
						<option value="gs_images/gif/21.gif">嘿嘿,我有主意了</option>
						<option value="gs_images/gif/22.gif">不咋的</option>
						<option value="gs_images/gif/23.gif">大哭</option>
						<option value="gs_images/gif/24.gif">哇，原来如此</option>
						<option value="gs_images/gif/25.gif">还是不行</option>
						<option value="gs_images/gif/26.gif">这下惨了</option>
						<option value="gs_images/gif/27.gif">警告</option>
					</select> <img name="theme_pic" height="15" src="gs_images/GIF/19.gif" width="15">
				</span>
			</div>
			<div>
				<span class="titleSpan">您的留言：</span>
				<span class="inputSpan"><textarea name="txtContent" class="inputStyle4"></textarea> * </span>
			</div>
			<div class="Clearance">&nbsp;</div>
			<div class="submitButtonDiv">
				<input type="submit" name="cmdok" value="送出留言">&nbsp;&nbsp;<input type="reset" name="cmdreset" value="重新填写">&nbsp;&nbsp;<input type="button" name="cmdreset2" value="放弃留言" onClick="javascript:history.back()">
			</div>
		</form>
	</div>
</div>