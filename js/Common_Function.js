// ��֤�ı����Ƿ�Ϊ��
function checkIsEmpty(Obj){
	var InputObj_Value = Obj.value;
	//alert(InputObj_Value);
	if(InputObj_Value.replace(/(^\s*)|(\s*$)/g, "") == ""){
		return true;
	}
	else{
		return false;	
	}
}
// ���ñ�
function resetFrm(Obj_ID){
	var frmObj = document.getElementById(Obj_ID);
	frmObj.reset();
	return false;
}
// checkboxԪ��ȫѡ�뷴ѡ
function QuanXuanAndReverseChoose(selectIdAll, selectId){
	var selectIdAll_Obj = document.getElementById(selectIdAll);
	var selectId_Obj = document.getElementsByName(selectId);
	//alert(selectId_Obj);
	if(selectIdAll_Obj.checked){
		for(i = 0;i < selectId_Obj.length;i ++){
			selectId_Obj[i].checked = true;
		}
	}
	else{
		for(i = 0;i < selectId_Obj.length;i ++){
			selectId_Obj(i).checked = false;
		}
	}
}
// ��ѭregָ���Ĺ�����֤�ı����֪�Ƿ���Ϲ���
function regInput(obj, reg, inputStr){
	var docSel	= document.selection.createRange()
	if (docSel.parentElement().tagName != "INPUT")	return false
	oSel = docSel.duplicate()
	oSel.text = ""
	var srcRange	= obj.createTextRange()
	oSel.setEndPoint("StartToStart", srcRange)
	var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
	return reg.test(str)
}
//�����ͣ��ʾ�㣬����뿪������
function showme(divId){
	var oSon = window.document.getElementById(divId);
	if (oSon == null) return;
	with(oSon){
		style.display = "block";
		style.pixelLeft = window.event.clientX + window.document.body.scrollLeft + 6;
		style.pixelTop = window.event.clientY + window.document.body.scrollTop + 9;
	}
}

function hidme(divId){
	var oSon = window.document.getElementById(divId);
	if (oSon == null) return;
	oSon.style.display="none";
}

function searchStringForURL(searchString, compart){
	var ss = location.href.substring(location.href.indexOf(searchString) + 1);
	SSArray = ss.split(compart);
	return SSArray;
}
//�����ı����Ĭ��ֵ
function setInputDefaultValue(objId, objValue, ComparativeValue){
	var inputObj = document.getElementById(objId);
	if(objValue == ComparativeValue){
		inputObj.value = "";
	}
	if(objValue.replace(/(^\s*)|(\s*$)/g, "") == ""){
		inputObj.value = ComparativeValue;
	}
}
//����ƶ���ѡ��ϣ�������ʽ
function fsecBoard(n, s, k, YClass, NClass, iframeAddress, showIfrmeId, moreLinkId, moreLinkAddress){ 
	for(i = s;i <= k;i ++){
		document.getElementById("fl0" + i).className = NClass; 
		document.getElementById("fbx0" + i).style.display = "none"; 
	} 
	document.getElementById("fl0" + n).className = YClass; 
	document.getElementById("fbx0" + n).style.display = "block";
	document.getElementById(moreLinkId).href = moreLinkAddress;
	if(n != 1){
		iframeAddress = iframeAddress + "&showDivId=fbx0" + n;
		document.getElementById(showIfrmeId).src = iframeAddress;
	}
}

function updateSubLmStyle(subLmImgTemp, subLmLinkTemp){
	var subLmImgObj = document.getElementById(subLmImgTemp);
	var subLmLinkObj = document.getElementById(subLmLinkTemp);
	if(subLmImgObj.className == "sidebarDivStyle8"){
		subLmImgObj.className = "sidebarDivStyle9";
	}
	else{
		subLmImgObj.className = "sidebarDivStyle8";
	}
	if(subLmLinkObj.className == "sidebarLinkStyle2"){
		subLmLinkObj.className = "sidebarLinkStyle3";
	}
	else{
		subLmLinkObj.className = "sidebarLinkStyle2";
	}
}

function checkSurveyQuestion(radioName, iframeID){
	var radioArrayObj = document.getElementsByName(radioName);
	var radioChecked = false;
	var checkedRadioVale = 0;
	var iframeObj = document.getElementById(iframeID);
	for(i = 0;i < radioArrayObj.length;i ++){
		if(radioArrayObj[i].checked){
			checkedRadioVale = radioArrayObj[i].value;
			radioChecked = true;
			break;
		}	
	}
	if(radioChecked){
		jumpURL = "processingSurveyQuestion.asp?id=" + checkedRadioVale;
		iframeObj.src = jumpURL;
	}
	else{
		alert("�Բ����㻹û�������κ�ѡ�񣬲����ύ��������ѡ��");
	}
}

function jsleft(lefts, leftn){ 
	var sl = lefts; 
	sl = sl.substring(0, leftn); 
	return sl; 
}
function jsright(rights, rightn){ 
	var sr = rights; 
	sr = sr.substring(sr.length - rightn, sr.length); 
	return sr; 
}
String.prototype.Trim = function() {
	var m = this.match(/^\s*(\S+(\s+\S+)*)\s*/);
	return (m == null) ? "" : m[1];
}
String.prototype.isMobile = function() {
	return (/^(?:13\d|15[2589]|18[79])-?\d{5}(\d{3}|\*{3})/.test(this.Trim()));
}
String.prototype.isTel = function(){
	//"���ݸ�ʽ: ���Ҵ���(2��3λ)-����(2��3λ)-�绰����(7��8λ)-�ֻ���(3λ)"
	//return (/^(([0\+]\d{2,3}-)?(0\d{2,3})-)?(\d{7,8})(-(\d{3,}))?/.test(this.Trim()));
	return (/^(([0\+]\d{2,3}-)?(0\d{2,3})-)(\d{7,8})(-(\d{3,}))?/.test(this.Trim()));
}

function PCAS() {
	this.SelP = document.getElementsByName(arguments[0])[0];
	this.SelC = document.getElementsByName(arguments[1])[0];
	this.SelA = document.getElementsByName(arguments[2])[0];
	this.DefP = this.SelA ? arguments[3] : arguments[2];
	this.DefC = this.SelA ? arguments[4] : arguments[3];
	this.DefA = this.SelA ? arguments[5] : arguments[4];
	this.SelP.PCA = this;
	this.SelC.PCA = this;
	this.SelP.onchange = function() { PCAS.SetC(this.PCA) };
	if (this.SelA) this.SelC.onchange = function() { PCAS.SetA(this.PCA) };
	PCAS.SetP(this)
};

PCAS.SetP = function(PCA) {
	var p_i = 0;
	for (i = 0; i < PCAN.length; i++) {
		//document.write(city_arr[i-1].substring(2,6)+"<br>");
		if (PCAN[i].substring(2, 6) == "0000") {
			PCAPV = PCAN[i].split('|')[0];
			PCAPT = PCAN[i].split('|')[1];
			PCA.SelP.options.add(new Option(PCAPT, PCAPV));
			if (PCA.DefP == PCAPV) PCA.SelP[p_i].selected = true
			p_i++;
		} 
	}
	PCAS.SetC(PCA);
};

PCAS.SetC = function(PCA) {
	PCA.SelC.length = 1;
	var c_i=0;
	var city1_str = PCA.SelP.value;
	var str_city1 = city1_str / 10000;
	//alert(str_city1);
	for (i = 0; i < PCAN.length; i++) {
		if (PCAN[i].substring(0, 2) == str_city1 && PCAN[i].substring(2, 6) != "0000" && PCAN[i].substring(4, 6) == "00") {
			PCACV = PCAN[i].split('|')[0];
			PCACT = PCAN[i].split('|')[1];
			PCA.SelC.options.add(new Option(PCACT, PCACV));
			if (PCA.DefC == PCACV) PCA.SelC[c_i+1].selected = true
			c_i++;
		}
	} if (PCA.SelA) PCAS.SetA(PCA)
};

PCAS.SetA = function(PCA) {
	PCA.SelA.length = 1;
	var a_i=0;
	var city2_str = PCA.SelC.value;
	var str_city2 = city2_str / 100;
	//alert(str_city1);
	for (i = 0; i < PCAN.length; i++) {
		if (PCAN[i].substring(0, 4) == str_city2 && PCAN[i].substring(4, 6) != "00") {
			PCAAV = PCAN[i].split('|')[0];
			PCAAT = PCAN[i].split('|')[1];
			PCA.SelA.options.add(new Option(PCAAT, PCAAV));
			if (PCA.DefA == PCAAV) PCA.SelA[a_i+1].selected = true
			a_i++;
		}
	}
}

/*�ж�������Ƿ�װ��Flash�ؼ�*/
function MM_CheckFlashVersion(reqVerStr, msg){
	with(navigator){
		var isIE  = (appVersion.indexOf("MSIE") != -1 && userAgent.indexOf("Opera") == -1);
		var isWin = (appVersion.toLowerCase().indexOf("win") != -1);
		if (!isIE || !isWin){  
			var flashVer = -1;
			if (plugins && plugins.length > 0){
				var desc = plugins["Shockwave Flash"] ? plugins["Shockwave Flash"].description : "";
				desc = plugins["Shockwave Flash 2.0"] ? plugins["Shockwave Flash 2.0"].description : desc;
				if (desc == ""){
					flashVer = -1;
				}
				else{
					var descArr = desc.split(" ");
					var tempArrMajor = descArr[2].split(".");
					var verMajor = tempArrMajor[0];
					var tempArrMinor = (descArr[3] != "") ? descArr[3].split("r") : descArr[4].split("r");
					var verMinor = (tempArrMinor[1] > 0) ? tempArrMinor[1] : 0;
					flashVer =  parseFloat(verMajor + "." + verMinor);
				}
			}
			// WebTV has Flash Player 4 or lower -- too low for video
			else if (userAgent.toLowerCase().indexOf("webtv") != -1){
				flashVer = 4.0;
			}
			var verArr = reqVerStr.split(",");
			var reqVer = parseFloat(verArr[0] + "." + verArr[2]);
			if (flashVer < reqVer){
				if (confirm(msg)){
					window.location = "http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash";
				}
			}
		}
	} 
}

// �滮���ԺƵ�� �����Ӳ˵�Begin
var def="1";
function mover(object){
	//���˵�
	var mm=document.getElementById("m_"+object);
	mm.className="m_li_a";
	//��ʼ���˵�������Ч��
	if(def!=0){
		var mdef=document.getElementById("m_"+def);
		mdef.className="m_li";
	}
	//�Ӳ˵�
	var ss=document.getElementById("s_"+object);
	ss.style.display="block";
	//��ʼ�Ӳ˵�������Ч��
	if(def!=0){
		var sdef=document.getElementById("s_"+def);
		sdef.style.display="none";
	}
}
	
function mout(object){
	//���˵�
	var mm=document.getElementById("m_"+object);
	mm.className="m_li";
	//��ʼ���˵���ԭЧ��
	if(def!=0){
		var mdef=document.getElementById("m_"+def);
		mdef.className="m_li_a";
	}
	//�Ӳ˵�
	var ss=document.getElementById("s_"+object);
	ss.style.display="none";
	//��ʼ�Ӳ˵���ԭЧ��
	if(def!=0){
		var sdef=document.getElementById("s_"+def);
		sdef.style.display="block";
	}
}
// �滮���ԺƵ�� �����Ӳ˵�End

//վ������Begin
function searchInfo(searchWordStrId){
	var searchWordStrObj = document.getElementById(searchWordStrId);
	if(checkIsEmpty(searchWordObj)){
		alert("��������Ч�������ؼ��ʣ�")
		return false;
	}
}
//վ������End
