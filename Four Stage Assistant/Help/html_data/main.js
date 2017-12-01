var width, height;
if (document.body && document.body.offsetWidth){
	width = document.body.offsetWidth;
	height = document.body.offsetHeight;
}
if (document.compatMode=='CSS1Compat' && document.documentElement && document.documentElement.offsetWidth ){
	width = document.documentElement.offsetWidth;
	height = document.documentElement.offsetHeight;
}
if (window.innerWidth && window.innerHeight){
	width = window.innerWidth;
	height = window.innerHeight;
}
width = (width-1280)/2;
height = (height-720)/2;
function setMargin(){
	document.getElementById("content").style.marginLeft =  width + "px";
	document.getElementById("content").style.marginTop =  height + "px";
}
function switchLanguage(show, hide){
	document.getElementById("frame_" + show).style.visibility = "visible";
	document.getElementById("frame_" + hide).style.visibility = "hidden";
	document.getElementById("menu_" + show).style.visibility = "visible";
	document.getElementById("menu_" + hide).style.visibility = "hidden";
	document.getElementById("link_" + show).style.color = "#CACF43";
	document.getElementById("link_" + hide).style.color = "#3D3D3D";
	switchMenuDivs("main", "main_" + show);
}
function switchMenuDivs(source, show){
	var divs;
	if (source == "main")
		divs = new Array("main_pl", "kontakt_pl", "main_en", "about_en", "offer_en", "products_en", "contact_en");
	if (source == "maplayer")
		divs = new Array("about", "video");
	for (var i in divs){
		document.getElementById(divs[i]).style.visibility = (divs[i] == show) ? "visible" : "hidden";
		document.getElementById("link_" + divs[i]).style.color = (divs[i] == show) ? "#CACF43" : "#3D3D3D";
	}
}
function initMain(){
	setMargin();
	switchLanguage('en', 'pl');
	switchMenuDivs('main', 'main_en');
}
function initMapLayer(){
	setMargin();
	switchMenuDivs('maplayer', 'about');
}