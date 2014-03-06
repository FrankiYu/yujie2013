// JavaScript Document

// Developed by Franki Yu


window.addEventListener('load', function(){  
	setTimeout(scrollTo, 0, 0, 1);  
}, false);  

var isWebApp = navigator.standalone;
if(isWebApp) {
	$("#navbar").before('<div class="statusbar"></div>');
	$(".navbar-fixed-top").css("top","20px");
}

function completed() {
	$(".panel").show();
	
	$("#loading").fadeOut();
	
	$(".profile").removeClass('loading');

	$('#photograph').carousel({
		 interval: 5000
		});
	$('#showcase').carousel({
		 interval: 8000
		});	
};

function email_hover() {
	$('.email').addClass("hover");
}
function email_out() {
    $('.email').removeClass("hover");
}
function sns_hover() {
	$('sns').addClass("hover");
	$('at').addClass("hover");
	$('.email').addClass("fadeout");
	$('.web').addClass("fadeout");
}
function sns_out() {
    $('sns').removeClass("hover");
	$('at').removeClass("hover");
	$('.email').removeClass("fadeout");
	$('.web').removeClass("fadeout");
}

// --- icon hover --- ///
$('#icon-facebook').mouseover(function(){
	email_hover();
	$(this).tooltip('show');
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-facebook').mouseout(function(){
	email_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});
$('#icon-twitter').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');	
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-twitter').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-google').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');	
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-google').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-qzone').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');	
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-qzone').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-sinaweibo').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');	
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-sinaweibo').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-tencentweibo').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-tencentweibo').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-wechat').mouseover(function(){
	sns_hover();
	$(this).tooltip('show');
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-wechat').mouseout(function(){
	sns_out();
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

$('#icon-qq').mouseover(function(){
	$(this).tooltip('show');	
	$('.sns-icon').not(this).addClass("sns-fadeout");	
});

$('#icon-qq').mouseout(function(){
	$('.sns-icon').not(this).removeClass("sns-fadeout");	
});

// --- text hover --- ///

$('email').mouseover(function(){
    $('.email').addClass("hover");
	$('email').tooltip('show');
	$('.sns-icon').not("#icon-facebook").addClass("sns-fadeout");	
});
$('email').mouseout(function(){
    $('.email').removeClass("hover");
	$('.sns-icon').not("#icon-facebook").removeClass("sns-fadeout");	
	
});
$('at').mouseover(function(){
    $('sns').addClass("hover");
	$('at').addClass("hover");
	$('.email').addClass("fadeout");
	$('.web').addClass("fadeout");
	$('sns').tooltip('show');
	$('#icon-facebook').addClass("sns-fadeout");
	$('#icon-qq').addClass("sns-fadeout");
});
$('at').mouseout(function(){
    $('sns').removeClass("hover");
	$('at').removeClass("hover");
	$('.email').removeClass("fadeout");
	$('.web').removeClass("fadeout");
	$('sns').tooltip('hide');
	$('#icon-facebook').removeClass("sns-fadeout");
	$('#icon-qq').removeClass("sns-fadeout");
});
$('sns').mouseover(function(){
    $('sns').addClass("hover");
	$('at').addClass("hover");
	$('.email').addClass("fadeout");
	$('.web').addClass("fadeout");
	$('sns').tooltip('show');
	$('#icon-facebook').addClass("sns-fadeout");
	$('#icon-qq').addClass("sns-fadeout");
	
});
$('sns').mouseout(function(){
    $('sns').removeClass("hover");
	$('at').removeClass("hover");
	$('.email').removeClass("fadeout");
	$('.web').removeClass("fadeout");
	$('sns').tooltip('hide');
	$('#icon-facebook').removeClass("sns-fadeout");
	$('#icon-qq').removeClass("sns-fadeout");
});
$('web').mouseover(function(){
    $('web').addClass("hover");
	$('sns').addClass("hover");
	$('at').addClass("fadeout");
	$('.email').addClass("fadeout");
	$('web').tooltip('show');
});
$('web').mouseout(function(){
    $('web').removeClass("hover");
	$('sns').removeClass("hover");
	$('at').removeClass("fadeout");
	$('.email').removeClass("fadeout");
});
