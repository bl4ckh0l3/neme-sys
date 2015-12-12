/**********************************************   PRIMA FUNZIONE SLIDESHOW   **********************************************/


var pictures = new Array()

var t
var j = 0
var slideShowSpeed = 2000
var crossFadeDuration = 3

function runSlideShow(){
   if (document.all){
      document.images.SlideShow.style.filter="blendTrans(duration=2)"
      document.images.SlideShow.style.filter="blendTrans(duration=crossFadeDuration)"
      document.images.SlideShow.filters.blendTrans.Apply()      
   }
   document.images.SlideShow.src = pictures[j].src
   if (document.all){
      document.images.SlideShow.filters.blendTrans.Play()
   }
   j = j + 1
   if (j > (pictures.length-1)) j=0
   t = setTimeout('runSlideShow()', slideShowSpeed)
}

/**********************************************   FINE PRIMA FUNZIONE SLIDESHOW   **********************************************/


/**********************************************   SECONDA FUNZIONE SLIDESHOW   **********************************************/

/*
imgPath = new Array;
SiClickGoTo = new Array;

if (document.images){
	i0 = new Image;
	i0.src = '01.gif';
	SiClickGoTo[0] = "http://www.html.it";
	imgPath[0] = i0.src;
	i1 = new Image;
	i1.src = '02.gif';
	SiClickGoTo[1] = "http://www.html.it";
	imgPath[1] = i1.src;
	i2 = new Image;
	i2.src = '03.gif';
	SiClickGoTo[2] = "http://www.html.it";
	imgPath[2] = i2.src;
}
*/

a = 0;

function ejs_img_fx(img){
	if(img && img.filters && img.filters[0]){
		img.filters[0].apply();
		img.filters[0].play();
	}
}

function StartAnim(){
	if (document.images){
		document.images.SlideShow.style="filter:progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=100,Duration=1)";
		//document.write('<IMG SRC="'+base_path+'" BORDER=0 NAME=SlideShow style="filter:progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=100,Duration=1)">');
		defilimg()
	}
	/*else{
		document.write('<IMG SRC="'+base_path+'" BORDER=0>')
	}*/
}
	
/*function ImgDest(){
	document.location.href = SiClickGoTo[a-1];
}*/
	
function defilimg(){
	if (a == 3){
		a = 0;
	}
	
	if (document.images){
		ejs_img_fx(document.SlideShow)
		document.SlideShow.src = pictures[a];
		tempo3 = setTimeout("defilimg()",2000);
		a++;
	}
}

/**********************************************   FINE SECONDA FUNZIOE SLIDESHOW   **********************************************/


/**********************************************   TERZA FUNZIONE SLIDESHOW   **********************************************/


//var demoTabs;
function slideSwhowFading(){
	/*Event.observe(window, "load", function() {
		//Immagini
		var images = [
			"beach.jpg",
			"play.jpg",
			"bone.jpg",
			"snow.jpg",
			"sunrise.jpg"
		];*/

		new Widget.Fader("SlideShow", pictures);
	//});
}

/**********************************************   FINE TERZA FUNZIOE SLIDESHOW   **********************************************/
