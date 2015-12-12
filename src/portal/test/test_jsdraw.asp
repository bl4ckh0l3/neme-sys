<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<script type="text/JavaScript" src="/common/js/jsDraw2D_Uncompressed.js"></script> 
</head>
<body>

<div id="canvas" style="overflow:hidden;position:relative;width:600px;height:370px;border:#999999 1px solid;"></div>


<script type="text/JavaScript">


	var points=new Array();
	
	//Create jsGraphics object
	var gr = new jsGraphics(document.getElementById("canvas"));
	gr.setCoordinateSystem("cartecian");
	gr.setOrigin(new jsPoint(20,350));
	gr.setScale(3);
	gr.showGrid(10, true);

	//Create jsColor object
	var col = new jsColor("black");
	var col2 = new jsColor("red");
	var col3 = new jsColor("blue");

	//Create jsPen object
	var pen = new jsPen(col,2);
	var pen2 = new jsPen(col2,2);
	var pen3 = new jsPen(col3,2);
	
	//Create jsFont object
	var font = new jsFont("verdana", "normal", 10, "normal", "small-cups");

	//Draw filled circle with pt1 as center point and radius 30.
	gr.fillCircle(col,new jsPoint(30,30),20);
	//You can also code with inline object instantiation like below
	gr.drawLine(pen2,new jsPoint(30,30),new jsPoint(50,30));    
	gr.drawLine(pen2,new jsPoint(30,30),new jsPoint(20,13));  
	gr.drawLine(pen2,new jsPoint(30,30),new jsPoint(25,49));     
	
	
	gr.drawEllipse(pen, new jsPoint(80,30), 40, 40);	
	gr.drawArc(pen, new jsPoint(80,30), 40,40, 0, 30)
	gr.fillArc(col, new jsPoint(80,30), 40, 40, 0, 30)
	gr.drawArc(pen2, new jsPoint(80,30), 40,40, 30, 150)
	gr.fillArc(col2, new jsPoint(80,30), 40, 40, 30, 150)
	gr.drawArc(pen3, new jsPoint(80,30), 40,40, 180, 180)
	gr.fillArc(col3, new jsPoint(80,30), 40, 40, 180, 180)
	
	gr.drawText("elemento 1", new jsPoint(80,55), font, col, 50, "left")
	gr.drawText("elemento 2", new jsPoint(102,35), font, col, 50, "left")
	gr.drawText("elemento 3", new jsPoint(80,10), font, col, 50, "left")


	//Draw a Curve
	var points=new Array(new jsPoint(10,70),
					    new jsPoint(30,80),
					    new jsPoint(55,100),
					    new jsPoint(65,110),
					    new jsPoint(75,75),
					    new jsPoint(85,90),
					    new jsPoint(95,110),
					    new jsPoint(105,95),
					    new jsPoint(125,85),
					    new jsPoint(135,100),
					    new jsPoint(150,76),
					    new jsPoint(160,89),
					    new jsPoint(180,80)
						);
	for(var i=0;i<points.length;i++){
		drawPoint(points[i].x,points[i].y);
	}
	gr.drawCurve(pen, points);
	//gr.drawClosedCurve(pen, points);
	//gr.fillClosedCurve(col, points);
	
	
	
	//Draw a Line between 2 points
	//gr.drawLine(pen,new jsPoint(65,80),new jsPoint(100,80));
	
	
	// disegno un grafico a barre 
	gr.drawRectangle(pen, new jsPoint(140,30), 5, 30);
	gr.fillRectangle(col, new jsPoint(140,30), 5, 30);    
	
	gr.drawRectangle(pen, new jsPoint(150,45), 5, 45);
	gr.fillRectangle(col, new jsPoint(150,45), 5, 45);  
	
	gr.drawRectangle(pen, new jsPoint(160,50), 5, 50);
	gr.fillRectangle(col, new jsPoint(160,50), 5, 50);  
	
	gr.drawRectangle(pen, new jsPoint(170,40), 5, 40);
	gr.fillRectangle(col, new jsPoint(170,40), 5, 40); 
	
	gr.drawRectangle(pen, new jsPoint(180,60), 5,60);
	gr.fillRectangle(col, new jsPoint(180,60), 5, 60);  


	function drawPoint(x,y){
		gr.fillRectangle(new jsColor("green"),new jsPoint(x-1,y+1),2,2);
	}
</script>
</body>
</html>