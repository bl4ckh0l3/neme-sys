<!-- #include virtual="/common/include/IncludeObjectList.inc" -->
<html>
<head>
<!-- #include virtual="/common/include/initCommonJs.inc" -->
<style>
/*.droptargetsx, .droptargetdx { list-style-type: none; margin: 0; float: left; margin-right: 10px; background: #eee; padding: 5px; width: 143px;height: 350px;}
.droptargetsx li, .droptargetdx li { margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; }*/
</style>
<script>
var resultTargetCatList ="1|2|3|4|5|";
var resultTargetLangList ="7|4|2|9|8|";
	
function manageTargetWidget(resultList, idlist_sx, idlist_dx) {

	alert("resultList start:"+resultList);
	

	$( "#"+idlist_sx ).sortable({
		connectWith: "ul"
		,receive: function(event, ui) {
			//alert("receive sx - li.id: "+ui.item.attr("id"));
			//alert("resultList start:"+resultList);
			
			resultList+=ui.item.attr("id")+"|";	
			alert("resultList finish:"+resultList);
		}
		,remove: function(event, ui) {
			//alert("remove sx - li.id: "+ui.item.attr("id"));
			resultList=resultList.replace(ui.item.attr("id")+"|","");
			alert(resultList);
		}
	}).disableSelection();
	
	$( "#"idlist_dx ).sortable({
		connectWith: "ul"
	}).disableSelection();
}

$(document).ready(function() {
	manageTargetWidget(resultTargetCatList,"snaptarget_sx","snaptarget_dx");
});
</script>
</head>
<body>
	
	
<ul id="snaptarget_sx" style="list-style-type: none; margin: 0; float: left; margin-right: 10px; background: #eee; padding: 5px; width: 143px;height: 350px;">
	<li class="ui-state-default" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="1">Can be dropped..</li>
	<li class="ui-state-default" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="2">..on an empty list</li>
	<li class="ui-state-default" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="3">Item 3</li>
	<li class="ui-state-default" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="4">Item 4</li>
	<li class="ui-state-default" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="5">Item 5</li>
</ul>

<ul id="snaptarget_dx" style="list-style-type: none; margin: 0; float: left; margin-right: 10px; background: #eee; padding: 5px; width: 143px;height: 350px;">
	<li class="ui-state-highlight" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="6">Can be dropped..</li>
	<li class="ui-state-highlight" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="7">..on an empty list</li>
	<li class="ui-state-highlight" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="8">Item 3</li>
	<li class="ui-state-highlight" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="9">Item 4</li>
	<li class="ui-state-highlight" style="margin: 5px; padding: 5px; font-size: 1.2em; width: 120px; cursor:move;" id="10">Item 5</li>
</ul>



</body>
</html>
