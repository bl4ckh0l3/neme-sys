<script>
function sendMenuTips(numGer, actionUrl){
    if(actionUrl!="#"){
      document.menu_tips.gerarchia.value=numGer;
      document.menu_tips.action=actionUrl;
      document.menu_tips.submit();
    }
}
           
function ajaxLoadMenuTips(){
  var query_string = "gerarchia=<%=request("gerarchia")%>";
  //alert("query_string: "+query_string);

  $.ajax({
    async: true,
    type: "GET",
    cache: false,
    url: "<%=Application("baseroot")&"/public/layout/include/menutips_ajax.asp"%>",
    data: query_string,
    success: function(response) {
      //alert(response);
      //$("#loading-menu-tips").hide();
      $("#menuTips").empty();
      $("#menuTips").append(response);
      $("#menuTips").fadeIn(1000);
    },
    error: function() {
      //alert("errorrrrrrrrrr!");
      $("#menuTips").empty();
    }
  });
}

jQuery(document).ready(function(){
  ajaxLoadMenuTips();
});
</script>
<div id="menuTips">
<span id="loading-menu-tips"><img src="<%=Application("baseroot")&"/common/img/loading_icon2.gif"%>" hspace="0" vspace="0" border="0" width="15" height="15" align="center" alt="" style="padding-top:0px;padding-bottom:0px;"></span>
</div>
<form name="menu_tips" method="post" action="">
<input type="hidden" name="gerarchia" value="">
<input type="hidden" name="modelPageNum" value="1">
</form>