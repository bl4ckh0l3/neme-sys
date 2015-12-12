<link rel="shortcut icon" type="image/x-icon" href="/favicon.ico">

<link rel="stylesheet" href="<%=Application("baseroot") & "/common/css/jquery-ui-latest.custom.css"%>" type="text/css">
<!-- codice per gestire un menù orizzontale con jquery al posto del menù classico verticale a sinistra -->
<link rel="stylesheet" type="text/css" href="<%=Application("baseroot") & "/public/layout/css/jqueryslidemenu.css"%>" />
<!--[if lte IE 7]>
<style type="text/css">
html .jqueryslidemenu{height: 1%;} /*Holly Hack for IE7 and below*/
</style>
<![endif]-->

<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-latest.min.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery-ui-latest.custom.min.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.google-analytics.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.ae.image.resize.min.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.form.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/javascript_global.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jqueryslidemenu.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/jquery.timers.js"%>"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/highcharts.js"%>"></script>

<!-- carico l'editor html semplificato CLEditor -->
<link rel="stylesheet" type="text/css" href="<%=Application("baseroot")&"/cleditor/jquery.cleditor.css"%>" />      
<script type="text/javascript" src="<%=Application("baseroot")&"/cleditor/jquery.cleditor.js"%>"></script>

<%if(Trim(Application("googlemaps_key"))<>"")then%>
<!--  ****************************************** INTEGRAZIONE GOOGLEMAP API ****************************************** -->
<script src="https://maps.googleapis.com/maps/api/js?key=<%=Trim(Application("googlemaps_key"))%>&amp;sensor=false&amp;libraries=drawing,geometry" type="text/javascript"></script>
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/markerclusterer_compiled.js"%>"></script>
 <%end if%>

 <%if(Trim(Application("analytics_account"))<>"")then%>
<!--  ******************************************** START: SCRIPT TRACKING ANALYTICS ******************************************** -->
<script type="text/javascript">
  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', '<%=Trim(Application("analytics_account"))%>']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();
</script>
<!--  ******************************************** END: SCRIPT TRACKING ANALYTICS ******************************************** -->
 <%end if%>