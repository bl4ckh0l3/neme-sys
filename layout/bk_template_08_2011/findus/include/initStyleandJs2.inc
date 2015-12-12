
<link rel="stylesheet" href="<%=Application("baseroot") & "/public/layout/css/stile.css"%>" type="text/css">
<%if not(isNull(strCSS)) ANd not(strCSS = "") then%>
<link rel="stylesheet" href="<%=Application("baseroot") & strCSS%>" type="text/css">
<%end if%>
<script type="text/JavaScript" src="<%=Application("baseroot") & "/common/js/jsDraw2D_Uncompressed.js"%>"></script> 
<script type="text/javascript" src="<%=Application("baseroot")&"/common/js/dom-drag.js"%>"></script>
<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAA_fC5YbH2g1ksyJQVZFp4xxQvEvV86WDzDA3NIhU79VSEjDOoDxQrk_kJ_0noEdAFGbBsYCQJi1rSuQ&sensor=false" type="text/javascript"></script>
<script type="text/javascript">
    function initialize() {
      if (GBrowserIsCompatible()) {
        var map = new GMap2(document.getElementById("map_canvas"));
        map.setCenter(new GLatLng(37.4419, -122.1419), 13);
        map.setUIToDefault();
      }
    }    
</script>