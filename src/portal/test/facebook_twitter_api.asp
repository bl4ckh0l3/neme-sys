<html>
    <head>
      <title>My Great Website</title>
    </head>
    <body>
    <!--<div id='fb-root'></div>
    <script src='http://connect.facebook.net/en_US/all.js'></script>
    <p><a onclick="postToFeed('passo direttamente il testo di prova del messaggio'); return false;">Post to Feed</a></p>
    <p id='msg'></p>

    <script> 
    FB.init({appId  : '207650472635802',
      status : true, // check login status
      cookie : true
    });

    function postToFeed(msgText) {

      // calling the API ...
      var obj = {
        method: 'feed',
        message: msgText,
        link: 'https://developers.facebook.com/docs/reference/dialogs/',
        //picture: 'http://fbrell.com/f8.jpg',
        //name: 'Facebook Dialogs',
        //caption: 'Reference Documentation',
        description: 'Using Dialogs to interact with users.'
      };

      function callback(response) {
        document.getElementById('msg').innerHTML = "Post ID: " + response['post_id'];
      }

      FB.ui(obj, callback);
    }
    </script>-->
    
    
    
      <div id="fb-root"></div>
      <script>(function(d, s, id) {
        var js, fjs = d.getElementsByTagName(s)[0];
        if (d.getElementById(id)) {return;}
        js = d.createElement(s); js.id = id;
        js.src = "//connect.facebook.net/en_US/all.js#xfbml=1";
        fjs.parentNode.insertBefore(js, fjs);
      }(document, 'script', 'facebook-jssdk'));</script>

      <div class="fb-comments" data-href="blackholenet.com" data-num-posts="2" data-width="500"></div>


      <br/><br/><br/>
      <script src="//platform.twitter.com/widgets.js" type="text/javascript"></script>
      <div>
      <a href="https://twitter.com/share" class="twitter-share-button" data-count="none">Tweet</a>
      </div>    
     </body>
 </html>


