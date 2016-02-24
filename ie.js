<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>ie test</title>
</head>
<body>

<script>
  (function() {
    'use strict';
  use_ie();
  
  function use_ie() {
      // IE起動
      var ie = WScript.CreateObject("InternetExplorer.Application")
      ie.Navigate( "http://www.google.co.jp/" );
      ie.Visible = true;
      waitIE( ie );
      // 検索キーワードを入力
      ie.document.getElementById("q").value = "test";
      WScript.Sleep( 100 );
      // 検索ボタンクリック
      ie.document.all("btnG").click();
      waitIE( ie );
      // 制御を破棄
      ie.Quit();
      ie = null;
  }
  
  
  // IEがビジー状態の間待ちます
  function waitIE( ie ) {
      while( ( ie.Busy ) || ( ie.readystate != 4 ) ) {
          WScript.Sleep( 100 );
      }
      WScript.Sleep( 1000 )
  }
  });
</script>
</body>
</html>
