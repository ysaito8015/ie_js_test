var ie = WScript.CreateObject("InternetExplorer.Application");
  // IE起動
  ie.Visible = true;
  ie.Navigate( 'https://www.google.co.jp/' );
  while ( (ie.busy) || (ie.redystate != 4) ) {
    WScript.Sleep(100);
  }
  //waitIE( ie );
  // 検索キーワードを入力
  ie.document.getElementById("q").value = "test";
  WScript.Sleep( 1000 );
  // 検索ボタンクリック
  ie.document.all("btnG").click();
  WScript.Sleep( 1000 );
  //waitIE( ie );
  // 制御を破棄
  ie.Quit();
  ie = null;


// IEがビジー状態の間待ちます
//function waitIE( ie ) {
//  while( ( ie.Busy ) || ( ie.readystate != 4 ) ) {
//    WScript.Sleep( 100 );
//  }
//  WScript.Sleep( 1000 )
//}
