var ws =WScript.CreateObject("WScript.Shell");
var ie = WScript.CreateObject("InternetExplorer.Application");
ie.Navigate( "http://www.yahoo.co.jp" );
ie.Visible = true;
while ( (ie.busy) || (ie.readystate != 4) ) {
  WScript.Sleep(1000);
}
// 検索キーワードを入力
//ie.document.all.Item("q").Value = 'Javascript';
var srchtxt = ie.document.all("srchtxt");
srchtxt.focus();
srchtxt.Value = "Javascript";
WScript.echo(srchtxt.Value);
//WScript.echo(ie.document.all("p").length);
WScript.Sleep( 1000 );
// 検索ボタンクリック
ie.document.all("srchbtn").click();
WScript.Sleep( 1000 );
// 制御を破棄
ie.Quit();
ie = null;


