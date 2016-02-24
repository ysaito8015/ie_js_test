// ------------------ jQueryでスクレイピング ------------------



// 現在のページ内でjQueryを有効化する（スクレイピングしやすいように）
function ie_inject_jquery( ie ){

	// 既に$があるか
	//if( ie.document.$ ) return;

	// script要素を新規作成
	var elem_head = ie.document.getElementsByTagName('head')[0];
	var elem_script = ie.document.createElement('script');
	elem_script.src =  "https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js";
		// https://developers.google.com/speed/libraries/devguide#jquery
	
	// ロード完了時のイベントを定義
	var load_complete = false;
	elem_script.onload = function(){
		load_complete = true;
	};
	
	// HEADにscriptタグを追加
	elem_head.appendChild( elem_script );
	
	// scriptのロード完了まで待つ
	while( load_complete ){
		WScript.Sleep( 500 );
	}
	
	// $を評価可能になるまで待つ
	while( ! ie.document.parentWindow.$ ){
		WScript.Sleep( 500 );
	}

	// WSHのグローバルで参照を定義
	$ = ie.document.parentWindow.$;
	//log("IEにjQuery注入完了");
	
	return;
}

/*
// IE上でjQueryを使う
function $( s ){
	return ie.document.parentWindow.$( s, ie.document );
}

↑ $.trim など様々な関数に対応したいので，
この関数は不要。

*/


// 遷移先の全ページ内でjQueryを有効化する
always_enable_JQuery = false;
function ie_enableJQuery(){
	always_enable_JQuery = true;
}



// ------------------ IEの基本操作 ------------------



// IE起動
function getIE()
{
	var ie = WScript.CreateObject("InternetExplorer.Application")
	ie.Visible = true;
	ie_goto_url( ie, "http://www.google.co.jp/" );
		//log("ブラウザでのアクセスを開始します。");
	
	return ie;
}


// IEがビジー状態の間待ちます
function ie_wait_while_busy( ie, _url )
{   
	var timeout_ms      = 45 * 1000;
	var step_ms         = 100;
	var total_waited_ms = 0;
	
	while( ( ie.Busy ) || ( ie.readystate != 4 ) )
	{
		WScript.Sleep( step_ms );
		
		// タイムアウトか？
		total_waited_ms += step_ms;
		if( total_waited_ms >= timeout_ms )
		{
			/*log(
				"警告：タイムアウトのため，リロードします。("
				+ ie.LocationURL
					// http://blog.livedoor.jp/programlog/archives/298228.html
				+ ")"
			);*/
			
			// どこかに移動中なら，そこへの移動を再試行
			if( _url )
			{
				//log( _url + "への遷移を再試行");
				ie_goto_url( ie, _url );
			}
			else
			{
				log( "リロード中");
				
				// 移動先が明示されていなければリロード
				ie.document.location.reload( true );
				ie_wait_while_busy( ie );
			}
			
			break;
		}
	}

	WScript.Sleep( 1000 )
}
	// http://d.hatena.ne.jp/language_and_engineering/20100310/p1
	// http://d.hatena.ne.jp/language_and_engineering/20100403/p1


// ページを移動
function ie_goto_url( ie, url ){
	//log("アクセスします：" + url);
	ie.Navigate( url );
	ie_wait_while_busy( ie, url );
	//log("ページを開きました。");
	
	// IEで常にjQueryを使うか
	if( always_enable_JQuery ){
		ie_inject_jquery( ie );
	}
}


// デバッグ用
function log(s){
	WScript.Echo(s);
}



// ------------------メイン処理 ------------------



// IE起動
var ie = getIE();

// 今後の遷移先の全ページ内でjQueryによるスクレイピングを有効化
ie_enableJQuery();

// 遷移
ie_goto_url( ie, "http://www.yahoo.co.jp/" );

// スクレイピング実行
log(
	$("div#topicsfb ul.emphasis li a") // jQueryセレクタが使える
		.map(function(){ 
			return $(this).html(); 
		})
		.get()
		.join("\n")
);
