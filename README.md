# clipboard_converter_for_teams
Microsoft Teams 向けのクリップボード変換ツールです。  
クリップボードに格納された投稿やフォルダへのリンクを短縮します。使い方は下記です。  
1. Teams上の投稿やファイルで「リンクのコピー」を選択し、クリップボードにリンクをコピーする
1. このプログラムを実行し、クリップボードの内容を更新する
1. Teams上でペーストを行うと、短縮された内容がペーストされる

### 機能
プログラムを実行すると、クリップボードに格納されたデータを判別し、下記のいずれかの１つを行います  

変換1: 「投稿へのリンク」を短くする
- before(Html形式)  
  >[投稿者名: 投稿内容](https://github.com/sonokagi/clipboard_converter_for_teams)  
  チーム名/チャンネル名 で XXXX年X月X日 XX:XX:XX に投稿しました
- after(Html形式)  
  >[投稿内容](https://github.com/sonokagi/clipboard_converter_for_teams)  

変換2: 「フォルダへのリンク」をHTMLタグによるリンクに変換する  
- before(Text形式)  
  >[https://～.sharepoint.com/～/～/～/～～～～～～～～～](https://github.com/sonokagi/clipboard_converter_for_teams)
- after(Html形式)  
  >[**_HERE!_**](https://github.com/sonokagi/clipboard_converter_for_teams)

### 使用前の準備
このプログラムは、Windowsのショートカットキーに割り当てて使います。まず、スタートメニューへのショートカット追加と、ショートカットキーの割り当てを行ってください。  
参考＞[Windows10 ショートカットキーの設定方法｜追加・変更など](https://litera.app/blog/windows-setting-shortcut/)

### 開発環境
- Windows 10
- Visual Studio 2019
- .Net Framework 4.7.2

### 利用しているソフトウェア
- ClipboardHelper.cs - The MIT License (MIT) Copyright (c) 2014 Arthur Teplitzki  
  >Helper to encode and set HTML fragment to clipboard.  
  >See http://theartofdev.com/2014/06/12/setting-htmltext-to-clipboard-revisited/