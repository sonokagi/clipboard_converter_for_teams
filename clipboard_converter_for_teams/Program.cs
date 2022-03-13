using System;
using System.Collections.Generic;

namespace clipboard_converter_for_teams
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // バージョン表示
            print_version();

            // クリップボード変換
            ClipboardConverterCollection.Execute();

            // 少し待つ。コンソールがすぐ消えると動作が分かり難いので
            System.Threading.Thread.Sleep(500);
        }

        static void print_version()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly().GetName();
            var ver = assembly.Version;
            Console.WriteLine($"{assembly.Name} v{ver.Major}.{ver.Minor}");
        }
    }

    public static class ClipboardConverterCollection
    {
        public static void Execute()
        {
            List<ClipboardConverter> converters = new List<ClipboardConverter>() {
                new LinkToPostConverter(),  // 「投稿へのリンク」を変換する
                new UrlTextConverter()      // 「テキストのURL」を変換する
            };

            foreach (ClipboardConverter c in converters)
            {
                if (c.Execute() == ClipboardConverter.Result.SUCCESS) return;
            }
        }
    }

    public abstract class ClipboardConverter
    {
        public enum Result { SUCCESS, SKIP };
        public abstract Result Execute();
        protected void outputResult(string before, string after)
        {
            Console.WriteLine("--- [before] ---");
            Console.WriteLine(before);
            Console.WriteLine("--- [after] ---");
            Console.WriteLine(after);
        }
    }

    public class LinkToPostConverter : ClipboardConverter
    {
        // Teamsのチャットで「リンクをコピー」を行うと、クリップボードにHtml形式で下記のようなデータが格納される
        // ここから必要なデータ(★で示した部分)を抜き出し、さらに「投稿者名: 」を削除して、クリップボードに上書きする
        //-----------------------------------------
        //Version:0.9
        //StartHTML:000000xxxx
        //EndHTML:000000xxxx
        //StartFragment:000000xxxx
        //EndFragment:000000xxxx
        //<html>
        //<body>
        //<!--StartFragment--><div itemprop="teams-copy-link">
        //★ここから★<a href="https://teams.microsoft.com/l/・・・" title="・・・">投稿者名: 投稿内容</a>★ここまで★
        //</div><div itemprop="teams-copy-link">Team名 / チャンネル名 で 2022年x月x日 xx:xx:xx に投稿しました</div><div>&nbsp;</div>
        //<!--EndFragment-->
        //</body>
        //</html>
        //-----------------------------------------

        public override Result Execute()
        {
            // 変換可否の事前チェック
            if (!canExecute()) return Result.SKIP;

            // クリップボードからデータをHTML/TEXT形式で取得
            string html = ClipboardWrapper.GetHtmlFragmentPart();
            string text = ClipboardWrapper.GetText();

            // 「投稿へのリンク部分」を抽出後に投稿者名を削除。異常時は抜ける
            string link = removeNameOfPoster(extractLinkPart(html));
            if (string.IsNullOrEmpty(link)) return Result.SKIP;

            // クリップボードに「投稿へのリンク部分」をHtml形式で設定
            // 必須ではないが、テキスト形式で元のクリップボードと同一データも設定しておく
            ClipboardWrapper.SetHtmlFragmentPartAndText(link, text);

            // 変換結果を出力
            outputResult(html, link);

            return Result.SUCCESS;
        }

        private bool canExecute()
        {
            // クリップボードにHTMLとTEXT形式のデータが両方存在し、
            // HTML形式データのFragment部分に"teams-copy-link"の文字列があれば、投稿へのリンクが格納されていると判断

            // TEXT形式 もしくは HTML形式のデータがなければ抜ける
            if (!ClipboardWrapper.ContainsText()) return false;
            if (!ClipboardWrapper.ContainsHtml()) return false;

            // HTML形式データのFragment部分を抜き出し
            // "teams-copy-link"の文字列があれば、投稿へのリンクが格納されていると判断
            string fragment_part = ClipboardWrapper.GetHtmlFragmentPart();
            return fragment_part.Contains("teams-copy-link");
        }

        private string extractLinkPart(string html)
        {
            // Htmlデータから投稿へのリンク部分「<a href=」～「</a>」の文字列を探す
            string start_keyword = "<a href=";              // 開始キーワード
            string end_keyword = "</a>";                    // 終了キーワード
            int start_index = html.IndexOf(start_keyword);  // 開始キーワードの検出位置
            int end_index = html.IndexOf(end_keyword);      // 終了キーワードの検出位置

            // 正常にキーワードが見つかった場合、各インデックスは0より大きくなるはず
            if (start_index > 0 && end_index > 0)
            {
                // 「投稿へのリンク部分」を返す
                int link_length = (end_index + end_keyword.Length) - start_index;
                return html.Substring(start_index, link_length);
            }
            else
            {
                return string.Empty;
            }
        }

        private string removeNameOfPoster(string link)
        {
            // 入力データは <a href="～">投稿者名: ～</a> を前提とする
            // このため「">」投稿者名「: 」の文字列を探す
            string start_keyword = "\">";                   // 開始キーワード
            string end_keyword = ": ";                      // 終了キーワード
            int start_index = link.IndexOf(start_keyword);  // 開始キーワードの検出位置
            int end_index = link.IndexOf(end_keyword);      // 終了キーワードの検出位置

            // 正常にキーワードが見つかった場合、各インデックスは0より大きくなるはず
            if (start_index > 0 && end_index > 0)
            {
                // 「投稿者名」を取り除いて返す
                string part_before_name = link.Substring(0, start_index + start_keyword.Length);
                string part_after_name = link.Substring(end_index + end_keyword.Length);
                return part_before_name + part_after_name;
            }
            else
            {
                return string.Empty;
            }
        }
    }

    public class UrlTextConverter : ClipboardConverter
    {
        // Teams上のファイルやフォルダで「リンクをコピー」すると、クリップボードにText形式でSharePointのURLが格納される
        // これをHTMLタグのリンクに変換し、Html形式でクリップボードに上書きする

        public override Result Execute()
        {
            // 変換可否の事前チェック
            if (!canExecute()) return Result.SKIP;

            // クリップボードからデータをTEXT形式で取得
            string text = ClipboardWrapper.GetText();

            // HTMLタグで「URLへのリンク」を作る。文字列は固定
            string link = "<a href=\"" + text + "\"><b><i>HERE!</i></b></a>";

            // クリップボードに「URLへのリンク」をHtml形式で設定
            // 必須ではないが、テキスト形式で元のクリップボードと同一データも設定しておく
            ClipboardWrapper.SetHtmlFragmentPartAndText(link, text);

            // 変換結果を出力
            outputResult(text, link);

            return Result.SUCCESS;
        }

        private bool canExecute()
        {
            // クリップボードにTEXT形式データがあり
            if (ClipboardWrapper.ContainsText())
            {
                // そのデータが「http」から始まっていたら、URLと判断する
                if (ClipboardWrapper.GetText().StartsWith("http")) return true;
            }

            return false;
        }
    }
}