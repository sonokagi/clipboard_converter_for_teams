using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            // 「投稿へのリンク」を変換する
            if (link_to_post_is_stored_in_clipboard())
            {
                convert_link_to_post();
            }
            // 「テキストのURL」を変換する
            else if (url_text_is_stored_in_clipboard())
            {
                convert_url_text();
            }
            // その他
            else
            {
                // 処理無し
            }
        }

        static bool link_to_post_is_stored_in_clipboard()
        {
            // クリップボードにHTMLとTEXT形式のデータが両方存在し、
            // HTML形式データのFragment部分に"teams-copy-link"の文字列があれば、投稿へのリンクが格納されていると判断

            // TEXT形式 もしくは HTML形式のデータがなければ抜ける
            if (!Clipboard.ContainsText(TextDataFormat.Text)) return false;
            if (!Clipboard.ContainsText(TextDataFormat.Html)) return false;

            // HTML形式データのFragment部分を抜き出す(異常時は抜ける)
            string html = Clipboard.GetText(TextDataFormat.Html);
            string fragment_part = extractFragmentPart(html);
            if ( String.IsNullOrEmpty(fragment_part) ) return false;

            // "teams-copy-link"の文字列があれば、投稿へのリンクが格納されていると判断
            return fragment_part.Contains("teams-copy-link");
        }

        static string extractFragmentPart(string html)
        {
            string start_keyword = "<!--StartFragment-->";  // 開始キーワード
            string end_keyword = "<!--EndFragment-->";      // 終了キーワード
            int start_index = html.IndexOf(start_keyword);  // 開始キーワードの検出位置
            int end_index = html.IndexOf(end_keyword);      // 終了キーワードの検出位置

            // 正常にキーワードが見つかった場合、各インデックスは0より大きくなるはず
            if (start_index > 0 && end_index > 0)
            {
                int start = start_index + start_keyword.Length;
                int length = end_index - start;
                return html.Substring(start, length);
            }
            // 異常時は空文字列を返す
            else
            {
                return String.Empty;
            }
        }

        static void convert_link_to_post()
        {
            // Teamsのチャットで「リンクをコピー」を行うと、クリップボードにHtml形式で下記のようなデータが格納される
            // ここから必要なデータ(★で示した部分)を抜き出し、さらに「投稿者名: 」を削除してから、再度Html形式でクリップボードに上書きする
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

            // クリップボードからデータをHTML/TEXT形式で取得
            string html = Clipboard.GetText(TextDataFormat.Html);
            string text = Clipboard.GetText(TextDataFormat.Text);
            Console.WriteLine("--- [before] clipboard data ---");
            Console.WriteLine(html);

            // 「投稿へのリンク部分」を抽出後に投稿者名を削除。異常時は抜ける
            string link = remove_name_of_poster(extract_link_part(html));
            if (string.IsNullOrEmpty(link)) return;

            // 変換結果を表示
            Console.WriteLine("--- [after] link to post ---");
            Console.WriteLine(link);

            // クリップボードに「投稿へのリンク部分」をHtml形式で設定
            // 必須ではないが、テキスト形式で元のクリップボードと同一データも設定しておく
            ClipboardHelper.CopyToClipboard(link, text);
        }

        static string extract_link_part(string html)
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

        static string remove_name_of_poster(string link)
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

        static bool url_text_is_stored_in_clipboard()
        {
            // クリップボードにTEXT形式データがあり
            if (Clipboard.ContainsText(TextDataFormat.Text))
            {
                // そのデータが「http」から始まっていたら、URLと判断する
                if (Clipboard.GetText(TextDataFormat.Text).StartsWith("http")) return true;
            }

            return false;
        }

        static void convert_url_text()
        {
            // Teams上のファイルやフォルダで「リンクをコピー」すると、クリップボードにText形式でSharePointのURLが格納される
            // これをHTMLタグのリンクに変換し、Html形式でクリップボードに上書きする

            // クリップボードからデータをTEXT形式で取得
            string text = Clipboard.GetText(TextDataFormat.Text);
            Console.WriteLine("--- [before] clipboard data ---");
            Console.WriteLine(text);

            // HTMLタグで「URLへのリンク」を作る。文字列は固定
            string link = "<a href=\"" + text + "\"><b><i>HERE!</i></b></a>";
            Console.WriteLine("--- [after] link to url ---");
            Console.WriteLine(link);

            // クリップボードに「URLへのリンク」をHtml形式で設定
            // 必須ではないが、テキスト形式で元のクリップボードと同一データも設定しておく
            ClipboardHelper.CopyToClipboard(link, text);
        }
    }
}
