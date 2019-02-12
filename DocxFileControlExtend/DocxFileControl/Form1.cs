using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        // 参照にて
        // Microsoft Word 14.0 Object Library
        // 上記を追加してください。
        // そうすると、参照設定に
        // Microsoft.Office.Core
        // Microsoft.Office.Interop.Word

        public Form1()
        {
            InitializeComponent();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            try
            {
                // Word アプリケーションオブジェクトを作成
                Word.Application word = new Word.Application();
                // Word の GUI を起動しないようにする
                word.Visible = false;

                // 新規文書を作成
                Document document = word.Documents.Add();

                // ヘッダーを編集
                editHeaderSample(ref document, 10, WdColorIndex.wdPink, "Header Area");

                // フッターを編集
                editFooterSample(ref document, 10, WdColorIndex.wdBlue, "Footer Area");

                // 見出しを追加
                addHeadingSample(ref document, "見出し");

                // パラグラフを追加
                document.Content.Paragraphs.Add();

                // テキストを追加
                addTextSample(ref document, WdColorIndex.wdGreen, "Hello, ");
                addTextSample(ref document, WdColorIndex.wdRed, "World");

                // 名前を付けて保存
                object filename = System.IO.Directory.GetCurrentDirectory() + @"\out.docx";
                document.SaveAs2(ref filename);

                // 文書を閉じる
                document.Close();
                document = null;
                word.Quit();
                word = null;

                // 情報通知
                MessageBox.Show("【成功】Docxファイル作成成功");
            }
            catch (Exception ex)
            {
                // 情報通知
                MessageBox.Show("【失敗】Docxファイル作成 : " + ex.Message);
            }
        }

        /// <summary>
        /// ■リプレース
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReplace_Click(object sender, EventArgs e)
        {
            Word.Application word = new Word.Application();
            Document documentList = new Document();

            // Define an object to pass to the word API for missing parameters
            object missing = Type.Missing;

            try
            {
                // ■【環境】ファイル制御ベース
                object fileName = @"D:\workspace\Sample\\DocxFileControlExtend\DocxFileControl\Template.docx";
                Document baseDoc = word.Documents.Add(fileName, ref missing, ref missing, ref missing);

                string startTime = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                string endTime2;
                string endTime3;

                for (int i = 1; i <= 3; i++)
                {
                    // ①．ドキュメント追加
                    documentList = word.Documents.Add(baseDoc, ref missing, ref missing, ref missing);
                    
                    if (i == 1) endTime2 = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);

                    // ②．本文データ置換処理
                    for (int j = 1; j <= 20; j++)
                    {
                        string txtReplace1_ = txtReplace1.Text;
                        this.FindAndReplace2(word, "@REPLACE" + j.ToString(), txtReplace1_ + j.ToString());
                    }
                    if (i == 1) endTime3 = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                }

                string endTime = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);

                // ■【整理】最終的ファイルコントロール
                object newFileName = @"D:\workspace\Sample\\DocxFileControlExtend\DocxFileControl\Template_Replace.docx";

                object copies = "1";
                object pages = "";
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object oTrue = true;
                object oFalse = false;

                foreach (var tar in word.Documents)
                {
                    (tar as Document).PrintOut(ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
                      ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                      ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
                }

                // 情報通知
                MessageBox.Show("【成功】Docxファイル作成");
            }
            catch (Exception ex)
            {
                // 情報通知
                MessageBox.Show("【失敗】Docxファイル作成 : " + ex.Message);
            }
            finally
            {
                if (word != null) word = null;
                if (documentList != null) documentList = null;
            }
        }


        private void FindAndReplace2(Word.Application WordApp,
                                    object findText,
                                    object replaceWithText)
        {
            object missing = Type.Missing;
            bool result =
                WordApp.Application.Selection.Find.Execute(ref findText,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing);
            if (result) { WordApp.Application.Selection.Text = (string)replaceWithText; }
        }

        /// <summary>
        /// 文言を探して該当文字を置換する
        /// </summary>
        /// <param name="WordApp"></param>
        /// <param name="findText"></param>
        /// <param name="replaceWithText"></param>
        private void FindAndReplace(Word.Application WordApp,
                                    object findText,
                                    object replaceWithText)
        {
            object missing = Type.Missing;

            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            if (replaceWithText.ToString().Length < 250) // Normal execution
            {
                WordApp.Selection.Find.Execute(ref findText,
                                               ref matchCase,
                                               ref matchWholeWord,
                                               ref matchWildCards,
                                               ref matchSoundLike,
                                               ref nmatchAllWordForms,
                                               ref forward,
                                               ref wrap,
                                               ref format,
                                               ref replaceWithText,
                                               ref replace,
                                               ref matchKashida,
                                               ref matchDiacritics,
                                               ref matchAlefHamza,
                                               ref matchControl);
            }
            else
            {
                WordApp.Application.Selection.Find.Execute(ref findText,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing,
                                                           ref missing);

                WordApp.Application.Selection.Text = (string)replaceWithText;
            }
        }

        /// <summary>
        /// 文書のヘッダーを編集する.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="fontSize"></param>
        /// <param name="color"></param>
        /// <param name="text"></param>
        private static void editHeaderSample(ref Document document, int fontSize, WdColorIndex color, string text)
        {
            foreach (Section section in document.Sections)
            {
                //Get the header range and add the header details.
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = color;
                headerRange.Font.Size = fontSize;
                headerRange.Text = text;
            }
        }

        /// <summary>
        /// 文書のフッターを編集する.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="fontSize"></param>
        /// <param name="color"></param>
        /// <param name="text"></param>
        private static void editFooterSample(ref Document document, int fontSize, WdColorIndex color, string text)
        {
            foreach (Section wordSection in document.Sections)
            {
                //Get the footer range and add the footer details.
                Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = color;
                footerRange.Font.Size = fontSize;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = text;
            }
        }

        /// <summary>
        /// 文書に見出しを追加する.
        /// </summary>
        private static void addHeadingSample(ref Document document, string text)
        {
            Paragraph para = document.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            object styleHeading1 = "見出し 1";
            para.Range.set_Style(ref styleHeading1);
            para.Range.Text = text;
            para.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// 文書の末尾位置を取得する.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private static int getLastPosition(ref Document document)
        {
            return document.Content.End - 1;
        }

        /// <summary>
        /// 文書の末尾にテキストを追加する.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="color"></param>
        /// <param name="text"></param>
        private static void addTextSample(ref Document document, WdColorIndex color, string text)
        {
            int before = getLastPosition(ref document);
            Range rng = document.Range(document.Content.End - 1, document.Content.End - 1);
            rng.Text += text;
            int after = getLastPosition(ref document);

            document.Range(before, after).Font.ColorIndex = color;
        }
    }
}
