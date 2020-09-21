using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.parser;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ApExamAnsPdfToExcel
{
    class Program
    {
        static string DirPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(Environment.GetCommandLineArgs()[0]));
        static string ResultFilePath = DirPath + @"\応用情報技術者試験過去問解答.xlsx";

        static void Main(string[] args)
        {
            string pdfFilePath = string.Empty;
            bool roop = true;

            // 出力先ファイルの存在チェック
            if(!File.Exists(ResultFilePath))
            {
                Console.WriteLine("出力先ファイルが存在しません。");
                Console.WriteLine("処理を終了します。");
                Console.ReadLine();
                return;
            }

            // 処理対象PDFファイル名を標準入力
            while (roop)
            {
                Console.WriteLine("---------------------------------------------");
                Console.WriteLine("ファイル名を入力してください。");
                Console.WriteLine("　※処理終了：[n]");
                string pdfFile = Console.ReadLine();
                pdfFilePath = DirPath + @"\" + pdfFile;

                if (String.IsNullOrEmpty(pdfFile))
                    continue;
                if (pdfFile.ToLower() == "n")
                    return;
                if (System.IO.Path.GetExtension(pdfFile).ToLower() != ".pdf")
                {
                    Console.WriteLine("PDFファイルのみ指定可能です。");
                    Console.WriteLine("　※拡張子まで入力して下さい。");
                    continue;
                }
                if (!File.Exists(pdfFilePath))
                {
                    Console.WriteLine("指定されたファイルが存在しません。");
                    continue;
                }

                roop = false;
            }

            // PDFファイルのテキストを取得
            string pdfTxt = GetTextFromAllPages(pdfFilePath);
            // 1行ごとに分割して配列化
            string[] lines = pdfTxt.Replace(@"\r\n", @"\n").Split(new[] { '\n', '\r' });
            // 解答行のみを抽出
            List<ApAmAns> ansList = new List<ApAmAns>();
            foreach (var line in lines)
            {
                string[] words = line.Split(' ');
                if (words.Length < 4)
                    continue;

                if(int.TryParse(words[1], out int i))
                {
                    ApAmAns apAmAns = new ApAmAns(i, words[2], words[3]);
                    ansList.Add(apAmAns);
                }
            }
            
            using (var wb = new XLWorkbook(ResultFilePath))
            {
                // シート存在チェック
                string fileName = System.IO.Path.GetFileNameWithoutExtension(pdfFilePath);
                IXLWorksheet ws = DetectWorkSheet(fileName, wb);

                if (ws == null)
                {
                    Console.WriteLine("出力先ファイルにPDFファイルと同名のシートが存在しません。");
                    Console.WriteLine("処理を終了します。");
                    Console.ReadLine();
                }

                // Excelに書き込み
                int row = 3;
                foreach (var ans in ansList)
                {
                    if(row-2 < ans.No)
                    {
                        // 読み込んだ解答№に歯抜けがある場合、書き込み行もずらす
                        while (row-2 != ans.No)
                            row++;
                    }

                    ws.Cell(row, 1).Value = ans.No;
                    ws.Cell(row, 3).Value = ans.AmAnsTypeToString();
                    ws.Cell(row, 4).Value = ans.QuestionTypeToString();
                }

                // 正答欄のアウトラインを作成
                ws.Columns(3, 5).Group();
                ws.Columns(3, 5).Collapse();

                wb.Save();
            }

            Console.WriteLine("処理正常終了。");
            Console.ReadLine();
        }

        public static string GetTextFromAllPages(String pdfPath)
        {
            PdfReader reader = new PdfReader(pdfPath);

            StringWriter output = new StringWriter();

            for (int i = 1; i <= reader.NumberOfPages; i++)
                output.WriteLine(PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy()));

            return output.ToString();
        }

        private static IXLWorksheet DetectWorkSheet(string fileName, IXLWorkbook wb)
        {
            foreach (var ws in wb.Worksheets)
            {
                if (ws.Name.Equals(fileName))
                    return ws;
            }
            return null;
        }
    }

    public class ApAmAns
    {
        public int No { get; }
        public AmAnsType Ans { get; }
        public QuestionType Type { get; }

        public ApAmAns(int no, string ans, string type)
        {
            No = no;
            Ans = StringToAmAnsType(ans);
            Type = StringToQuestionType(type);
        }

        private AmAnsType StringToAmAnsType(string ans)
        {
            switch (ans)
            {
                case "ア":
                    return AmAnsType.A;
                case "イ":
                    return AmAnsType.I;
                case "ウ":
                    return AmAnsType.U;
                case "エ":
                    return AmAnsType.E;
                default:
                    return AmAnsType.None;
            }
        }

        public string AmAnsTypeToString()
        {
            switch (Ans)
            {
                case AmAnsType.A:
                    return "ア";
                case AmAnsType.I:
                    return "イ";
                case AmAnsType.U:
                    return "ウ";
                case AmAnsType.E:
                    return "エ";
                default:
                    return "[不明]";
            }
        }

        private QuestionType StringToQuestionType(string ans)
        {
            switch (ans)
            {
                case "Ｔ":
                    return QuestionType.T;
                case "Ｍ":
                    return QuestionType.M;
                case "Ｓ":
                    return QuestionType.S;
                default:
                    return QuestionType.None;
            }
        }

        public string QuestionTypeToString()
        {
            switch (Type)
            {
                case QuestionType.T:
                    return "Ｔ";
                case QuestionType.M:
                    return "Ｍ";
                case QuestionType.S:
                    return "Ｓ";
                default:
                    return "[不明]";
            }
        }
    }

    public enum AmAnsType : int
    {
        A, I, U, E, None
    }

    public enum QuestionType : int
    {
        T, M, S, None
    }

}
