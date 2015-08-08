using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

namespace NumCsvOpener
{
    static class Program
    {
        static Dictionary<string, Encoding> encodings =
            new Dictionary<string, Encoding>()
            {
                { "BIG5", Encoding.GetEncoding(950) },
                { "UTF8", Encoding.UTF8 },
                { "簡體中文", Encoding.GetEncoding(936) },
                { "日文", Encoding.GetEncoding(932) }
            };
        static Encoding encoding = encodings["BIG5"];

        //彈出對話框選擇編碼
        static void PickupEncoding()
        {
            using (Form prompt = new Form())
            {
                prompt.Width = 144;
                prompt.Height = 80 + 24 * encodings.Count;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.MaximizeBox = prompt.MinimizeBox = false;
                prompt.Text = "請選擇檔案編碼";
                prompt.StartPosition = FormStartPosition.CenterScreen;
                int top = 12;
                foreach (string key in encodings.Keys)
                {
                    Button button = new Button() {
                        Text = key,
                        Left = 12, Top = top,
                        Width = 100, Height = 24,
                        DialogResult = DialogResult.OK
                    };
                    top += 30;
                    button.Click += (sender, e) => {
                        //選取指定的編碼
                        encoding = encodings[key];
                        prompt.Close();
                    };
                    if (key == "BIG5")
                        prompt.AcceptButton = button;
                    prompt.Controls.Add(button);
                }
                prompt.ShowDialog();
            }
        }

        [STAThread]
        static void Main()
        {
            try
            {
                //由使用者挑選CSV檔
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "*.csv|*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string fn = ofd.FileName;
                    //開放使用者選取編碼
                    PickupEncoding();

                    using (StreamReader sr = new StreamReader(fn, encoding, true))
                    {
                        StringBuilder sb = new StringBuilder();
                        bool quotMarkMode = false;
                        string newLineReplacement = "\x07";
                        //支援CSV雙引號內含換行符號規則，採逐字讀入解析
                        //雙引號內如需表示", 使用""代替
                        while (sr.Peek() >= 0)
                        {
                            var ch = (char)sr.Read();
                            if (quotMarkMode)
                            {
                                //雙引號包含區段內遇到雙引號有兩種情境
                                if (ch == '"')
                                {
                                    //連續兩個雙引號，為欄位內雙引號字元
                                    if (sr.Peek() == '"')
                                    {
                                        sb.Append(sr.Read());
                                    }
                                    //遇到結尾雙引號，雙引號包夾模式結束
                                    else
                                    {
                                        quotMarkMode = false;
                                    }
                                    sb.Append(ch);
                                }
                                //雙引號內遇到換行符號，先置換成特殊號，稍後換回
                                else if (ch == '\r' && sr.Peek() == '\n')
                                {
                                    sr.Read();
                                    sb.Append(newLineReplacement);
                                }
                                else
                                {
                                    sb.Append(ch);
                                }
                            }
                            else
                            {
                                sb.Append(ch);
                                if (ch == '"') quotMarkMode = true;
                            }
                        }
                        var fixedCsv = sb.ToString();
                        sb.Length = 0;
                        string line;
                        using (var lr = new StringReader(fixedCsv))
                        {
                            while ((line = lr.ReadLine()) != null)
                            {
                                string[] p = line.Split(',');
                                sb.AppendLine(string.Join(",",
                                    //若欄位以0起首，重新組裝成="...."格式
                                    p.Select(o => 
                                        o.StartsWith("0") ? 
                                            string.Format("=\"{0}\"", o) : 
                                            //還原換行符號
                                            o.StartsWith("\"") ? 
                                                o.Replace(newLineReplacement, "\r\n") : 
                                                o
                                    ).ToArray()));
                            }
                        }

                        //調整結果另存為同目錄下*.fixed.csv檔
                        string fixedFile = Path.Combine(
                            Path.GetDirectoryName(fn), 
                            Path.GetFileNameWithoutExtension(fn) + ".fixed.csv");
                        //一律存為UTF8
                        File.WriteAllText(fixedFile, sb.ToString(), Encoding.UTF8);
                        //開啟CSV
                        Process proc = new Process();
                        proc.StartInfo = new ProcessStartInfo(fixedFile);
                        proc.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
