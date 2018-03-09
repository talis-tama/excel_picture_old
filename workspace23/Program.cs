using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace workspace23
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
    public class Form1 : Form 
    {
        TextBox textbox1, textbox2;
        public Form1()
        {
            Width = 400;
            Height = 900;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Label label1 = new Label();
            label1.Location = new Point(10, 10);
            label1.Size = new Size(300, 15);
            label1.Text = "読み込む画像ファイルを指定してください";
            Controls.Add(label1);
            textbox1 = new TextBox();
            textbox1.Location = new Point(10, 35);
            textbox1.Size = new Size(300, 15);
            Controls.Add(textbox1);
            Button button = new Button();
            button.Location = new Point(10, 60);
            button.Size = new Size(300, 30);
            button.Text = "開始";
            button.Click += new EventHandler(button_click);
            Controls.Add(button);
            textbox2 = new TextBox();
            textbox2.Location = new Point(10, 100);
            textbox2.Size = new Size(300, 750);
            textbox2.Multiline = true;
            textbox2.ScrollBars = ScrollBars.Vertical;
            textbox2.ReadOnly = true;
            Controls.Add(textbox2);
        }
        void button_click(object sender, EventArgs e)
        {
            string imagename;
                imagename = textbox1.Text;
                int wid, hig/*, r, g, bb*/;
            string r, g, bb;
            Bitmap img = new Bitmap(imagename);
        //StreamWriter output = new StreamWriter(@"D:\test.txt", false, Encoding.GetEncoding("shift_jis"));*/
        using (var book = new XLWorkbook(@"C:\Users\" + Environment.UserName + "\\Desktop\\test.xlsx", XLEventTracking.Disabled))
        {
            var sheet1 = book.Worksheet(1);
            sheet1.Columns().Width = 0.08;
            sheet1.Rows().Height = 4.50;
            wid = img.Width;
            hig = img.Height;
            book.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
            book.Dispose();
        }/*
                using (var bookaft = new XLWorkbook(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx", XLEventTracking.Disabled))
                {
            var sheet1 = bookaft.Worksheet(1);
            for (int a = 0; a < hig; a++)
            {
                    for (int b = 0; b < wid; b++)
                    {
                        Color col = img.GetPixel(b, a);
                        r = Convert.ToString(col.R, 16);
                        g = Convert.ToString(col.G, 16);
                        bb = Convert.ToString(col.B, 16);
                        if (r.Length == 1) { r = "0" + r; }
                        if (g.Length == 1) { g = "0" + g; }
                        if (bb.Length == 1) { bb = "0" + bb; }
                        var cell = sheet1.Cell(a + 1, b + 1);
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#" + r + g + bb);
                        textbox2.AppendText(b.ToString() + "," + a.ToString() + Environment.NewLine);
                    }
                    bookaft.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testafter.xlsx");
                sheet1.Dispose();
                bookaft.Dispose();
            }
        }*/
            //output.Close();
            EXCEL.Application task;
            EXCEL.Workbook bookafter;
            EXCEL.Worksheet sheet;
            EXCEL.Range cell;
            task = new EXCEL.Application();
            bookafter = (EXCEL.Workbook)(task.Workbooks.Open(@"C:\Users\"+ Environment.UserName + "\\Desktop\\testafter.xlsx"));
            task.Application.Visible = true;
            sheet = (EXCEL.Worksheet)bookafter.Sheets[1];
            sheet.Select(Type.Missing);
            for (int a = 0; a < hig; a++)
            {
                for (int b = 0; b < wid; b++)
                {
                    Color col = img.GetPixel(b, a);
                    r = Convert.ToString(col.R, 16);
                    g = Convert.ToString(col.G, 16);
                    bb = Convert.ToString(col.B, 16);
                    if (r.Length == 1) { r = "0" + r; }
                    if (g.Length == 1) { g = "0" + g; }
                    if (bb.Length == 1) { bb = "0" + bb; }
                    cell = (EXCEL.Range)sheet.Cells[a + 1, b + 1];
                    cell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#"+r+g+bb));
                    textbox2.AppendText(b.ToString() + "," + a.ToString() + Environment.NewLine);
                }
            }
            bookafter.SaveAs(@"C:\Users\" + Environment.UserName + "\\Desktop\\testoutput.xlsx");
            bookafter.Close();
            task.Quit();
        }
    }
}
