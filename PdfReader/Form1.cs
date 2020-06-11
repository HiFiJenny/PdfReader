using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using org.pdfbox.pdmodel;
using org.pdfbox.util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Runtime.InteropServices;
using NPOI.SS.UserModel;

namespace PdfReader
{
    public partial class Form1 : Form
    {
        NPOI.SS.UserModel.ISheet st;
        pdfexcelpath address = new pdfexcelpath();

        /// <summary>
        /// 窗体动画函数
        /// </summary>
        /// <param name="hwnd">指定产生动画的窗口的句柄</param>
        /// <param name="dwTime">指定动画持续的时间</param>
        /// <param name="dwFlags">指定动画类型，可以是一个或多个标志的组合。</param>
        /// <returns></returns>
        [DllImport("user32")]
        private static extern bool AnimateWindow(IntPtr hwnd, int dwTime, int dwFlags);

        //下面是可用的常量，根据不同的动画效果声明自己需要的
        private const int AW_HOR_POSITIVE = 0x0001;//自左向右显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_HOR_NEGATIVE = 0x0002;//自右向左显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_VER_POSITIVE = 0x0004;//自顶向下显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_VER_NEGATIVE = 0x0008;//自下向上显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志该标志
        private const int AW_CENTER = 0x0010;//若使用了AW_HIDE标志，则使窗口向内重叠；否则向外扩展
        private const int AW_HIDE = 0x10000;//隐藏窗口
        private const int AW_ACTIVE = 0x20000;//激活窗口，在使用了AW_HIDE标志后不要使用这个标志
        private const int AW_SLIDE = 0x40000;//使用滑动类型动画效果，默认为滚动动画类型，当使用AW_CENTER标志时，这个标志就被忽略
        private const int AW_BLEND = 0x80000;//使用淡入淡出效果
        //窗体代码（将窗体的FormBorderStyle属性设置为none）：

        public Form1()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AnimateWindow(this.Handle, 200, AW_BLEND | AW_ACTIVE);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog.Filter = "文本文件|*.*|C#文件|*.cs|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                address.pdf_path = openFileDialog.FileName;
                if (address.pdf_path.Contains("pdf"))
                {
                    PDFpath.Text = address.pdf_path;
                }
                else
                {
                    MessageBox.Show("文件选择错误");
                    PDFpath.Text = " ";
                    address.pdf_path = " ";
                }

            }
        }

        public class pdfexcelpath
        {
            public string pdf_path;
            public string excel_path;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            SaveFileDialog dialogs = new SaveFileDialog();
            dialogs.Filter = "Excel|*.xlsx";
            // dialogs.DefaultExt = "*.xlsx";
            dialogs.FileName = "CalCer.xlsx";
            if (dialogs.ShowDialog() == DialogResult.OK)
            {
                address.excel_path = dialogs.FileName;
                EXCELpath.Text = address.excel_path;
            }
        }

            private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            FileStream fs;
            XSSFWorkbook wk;
            PDDocument doc = PDDocument.load(address.pdf_path);
            PDFTextStripper pdfstripper = new PDFTextStripper();
            string str = pdfstripper.getText(doc);
            string[] str1 = str.Split(new string[] { "Sensor" }, StringSplitOptions.RemoveEmptyEntries);       //提取sensor之后的内容：数据
            string[] str2 = str1[1].Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);  //提取每一行的数据
            string[] standard = str2[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries); //提取第一行标准值   
            //   swPdfChange.Write(str1[1]);  //str1 中保存所需要的内容： 标准测量点 校准值
            int num = str2.Length - 1;   //获取的值的行数
                                         //    swPdfChange.Close();
            string TempletFileName =  @".\excel.xlsx";  //模板文件  

            using ( fs = File.OpenRead(TempletFileName))
            {
                FileStream fs2 = File.Create(address.excel_path);
                wk = new XSSFWorkbook(fs);

                ICellStyle tableStyle = wk.CreateCellStyle();
                st = wk.GetSheet("52-301,52-301.1");
                IRow r = st.CreateRow(0);
                ICell c = null;
                int cnt = 0;
                for (int i = 0; i < num; i++)
                {
                    string[] value = str2[i + 1].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if ((i % 4 ==0) && (i >= 4 ))
                    {
                        cnt += 24;
                        goto A;
                    }
                A:

                    for (int j = 0; j < standard.Length; j++)        //标准值/测量点写入excel
                    {
                            char[] TrimChar = { '(', ')' };
                            int[] col = { 1, 2, 14 };
                            string[] val = { value[0], standard[j], value[(j + 1) * 2].Trim(TrimChar) };

                            r = st.CreateRow(i * 5 + 68 + j + cnt);
                            for (int dic = 0; dic < 3; dic++)
                            {
                                c = r.CreateCell(col[dic]);
                                c.SetCellValue(val[dic]);
                        }

                            for (int k = 0; k < 8; k++)        //4次读数
                            {
                                c = r.CreateCell(k + 5);
                                c.SetCellValue(value[(j + 1) * 2 - 1]);
                            }

                    }
                    c.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    c.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    c.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                }

                wk.Write(fs2);
                fs2.Close();
                fs.Close();
            }
            this.Cursor = Cursors.Arrow;
            MessageBox.Show("转换完成", "PDF->Excel", MessageBoxButtons.OK);
        }


        /// <summary>
        /// 顶部控件： 关闭窗口 拖动窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private Point mousePoint = new Point();
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            base.OnMouseDown(e);
            this.mousePoint.X = e.X;
            this.mousePoint.Y = e.Y;
        }
        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.Button == MouseButtons.Left)
            {
                this.Top = Control.MousePosition.Y - mousePoint.Y;
                this.Left = Control.MousePosition.X - mousePoint.X;
            }
        }



        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            AnimateWindow(this.Handle, 200, AW_BLEND | AW_HIDE);
        }

        private void PDFpath_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
