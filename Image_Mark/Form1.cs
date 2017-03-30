using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Image_Mark
{
    public partial class Form1 : Form
    {
        String path;
        Image img;

        private static int maxn = 500;
        public String[] list = new String[maxn];
        public int currentPosition;
        public int imgCount;

        Microsoft.Office.Interop.Excel.Application app = null;
        Microsoft.Office.Interop.Excel.Workbook wb = null;
        Microsoft.Office.Interop.Excel.Worksheet ws = null;

        public void init()
        {
            try
            {
                list = System.IO.File.ReadAllLines("list.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not found list.txt on Current directory");

                Application.Exit();
                Environment.Exit(0);
            }

            try
            {
                String[] reader = System.IO.File.ReadAllLines("confugre.ini");

                currentPosition = Int32.Parse(reader[0]);
                imgCount = Int32.Parse(reader[1]);

                MessageBox.Show("当前处理到第 " + (currentPosition + 1) + " 张, 共需处理 " + imgCount + " 张");
            }
            catch (Exception ex)
            {
                currentPosition = 0;
                imgCount = list.Length;

                System.IO.File.WriteAllText("confugre.ini", "" + currentPosition + "\n" + imgCount);
            }

        }

        public void openExl()
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;

            wb = app.Workbooks.Add(Type.Missing);

            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];

            ws.Activate();

            ws.Name = "Result";

            ws.Cells[1, 1] = "ImagePath";
            ws.Cells[1, 2] = "一处稀线或开口";
            ws.Cells[1, 3] = "多跟稀线";
            ws.Cells[1, 4] = "帘线交叉或重叠";
            ws.Cells[1, 5] = "胎体帘线弯曲";
            ws.Cells[1, 6] = "胎体帘线折断";
            ws.Cells[1, 7] = "散线";
            ws.Cells[1, 8] = "杂物";
            ws.Cells[1, 9] = "气泡";
            ws.Cells[1, 10] = "带束层中心和胎冠中心线偏差<300mm";
            ws.Cells[1, 11] = "带束层中心和胎冠中心线偏差>300mm";
            ws.Cells[1, 12] = "0'带束层外端和第二带束层端点偏离左右差长度<=300mm";
            ws.Cells[1, 13] = "0'带束层外端和第二带束层端点偏离左右差长度>300mm";
            ws.Cells[1, 14] = "0'带束层和第三带束层重叠";
            ws.Cells[1, 15] = "0’带束层和第三带束层间隙<=300mm";
            ws.Cells[1, 16] = "0’带束层和第三带束层间隙>300mm";
            ws.Cells[1, 17] = "接头重叠、钢丝重叠";
            ws.Cells[1, 18] = "接头开、变稀、缺线和散头";
            ws.Cells[1, 19] = "0‘散线";
            ws.Cells[1, 20] = "带束层缺失";
            ws.Cells[1, 21] = "带束层方向";
            ws.Cells[1, 22] = "杂物";
            ws.Cells[1, 23] = "带束层内气泡";
            ws.Cells[1, 24] = "胎体反包层和钢丝包布差级，与标准值的差";
            ws.Cells[1, 25] = "变稀.接头开.缺线和散头";
            ws.Cells[1, 26] = "散线、 帘线、 长度（包括钢丝包布)";
            ws.Cells[1, 27] = "钢丝帘线连续重叠";
            ws.Cells[1, 28] = "杂物";
            ws.Cells[1, 29] = "气泡";
            ws.Cells[1, 30] = "胎体反包层左右差";
            ws.Cells[1, 31] = "钢丝包布交叉";
            ws.Cells[1, 32] = "无内胎胎体反包高度偏低";
            ws.Cells[1, 33] = "钢包缺失";
            ws.Cells[1, 34] = "胎圈变形";
            ws.Cells[1, 35] = "胎圈张嘴";
            ws.Cells[1, 36] = "胎圈折断";

            wb.SaveAs("Result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }



        public Form1()
        {
            init();
            openExl();

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //保存exl
            ws.Cells[currentPosition + 2, 1] = path;

            if(checkBox1.Checked) ws.Cells[currentPosition + 2, 2] = "1";
                else ws.Cells[currentPosition + 2, 2] = "0";

            if (checkBox2.Checked) ws.Cells[currentPosition + 2, 3] = "1";
                else ws.Cells[currentPosition + 2, 3] = "0";

            if (checkBox3.Checked) ws.Cells[currentPosition + 2, 4] = "1";
                else ws.Cells[currentPosition + 2, 4] = "0";

            if (checkBox4.Checked) ws.Cells[currentPosition + 2, 5] = "1";
                else ws.Cells[currentPosition + 2, 5] = "0";

            if (checkBox5.Checked) ws.Cells[currentPosition + 2, 6] = "1";
                else ws.Cells[currentPosition + 2, 6] = "0";

            if (checkBox6.Checked) ws.Cells[currentPosition + 2, 7] = "1";
                else ws.Cells[currentPosition + 2, 7] = "0";

            if (checkBox7.Checked) ws.Cells[currentPosition + 2, 8] = "1";
                else ws.Cells[currentPosition + 2, 8] = "0";

            if (checkBox8.Checked) ws.Cells[currentPosition + 2, 9] = "1";
                else ws.Cells[currentPosition + 2, 9] = "0";

            if (checkBox9.Checked) ws.Cells[currentPosition + 2, 10] = "1";
                else ws.Cells[currentPosition + 2, 10] = "0";

            if (checkBox10.Checked) ws.Cells[currentPosition + 2, 11] = "1";
                else ws.Cells[currentPosition + 2, 11] = "0";

            if (checkBox11.Checked) ws.Cells[currentPosition + 2, 12] = "1";
                else ws.Cells[currentPosition + 2, 12] = "0";

            if (checkBox12.Checked) ws.Cells[currentPosition + 2, 13] = "1";
                else ws.Cells[currentPosition + 2, 13] = "0";

            if (checkBox13.Checked) ws.Cells[currentPosition + 2, 14] = "1";
                else ws.Cells[currentPosition + 2, 14] = "0";

            if (checkBox14.Checked) ws.Cells[currentPosition + 2, 15] = "1";
                else ws.Cells[currentPosition + 2, 15] = "0";

            if (checkBox15.Checked) ws.Cells[currentPosition + 2, 16] = "1";
                else ws.Cells[currentPosition + 2, 16] = "0";

            if (checkBox16.Checked) ws.Cells[currentPosition + 2, 17] = "1";
                else ws.Cells[currentPosition + 2, 17] = "0";

            if (checkBox17.Checked) ws.Cells[currentPosition + 2, 18] = "1";
                else ws.Cells[currentPosition + 2, 18] = "0";

            if (checkBox18.Checked) ws.Cells[currentPosition + 2, 19] = "1";
                else ws.Cells[currentPosition + 2, 19] = "0";

            if (checkBox19.Checked) ws.Cells[currentPosition + 2, 20] = "1";
                else ws.Cells[currentPosition + 2, 20] = "0";

            if (checkBox20.Checked) ws.Cells[currentPosition + 2, 21] = "1";
                else ws.Cells[currentPosition + 2, 21] = "0";

            if (checkBox21.Checked) ws.Cells[currentPosition + 2, 22] = "1";
                else ws.Cells[currentPosition + 2, 22] = "0";

            if (checkBox22.Checked) ws.Cells[currentPosition + 2, 23] = "1";
                else ws.Cells[currentPosition + 2, 23] = "0";

            if (checkBox23.Checked) ws.Cells[currentPosition + 2, 24] = "1";
                else ws.Cells[currentPosition + 2, 24] = "0";

            if (checkBox24.Checked) ws.Cells[currentPosition + 2, 25] = "1";
                else ws.Cells[currentPosition + 2, 25] = "0";

            if (checkBox25.Checked) ws.Cells[currentPosition + 2, 26] = "1";
                else ws.Cells[currentPosition + 2, 26] = "0";

            if (checkBox26.Checked) ws.Cells[currentPosition + 2, 27] = "1";
                else ws.Cells[currentPosition + 2, 27] = "0";

            if (checkBox27.Checked) ws.Cells[currentPosition + 2, 28] = "1";
                else ws.Cells[currentPosition + 2, 28] = "0";

            if (checkBox28.Checked) ws.Cells[currentPosition + 2, 29] = "1";
                else ws.Cells[currentPosition + 2, 29] = "0";

            if (checkBox29.Checked) ws.Cells[currentPosition + 2, 30] = "1";
                else ws.Cells[currentPosition + 2, 30] = "0";

            if (checkBox30.Checked) ws.Cells[currentPosition + 2, 31] = "1";
                else ws.Cells[currentPosition + 2, 31] = "0";

            if (checkBox31.Checked) ws.Cells[currentPosition + 2, 32] = "1";
                else ws.Cells[currentPosition + 2, 32] = "0";

            if (checkBox32.Checked) ws.Cells[currentPosition + 2, 33] = "1";
                else ws.Cells[currentPosition + 2, 33] = "0";

            if (checkBox33.Checked) ws.Cells[currentPosition + 2, 34] = "1";
                else ws.Cells[currentPosition + 2, 34] = "0";

            if (checkBox34.Checked) ws.Cells[currentPosition + 2, 35] = "1";
                else ws.Cells[currentPosition + 2, 35] = "0";
            
            if (checkBox35.Checked) ws.Cells[currentPosition + 2, 36] = "1";
                else ws.Cells[currentPosition + 2, 36] = "0";


            wb.SaveAs("Result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //保存confugre.ini
            System.IO.File.WriteAllText("confugre.ini", "" + currentPosition + "\n" + imgCount);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            path = list[currentPosition];


            try
            {
                img = Image.FromFile(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Not find the file!");
                return;
            }
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.Image = img;
            textBox1.Text = path;
            label5.Text = "" + (currentPosition + 1) + "/" + imgCount;
        }
        
        ~Form1()
        {
            wb.SaveAs("Result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close(true, Type.Missing, Type.Missing);

            wb = null;
            app.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ++currentPosition;
            if(currentPosition >= imgCount)
            {
                MessageBox.Show("This is the last Image");
                return;
            }

            path = list[currentPosition];
            img = Image.FromFile(path);
            pictureBox1.Image = img;
            textBox1.Text = path;
            label5.Text = "" + (currentPosition + 1) + "/" + imgCount;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            wb.SaveAs("Result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close(true, Type.Missing, Type.Missing);

            wb = null;
            app.Quit();
            base.OnClosing(e);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            --currentPosition;

            if(currentPosition < 0)
            {
                MessageBox.Show("This is the first Image");
                return;
            }

            path = list[currentPosition];
            img = Image.FromFile(path);
            pictureBox1.Image = img;
            textBox1.Text = path;
            label5.Text = "" + (currentPosition + 1) + "/" + imgCount;
        }
    }
}
