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

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;

namespace _1_excel02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            readExcel();

        }
        private static void createExcel()
        {
            //创建工作
            HSSFWorkbook wk = new HSSFWorkbook();
            //创建一个名称为mySheet的表
            ISheet tb = wk.CreateSheet("mySheet");
            //创建第一行
            IRow rowHeader = tb.CreateRow(0);
            //设置行的高度，行高设置数值好像是像素点的1/20，所以*20达到效果;
            rowHeader.Height = 20 * 20;
            ICell cell_one = rowHeader.CreateCell(0);
            cell_one.SetCellValue("这是标题行");
            //设置单元的宽度，宽度数值好像是字符的1/256，所以以便达到设置的要求
            tb.SetColumnWidth(0, 20 * 256);
            //创建第二行
            IRow row = tb.CreateRow(1);
            for (int i = 0; i < 20; i++)
            {
                ICell cell = row.CreateCell(i);   //在第二行中创建单元格
                cell.SetCellValue(i);             //循环往第二行的单元格添加数据 
            }
            string phsicalPath = AppDomain.CurrentDomain.BaseDirectory;
            using (FileStream fs = File.OpenWrite(string.Format(@"D:/高振铭.xls", phsicalPath)))
            {
                wk.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
                Console.WriteLine("提示：创建成功！");
            }
        }
        private static void readExcel()
        {
            StringBuilder sbr = new StringBuilder();
            //生成文件的物理路径
            string phsicalPath = string.Format(@"{0}1.xls", AppDomain.CurrentDomain.BaseDirectory);
            using(FileStream fs = new FileStream(phsicalPath,FileMode.OpenOrCreate)) //打开1.xls文件
            {
                HSSFWorkbook wk = new HSSFWorkbook(fs); //把xls文件写入到数据写于wk中
                //NumberofSheet是Excel中总共的表数量
                for (int i = 0; i < wk.NumberOfSheets; i++)
                {
                    //读取表数据
                    ISheet sheet = wk.GetSheetAt(i);
                    //LastRowNum 是当前表的总行数
                    for (int j = 0; j <= sheet.LastRowNum; j++)
                    {
                        //读取当前行数据
                        IRow row = sheet.GetRow(j);
                        if (row != null)
                        {
                            sbr.Append("-------------------------------------\r\n");//行与行之间的界限
                            //LastCellNum是当前行的总列数
                            for (int k = 0; k <= row.LastCellNum; k++)
                            {
                                //当前表格
                                //当前表格
                                ICell cell = row.GetCell(k);
                                if (cell != null)
                                {
                                    //获取单元格的数据值 并 转换为字符串类型
                                    sbr.Append(cell.ToString());
                                }
                            }
                           
                        }
                    }
                }
            }
            string phsicalPath_txt = string.Format(@"{0}myText.txt", AppDomain.CurrentDomain.BaseDirectory);
            //把读取xls文件的数据写入myText.txt文件中
            using (StreamWriter wr = new StreamWriter(new FileStream(phsicalPath_txt, FileMode.Append)))
            {
                wr.Write(sbr.ToString());
                wr.Flush();
            }
            Console.WriteLine("读取成功");
            
        }
    }
}
