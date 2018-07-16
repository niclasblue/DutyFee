using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using static System.Console;

namespace DutyFee
{
    public static class DutyFeeTableBuilder
    {
        public static void Builde()
        {

            // 从文件中读取名字
            // 第一行为标题行
            //string path = Environment.CurrentDirectory;
            FileStream stream = new FileStream("names.txt", FileMode.Open);
            StreamReader reader = new StreamReader(stream);
            List<string> names = new List<string>();
            while (reader.Peek() >= 0)
            {
                names.Add(reader.ReadLine());
            }
            reader.Close();

            string title = names[0];
            names.RemoveAt(0);

            //创建xlsx文档，使用Spire.XLS
            Workbook workbook = new Workbook();
            workbook.CreateEmptySheet();
            Worksheet worksheet = workbook.Worksheets[0];

            //首先确定要生成数据表的行数和列数

            WriteLine("Please input the DateTime you want create?");
            //初始的时间是程序运行时间的上一个月。
            int year = DateTime.Now.Year;
            int month = DateTime.Now.AddMonths(-1).Month;

            int tempYear = year;
            int tempMonth = month;
            do
            {
                Write($"Please input the year:(default:{year})");
                try
                {
                    tempYear = Convert.ToInt32(ReadLine());
                }
                catch (Exception)
                {
                    tempYear = year;
                }
                Write($"Please input the month:(default:{month})");
                try
                {
                    tempMonth = Convert.ToInt32(ReadLine());
                }
                catch (Exception)
                {
                    tempMonth = month;
                }

            } while ((tempYear <= 0) || (tempMonth <= 0) || (tempMonth > 12));

            year = tempYear;
            month = tempMonth;
            //获取当月的天数
            int days = DateTime.DaysInMonth(year, month);

            //将要填充的数据先放到一个Object类型的二维数组当中
            int row = names.Count + 2;
            int col = days + 1;
            object[,] myData = new object[row, col];

            //填充值班数据
            for (int j = 1; j < row; j++)
            {
                for (int i = 1; i < col; i++)
                {
                    myData[j, i] = "";
                }
            }
            myData[0, 0] = "姓名";
            myData[row - 1, 0] = "备注";
            myData[row - 1, 1] = "加班☆；值班△；休假□；请事（病）假★。";


            //从json文件中读取节假日数据
            string configPath = string.Format("holiday.json");
            HolidayChecker checker = new HolidayChecker(configPath);
            for (int i = 1; i <= days; i++)
            {
                myData[0, i] = $"{i}日";
                worksheet.SetColumnWidth(i + 1, 3.5);
                //修改颜色
                DateTime currentDay = new DateTime(year, month, i);
                DateType currentDayType = checker.CheckDayType(currentDay);
                if ((currentDayType is DateType.Weekend) || (currentDayType is DateType.Holiday))
                {
                    string range = String.Format("{0}1:{0}{1}", NumberCharsConvert.NumberToChars(i + 1), names.Count-1);
                    worksheet.Range[range].Style.Color = (currentDayType is DateType.Weekend ? Color.Gray : Color.Red);
                }
            }
            for (int j = 1; j <= names.Count; j++)
            {
                myData[j, 0] = names[j - 1];

                worksheet.SetRowHeight(j + 1, 40);
            }


            worksheet.InsertArray(myData, 1, 1);
            //整理表格格式
            //设置列宽和行高
            worksheet.SetRowHeight(1, 40);

            // worksheet.SetRowHeight(names.Count + 2, 40);
           // worksheet.Rows[row-1].RowHeight = 40;
            //worksheet.Rows[row-1].Style.Color = Color.;

            worksheet.SetColumnWidth(1, 6);
            //设置边框
            CellRange xlsRange = worksheet.Range[$"A1:{NumberCharsConvert.NumberToChars(col)}{row}"];
            xlsRange.BorderInside(LineStyleType.Thin, Color.Black);
            xlsRange.BorderAround(LineStyleType.Thin, Color.Black);
            //设置对齐方式
            xlsRange.Style.HorizontalAlignment = HorizontalAlignType.Center;
            xlsRange.Style.VerticalAlignment = VerticalAlignType.Center;
            //合并单元格
            worksheet.Range[$"B{names.Count + 2}:{NumberCharsConvert.NumberToChars(days + 1)}{names.Count + 2}"].Merge();
            worksheet.Range[$"B{names.Count + 2}"].HorizontalAlignment = HorizontalAlignType.Left;
            //设置打印页面格式为A4
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            //set header&footer
            worksheet.PageSetup.CenterHeader = $"{title}{year}年{month}月考勤登记表";
            worksheet.PageSetup.LeftFooter = "制表人：                                 政工复核：                           审核人：";


            //保存文件
            workbook.SaveToFile(@"..\..\" + $"{year}年{month}月值（加）班考勤表.xlsx", ExcelVersion.Version2016);
        }
    }
}
