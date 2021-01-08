using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;

namespace GridViewExport
{
    public class ExportExcel
    {
        static int colIDWidth = 200;
        static int colNameWidth = 200;
        static int colSpecWidth = 200;
        static int colItemWidth = 50;

        public static void GridToExcel(string fileName, List<GridViewExport.Form1.ExportItem> list)
        {
            if (list==null|| list.Count == 0)
            {
                return;
            }
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel 2003格式|*.xls";
            sfd.FileName = fileName + DateTime.Now.ToString("yyyyMMddHHmmssms");
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet(fileName);
            HSSFRow headRow = (HSSFRow)sheet.CreateRow(0);
            HSSFRow headRow2 = (HSSFRow)sheet.CreateRow(1); //航头第二行

            sheet.SetColumnWidth(0, colIDWidth * 30);
            sheet.SetColumnWidth(1, colNameWidth * 30);
            sheet.SetColumnWidth(2, colSpecWidth * 30);

            ICellStyle cellStyle = Getcellstyle(wb);
            //合并航头
            for (int i = 0; i < 15; i++)
            {
                HSSFCell headCell = (HSSFCell)headRow.CreateCell(i, CellType.String);
                headCell.CellStyle = cellStyle;
                headCell = (HSSFCell)headRow2.CreateCell(i, CellType.String);
                headCell.CellStyle = cellStyle;

                sheet.SetColumnWidth(3 + i, colItemWidth * 30);
            }

            //行头

            HSSFCell headCell0 = (HSSFCell)headRow.GetCell(0);
            headCell0.SetCellValue("检测编号");
            headCell0.CellStyle = cellStyle;

            HSSFCell headCell1 = (HSSFCell)headRow.GetCell(1);
            headCell1.SetCellValue("样品名称");
            headCell1.CellStyle = cellStyle;

            HSSFCell headCell2 = (HSSFCell)headRow.GetCell(2);
            headCell2.SetCellValue("规格(mm)");
            headCell2.CellStyle = cellStyle;

            HSSFCell headCell3 = (HSSFCell)headRow.GetCell(3);
            headCell3.SetCellValue("任务说明(检测结果%)");
            headCell3.CellStyle = cellStyle;

            for (int i = 0; i < 12; i++)
            {
                HSSFCell headCell = (HSSFCell)headRow2.GetCell(3 + i);
                headCell.SetCellValue(i+1);
                headCell.CellStyle = cellStyle;
            }

            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 0));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 1,1));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1,2, 2));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 3, 14));


            cellStyle = Getcellstyle(wb,false);
            int rowindexstart=2;

            //样式
            if (list != null&&list.Count>0)
            {
                foreach (var item in list)
                {
                    if(item.Items!=null&&item.Items.Count>0)
                    {
                        int count=(int)Math.Ceiling((double)item.Items.Count/12.0);
                        for(int i=0;i<count;i++)
                        {
                            HSSFRow dataRow = (HSSFRow)sheet.CreateRow(rowindexstart); //航头第二行
                            HSSFCell datacell = (HSSFCell)dataRow.CreateCell(0);
                            datacell.SetCellValue(item.ID);
                            datacell.CellStyle = cellStyle;
                            datacell = (HSSFCell)dataRow.CreateCell(1);
                            datacell.SetCellValue(item.Name);
                            datacell.CellStyle = cellStyle;
                            datacell = (HSSFCell)dataRow.CreateCell(2);
                            datacell.SetCellValue(item.Spec);
                            datacell.CellStyle = cellStyle;
                            for(int j=0+i*12;j<(i+1)*12;j++)
                            {
                                if(j>item.Items.Count-1)
                                {
                                    //只设置保持样式一致
                                    datacell = (HSSFCell)dataRow.CreateCell(j + 3 - i * 12);
                                    datacell.CellStyle = cellStyle;
                                }
                                else
                                {
                                    datacell = (HSSFCell)dataRow.CreateCell(j + 3 - i * 12);
                                    datacell.SetCellValue(item.Items[j]);
                                    datacell.CellStyle = cellStyle;
                                }

                            }

                            rowindexstart++;
                        }
                    }
                }
            }

            using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create))
            {
                wb.Write(fs);
            }
            MessageBox.Show("导出成功！");
        }

        public enum stylexls
        {

            头,

            url,

            时间,

            数字,

            钱,

            百分比,

            中文大写,

            科学计数法,

            默认

        }


        static ICellStyle Getcellstyle(IWorkbook wb,bool aligncenter=true, stylexls str=stylexls.默认)
        {

            ICellStyle cellStyle = wb.CreateCellStyle();




            //定义几种字体

            //也可以一种字体，写一些公共属性，然后在下面需要时加特殊的

            IFont font12 = wb.CreateFont();

            font12.FontHeightInPoints = 10;

            font12.FontName = "微软雅黑";






            IFont font = wb.CreateFont();

            font.FontName = "微软雅黑";

            //font.Underline = 1;下划线







            IFont fontcolorblue = wb.CreateFont();

            fontcolorblue.Color = HSSFColor.OliveGreen.Black.Index;

            fontcolorblue.IsItalic = true;//下划线

            fontcolorblue.FontName = "微软雅黑";







            //边框

            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;

            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            //边框颜色

            cellStyle.BottomBorderColor = HSSFColor.OliveGreen.Black.Index;

            cellStyle.TopBorderColor = HSSFColor.OliveGreen.Black.Index;




            //背景图形，我没有用到过。感觉很丑

            //cellStyle.FillBackgroundColor = HSSFColor.OLIVE_GREEN.BLUE.index;

            //cellStyle.FillForegroundColor = HSSFColor.OLIVE_GREEN.BLUE.index;

            cellStyle.FillForegroundColor = HSSFColor.White.Index;

            // cellStyle.FillPattern = FillPatternType.NO_FILL;

            cellStyle.FillBackgroundColor = HSSFColor.Maroon.Index;



            //水平对齐
            if (aligncenter)
            {
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            }
            else
            {
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            }




            //垂直对齐

            cellStyle.VerticalAlignment = VerticalAlignment.Center;




            //自动换行

            cellStyle.WrapText = true;




            //缩进;当设置为1时，前面留的空白太大了。希旺官网改进。或者是我设置的不对

            cellStyle.Indention = 0;


            return cellStyle;

            //上面基本都是设共公的设置

            //下面列出了常用的字段类型

            switch (str)
            {

                case stylexls.头:

                    // cellStyle.FillPattern = FillPatternType.LEAST_DOTS;

                    cellStyle.SetFont(font12);

                    break;

                case stylexls.时间:

                    IDataFormat datastyle = wb.CreateDataFormat();




                    cellStyle.DataFormat = datastyle.GetFormat("yyyy/mm/dd");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.数字:

                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.钱:

                    IDataFormat format = wb.CreateDataFormat();

                    cellStyle.DataFormat = format.GetFormat("￥#,##0");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.url:

                    fontcolorblue.Underline = FontUnderlineType.Single;

                    cellStyle.SetFont(fontcolorblue);

                    break;

                case stylexls.百分比:

                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.中文大写:

                    IDataFormat format1 = wb.CreateDataFormat();

                    cellStyle.DataFormat = format1.GetFormat("[DbNum2][$-804]0");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.科学计数法:

                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00E+00");

                    cellStyle.SetFont(font);

                    break;

                case stylexls.默认:

                    cellStyle.SetFont(font);

                    break;

            }

            return cellStyle;
        }
 

    }
}
