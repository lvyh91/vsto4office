using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Lvyh.Core;
using Lvyh.EDomain;
//using Microsoft.Office.Tools.Excel;

namespace Homework02
{
    public class RibbonFunctionGrab
    {
        //private static RibbonFunctionGrab _func = new RibbonFunctionGrab();
        //public static RibbonFunctionGrab Default
        //{
        //    get { return _func; }
        //}
        //private RibbonFunctionGrab() { }


        public static void Grab()
        {
            try
            {
                string url = @"http://www.matrix67.com/blog/feed";
                List<string> txt = MyGrab.GetContents(url);

                Excel.Worksheet sheet = Geter.App.ActiveSheet;

                sheet.Cells[2, 1] = "序号";
                sheet.Cells[2, 2] = "标题";
                for (int i = 1; i <= txt.Count; i++)
                {
                    sheet.Cells[i + 2, 1] = i;
                    sheet.Cells[i + 2, 2] = txt[i - 1];
                }
                sheet.Range["a1:b1"].EntireColumn.AutoFit();
                sheet.Range["a2"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                sheet.Cells[1, 1] = url;
                sheet.Range["a1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
            catch (Exception ex)
            {
                Logger.Default.Log(ex);
            }
        }
    }
}
