﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Homework02
{
    public class Geter
    {
        public static Excel.Application App
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }
    }
}
