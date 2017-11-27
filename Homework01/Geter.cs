using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Homework01
{
    internal class Geter
    {
        internal static PowerPoint.Application App
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }
    }
}
