using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Lvyh.PDomain;

namespace Homework01
{
    public partial class PRibbon1
    {
        private void PRibbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnClearMargin_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = Geter.App.ActiveWindow.Selection;
            if (sel.Type== PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                sel.ShapeRange.ToList().ForEach(s => s.ClearMargin());
            }
        }
    }
}
