using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Lvyh.EDomain;

namespace Homework02
{
    public partial class ERibbon1
    {
        private void ERibbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnGrab_Click(object sender, RibbonControlEventArgs e)
        {
            new Action(RibbonFunctionGrab.Grab).BeginInvoke(null, null);
        }
    }
}
