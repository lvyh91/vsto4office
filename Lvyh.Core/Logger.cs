using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lvyh.Core
{
    public class Logger
    {
        private static Logger _logger = new Logger();
        public static Logger Default
        {
            get { return _logger; }
        }
        private Logger() { }


        public void Log(Exception ex)
        {

        }

        public void Log(string text)
        {

        }
    }
}
