using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Lvyh.EDomain
{
   public class MyGrab
    {
       public static List<string> GetContents(string url)
       {
           List<string> titles = null;

           //第一句，抓
           XElement xml = XElement.Load(url);

           //第二句，取
           //string txt = "------------ 数学大冒险 -------------" + "\r\n";
           titles = xml.Element("channel").Elements("item")
                         .Select((ele) => ele.Element("title").Value)
                         .ToList();
           //.Select((m, index1) => txt += index1.ToString() + ":" + m.Element("title").Value + "\r\n")
           //.Where((n, index2) => index2 < 5)
           //.ToList();

           //
           return titles;

       }
    }
}
