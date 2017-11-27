using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Lvyh.PDomain
{
    public static class ShapeExt
    {
        public static void ClearMargin(this PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
            {
                string msg = "下列形状不包含文本域：";
                msg += "\r\n ID：" + shape.Id;
                msg += "\r\n Name：" + shape.Name;
                throw new Exception(msg);
            }

            shape.TextFrame.MarginLeft = 0;
            shape.TextFrame.MarginTop = 0;
            shape.TextFrame.MarginRight = 0;
            shape.TextFrame.MarginBottom = 0;
        }


        public static List<PowerPoint.Shape> ToList(this PowerPoint.ShapeRange shapeRange)
        {
            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();

            foreach (PowerPoint.Shape shape in shapeRange)
            {
                shapes.Add(shape);
            }

            return shapes;
        }

        public static void ForEach(this List<PowerPoint.Shape> shapes, Action<PowerPoint.Shape> action, int startIndex = 0, int endIndex = 0)
        {
            if (endIndex == 0)
            {
                endIndex = shapes.Count;
            }

            for (int i = startIndex; i <= endIndex; i++)
            {
                action(shapes[i]);
            }
        }

        /// <summary>
        /// 逆序循环操作列表中的对象，下标从0开始
        /// </summary>
        /// <param name="shapes"></param>
        /// <param name="action"></param>
        /// <param name="endIndex"></param>
        /// <param name="startIndex"></param>
        public static void ForEach2(this List<PowerPoint.Shape> shapes, Action<PowerPoint.Shape> action, int endIndex = 0, int startIndex = 0)
        {
            if (endIndex == 0)
            {
                endIndex = shapes.Count - 1;
            }

            for (int i = endIndex; i >= startIndex; i--)
            {
                action(shapes[i]);
            }
        }
    }
}
