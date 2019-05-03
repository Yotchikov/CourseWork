using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphLibrary
{
    static class VisioColor
    {
        public static string ColorToRgb(string color)
        {
            switch(color)
            {
                case "white":
                    return "RGB(255, 255, 255)";
                case "red":
                    return "RGB(255,0,0)";
                case "green":
                    return "RGB(0,255,0)";
                case "aliceblue":
                    return "RGB(240, 248, 255)";
                case "antiquewhite":
                    return "RGB(250, 235, 215)";
                case "aquamarine":
                    return "RGB(127, 255, 212)";
                case "azure":
                    return "RGB(240, 255, 255)";
                case "beige":
                    return "RGB(245, 245, 220)";
                case "black":
                    return "RGB(0,0,0)";
                case "blue":
                    return "RGB(0,0,255)";
                case "blueviolet":
                    return "RGB(138, 43, 226)";
                case "brown":
                    return "RGB(165, 42, 42)";
                case "burlywood":
                    return "RGB(222, 184, 135)";
                case "cadetblue":
                    return "RGB(95, 158, 160)";
                case "chartreuse":
                    return "RGB(127, 255, 0)";
                case "chocolate":
                    return "RGB(210, 105, 30)";
                case "coral":
                    return "RGB(255, 127, 80)";
                case "cornflowerblue":
                    return "RGB(100, 149, 237)";
                case "cornsilk":
                    return "RGB(255, 248, 220)";
                case "crimson":
                    return "RGB(220, 20, 60)";
                case "cyan":
                    return "RGB(0, 255, 255)";
                case "darkgoldenrod":
                    return "RGB(184, 134, 11)";
                case "darkgreen":
                    return "RGB(0,100,0)";
                case "darkolivegreen":
                    return "RGB(85,107,47)";
                case "darkorange":
                    return "RGB(255,140,0)";
                case "darkorchild":
                    return "RGB(153,50,204)";
                default:
                    throw new ArgumentException("Обнаружен недопустимый цвет");
            }
        }
    }
}
