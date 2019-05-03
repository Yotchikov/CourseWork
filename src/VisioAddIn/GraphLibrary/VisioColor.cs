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
                    return "RGB(0, 100, 0)";
                case "darkolivegreen":
                    return "RGB(85, 107, 47)";
                case "darkorange":
                    return "RGB(255, 140, 0)";
                case "darkorchild":
                    return "RGB(153, 50, 204)";
                case "darksalmon":
                    return "RGB(233, 150, 122)";
                case "darkseagreen":
                    return "RGB(143, 188, 143)";
                case "darkslateblue":
                    return "RGB(72, 61, 139)";
                case "darkslategray":
                    return "RGB(47, 79, 79)";
                case "darkslategrey":
                    return "RGB(47, 79, 79)";
                case "darkturquoise":
                    return "RGB(0, 206, 209)";
                case "darkviolet":
                    return "RGB(148, 0, 211)";
                case "deeppink":
                    return "RGB(255, 20, 147)";
                case "deepskyblue":
                    return "RGB(0, 191, 255)";
                case "dimgray":
                    return "RGB(105, 105, 105)";
                case "dimgrey":
                    return "RGB(105, 105, 105)";
                case "dodgerblue":
                    return "RGB(30, 144, 255)";
                case "firebrick":
                    return "RGB(178, 34, 34)";
                case "floralwhite":
                    return "RGB(255, 250, 240)";
                case "forestgreen":
                    return "RGB(34, 139, 34)";
                case "gainsboro":
                    return "RGB(220, 220, 200)";
                case "ghostwhite":
                    return "RGB(248, 248, 255)";
                case "gold":
                    return "RGB(255, 215, 0)";
                case "goldenrod":
                    return "RGB(218, 165, 32)";
                case "gray":
                    return "RGB(192, 192, 192)";
                case "grey":
                    return "RGB(192, 192, 192)";
                default:
                    throw new ArgumentException("Обнаружен недопустимый цвет");
            }
        }
    }
}
