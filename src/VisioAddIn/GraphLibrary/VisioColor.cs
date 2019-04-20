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
                default:
                    throw new ArgumentException("Обнаружен недопустимый цвет");
            }
        }
    }
}
