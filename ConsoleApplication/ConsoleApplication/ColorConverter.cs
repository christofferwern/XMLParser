using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using DocumentFormat.OpenXml.Drawing;

namespace ConsoleApplication
{
    public class ColorConverter
    {
        public ColorConverter(){}


        public string SetTint(Color c, double tint)
        {

            //Console.WriteLine("Tint: " + tint);
            int hexValue = (int)(tint / 1000);

            double tintConv = (double)Convert.ToInt32(hexValue.ToString(), 16) / 255;

            //Console.WriteLine("tintConv: " + tintConv);

            double r = (double)c.R / 255;
            double g = (double)c.G / 255;
            double b = (double)c.B / 255;

            double rLin = RGB_to_linearRGB(r);
            double gLin = RGB_to_linearRGB(g);
            double bLin = RGB_to_linearRGB(b);

            //Console.WriteLine("Linear R: " + rLin + "\nLinear G: " + gLin + "\nLinear B: " + bLin);

            /**TINT**/

            if (tintConv > 0)
                rLin = rLin * (1 - tintConv) + tintConv;
            else
                rLin = rLin * (1 + tintConv);

            if (tintConv > 0)
                gLin = gLin * (1 - tintConv) + tintConv;
            else
                gLin = gLin * (1 + tintConv);

            if (tintConv > 0)
                bLin = bLin * (1 - tintConv) + tintConv;
            else
                bLin = bLin * (1 + tintConv);

            r = linearRGB_to_RGB(rLin);
            g = linearRGB_to_RGB(gLin);
            b = linearRGB_to_RGB(bLin);

            Console.WriteLine("R: " + (r * 255) + ", G: " + (g * 255) + ", B: " + (b * 255));

            Color outColor = Color.FromArgb((int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));

            return outColor.R.ToString("X2") + outColor.G.ToString("X2") + outColor.B.ToString("X2");
        }
        
        private double RGB_to_linearRGB(double val){

            if (val < 0.0)
                return 0.0;
            if (val <= 0.04045)
                return val / 12.92;
            if (val <= 1.0)
                return (double) Math.Pow(((val + 0.055) / 1.055),2.4);
            
            return 1.0;
        }

        private double linearRGB_to_RGB(double val)
        {

            if (val < 0.0)
                return 0.0;
            if (val <= 0.0031308)
                return val * 12.92;
            if (val < 1.0)
                return (1.055 * Math.Pow(val, (1.0 / 2.4))) - 0.055;

            return 1.0;
            /*
            Public Function linearRGB_to_sRGB(ByVal value As Double) As Double
                If value < 0.0# Then Return 0.0#
                If value <= 0.0031308# Then Return value * 12.92
                If value < 1.0# Then Return 1.055 * (value ^ (1.0# / 2.4)) - 0.055
                Return 1.0#
            End Function
            */

        }

        private HSLColor CreateHSL(int r, int g, int b)
        {
            HSLColor hsl = new HSLColor();

            double R = (double)r / 255;
            double G = (double)g / 255;
            double B = (double)b / 255;
            double max = 0, min = 0;
            double H = 0, S = 0, L = 0;
            bool rBool = false, gBool = false, bBool = false;

            //find max
            if ((R > G) && (R > B)){
                max = R;
                rBool = true;
            }
            else if ((G > R) && (G > B)){
                max = G;
                gBool = true;
            }
            else if ((B > G) && (B > R))
            {
                max = B;
                bBool = true;
            }

            if ((R < G) && (R < B))
                min = R;
            else if ((G < R) && (G < B))
                min = G;
            else if ((B < G) && (B < R))
                min = B;

            //set Luminance
            hsl.Luminance = (min + max) / 2;

            if (min == max)
            {
                hsl.Saturation = 0;
                hsl.Hue = 0;
            }

            //set saturation
            if (hsl.Luminance < 0.5)
                hsl.Saturation = (max - min) / (max + min);
            else if (hsl.Luminance > 0.5)
                hsl.Saturation = (max - min) / (2.0 - max - min);

            if (rBool)
            {
                H = ((G - B) / (max - min))*60.0;
                Console.WriteLine("red");
            }
            if (gBool)
            {
                H = ((2.0 + ((B - R) / (max - min)))*60.0);
                Console.WriteLine("green");
            }
            if (bBool)
            {
                H= (4.0 + (R - G) / (max - min))*60.0;
                Console.WriteLine("blue");
            }

            if (H < 0)
                H += 360;

            H = Math.Round(H);
            H = H / 360;

            hsl.Hue = H;

            return hsl;
        }

        public string SetShade(Color c, double shade)
        {
            //Console.WriteLine("Shade: " + shade);
            int hexValue = (int)(shade/1000);

            double convShade = (double) Convert.ToInt32(hexValue.ToString(), 16) / 255;

            //Console.WriteLine("convShade: " + convShade);

            double r = (double) c.R / 255;
            double g = (double) c.G / 255;
            double b = (double) c.B / 255;

            double rLin  = RGB_to_linearRGB(r);
            double gLin = RGB_to_linearRGB(g);
            double bLin = RGB_to_linearRGB(b);

            //Console.WriteLine("Linear R: " + rLin + "\nLinear G: " + gLin + "\nLinear B: " + bLin);

            /**SHADE**/
            if ((rLin * convShade) < 0)
                rLin = 0;
            if ((rLin * convShade) > 1)
                rLin = 0;
            else
                rLin *= convShade;

            if ((gLin * convShade) < 0)
                gLin = 0;
            if ((gLin * convShade) > 1)
                gLin = 0;
            else
                gLin *= convShade;

            if ((bLin * convShade) < 0)
                bLin = 0;
            if ((bLin * convShade) > 1)
                bLin = 0;
            else
                bLin *= convShade;

            /**SHADE**/

            r = linearRGB_to_RGB(rLin);
            g = linearRGB_to_RGB(gLin);
            b = linearRGB_to_RGB(bLin);

            //Console.WriteLine("R: " + (r*255) + ", G: " + (g*255) + ", B: " + (b*255) );

            Color outColor = Color.FromArgb((int)Math.Round(r * 255), (int)Math.Round(g * 255), (int)Math.Round(b * 255));

            return outColor.R.ToString("X2") + outColor.G.ToString("X2") + outColor.B.ToString("X2");

        }

        public string SetSaturation(Color c, double saturation)
        {
            double convSaturation = (saturation / 1000)*0.01;
            //Console.WriteLine("Saturation: " + convSaturation);
            HSLColor hsl = RGB_to_HSL(c); // CreateHSL(rR, gG, bB);

            //Console.WriteLine(hsl.toString());
            hsl.Saturation *= convSaturation;
            //Console.WriteLine(hsl.toString());
            c = HSL_to_RGB(hsl);

            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        public string SetBrightness(Color c, double brightness)
        {
            
            double convBrightness = brightness / 100000;

            HSLColor hsl = RGB_to_HSL(c);

            hsl.Luminance = convBrightness;
            c = HSL_to_RGB(hsl);

            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        private HSLColor RGB_to_HSL(Color c)
        {
            HSLColor hsl = new HSLColor();

            hsl.Hue = c.GetHue() / 360.0;
            hsl.Luminance = c.GetBrightness();
            hsl.Saturation = c.GetSaturation();

            return hsl;
        }

        private Color HSL_to_RGB(HSLColor hslColor)
        {

            double r = 0, g = 0, b = 0;
            double temp1, temp2;


            if (hslColor.Luminance == 0)
                r = g = b = 0;
            else
            {
                if (hslColor.Saturation == 0)
                {
                    r = g = b = hslColor.Luminance;
                }
                else
                {
                    temp2 = ((hslColor.Luminance <= 0.5) ? hslColor.Luminance * (1.0 + hslColor.Saturation) : hslColor.Luminance + hslColor.Saturation - (hslColor.Luminance * hslColor.Saturation));
                    temp1 = 2.0 * hslColor.Luminance - temp2;

                    double[] t3 = new double[] { hslColor.Hue + 1.0 / 3.0, hslColor.Hue, hslColor.Hue - 1.0 / 3.0 };
                    double[] clr = new double[] { 0, 0, 0 };
                    for (int i = 0; i < 3; i++)
                    {
                        if (t3[i] < 0)
                            t3[i] += 1.0;
                        if (t3[i] > 1)
                            t3[i] -= 1.0;

                        if (6.0 * t3[i] < 1.0)
                            clr[i] = temp1 + (temp2 - temp1) * t3[i] * 6.0;
                        else if (2.0 * t3[i] < 1.0)
                            clr[i] = temp2;
                        else if (3.0 * t3[i] < 2.0)
                            clr[i] = (temp1 + (temp2 - temp1) * ((2.0 / 3.0) - t3[i]) * 6.0);
                        else
                            clr[i] = temp1;
                    }
                    r = clr[0];
                    g = clr[1];
                    b = clr[2];
                }
            }

            return Color.FromArgb((int)(255 * r), (int)(255 * g), (int)(255 * b));
            
        }

    }
    
}
