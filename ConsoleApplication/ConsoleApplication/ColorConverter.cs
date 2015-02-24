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
            double tintConv = (tint/1000)*0.01;
            Console.WriteLine("Tint: " + tintConv);
            HSLColor hsl = RGB_to_HSL(c);
            hsl.Luminance = (hsl.Luminance * (tintConv * 0.01)) + (1- (tintConv * 0.01));
            c = HSL_to_RGB(hsl);
            
            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");

            /*double convTint = tint / 1000;
            Color temp = new Color();
            double r = 0, g = 0, b = 0;

            r = ((1 - (60*0.01)) * (255 - c.R)) + c.R;
            g = ((1 - (60*0.01)) * (255 - c.G)) + c.G;
            b = ((1 - (60*0.01)) * (255 - c.B)) + c.B;

            temp = Color.FromArgb((int)(r), (int)(g), (int)(b));
            
            return temp.R.ToString("X2") + temp.G.ToString("X2") + temp.B.ToString("X2");*/
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
            Console.WriteLine("Shade: " + shade);
            int hexValue = (int)(shade/1000);

            double convShade = (double) Convert.ToInt32(hexValue.ToString()) / 255;

            Console.WriteLine("convShade: " + convShade);
            
            int rR = Convert.ToInt16(c.R);
            int gG = Convert.ToInt16(c.G);
            int bB = Convert.ToInt16(c.B);
            
            /*HSLColor newHsl = CreateHSL(c.R, c.G, c.B);
            HSLColor newHsl = CreateHSL(r, g, b);
            Color cs = Color.FromArgb(24, 98, 118);
            HSLColor cHsl = RGB_to_HSL(c);
            Console.WriteLine(newHsl.Hue + " " + newHsl.Saturation + " " + newHsl.Luminance);
            Console.WriteLine(cHsl.Hue + " " + cHsl.Saturation + " " + cHsl.Luminance);*/

/*          Color outColor = c;
            double r = 0, g = 0, b = 0;

            Console.WriteLine(rR + " " + gG + " " + bB);
            r = (1 - (convShade*0.01)) * rR;
            g = (1 - (convShade*0.01)) * gG;
            b = (1 - (convShade*0.01)) * bB;

            Console.WriteLine(r + " " + g + " " + b);
            outColor = Color.FromArgb((int)Math.Round(r), (int)Math.Round(g), (int)Math.Round(b));

            Console.WriteLine("temp: " + outColor.R.ToString("X2") + outColor.G.ToString("X2") + outColor.B.ToString("X2"));*/


            Color outColor = c;// ColorTranslator.FromHtml("#BBE0E3");

            //HSLColor hsl = CreateHSL(rR, gG, bB);

            HSLColor hsl = RGB_to_HSL(outColor);
            //Console.WriteLine(hsl.toString());
            hsl.Luminance *= convShade;
            //Console.WriteLine(hsl.toString());

            outColor = HSL_to_RGB(hsl);

            Console.WriteLine("R: " + outColor.R + ", G: " + outColor.G + ", B: " + outColor.B);
            Console.WriteLine(outColor.R.ToString("X2") + outColor.G.ToString("X2") + outColor.B.ToString("X2"));
            return outColor.R.ToString("X2") + outColor.G.ToString("X2") + outColor.B.ToString("X2");

        }

        public string SetSaturation(Color c, double saturation)
        {
            double convSaturation = (saturation / 1000)*0.01;
            Console.WriteLine("Saturation: " + convSaturation);
            int rR = Convert.ToInt16(c.R);
            int gG = Convert.ToInt16(c.G);
            int bB = Convert.ToInt16(c.B);
            HSLColor hsl = CreateHSL(rR, gG, bB);

            Console.WriteLine(hsl.toString());
            hsl.Saturation = hsl.Saturation*convSaturation;
            Console.WriteLine(hsl.toString());
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
