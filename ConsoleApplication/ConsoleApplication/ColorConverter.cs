using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    public class ColorConverter
    {
        public ColorConverter(){}


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

        private Color HSL_to_RGB(HSLColor hsl)
        {
            double r = 0, g = 0, b = 0;
            double temp1, temp2;


            if (hsl.Luminance == 0)
                r = g = b = 0;
            else
            {
                if (hsl.Saturation == 0)
                {
                    r = g = b = hsl.Luminance;
                }
                else
                {
                    temp2 = ((hsl.Luminance <= 0.5) ? hsl.Luminance * (1.0 + hsl.Saturation) : hsl.Luminance + hsl.Saturation - (hsl.Luminance * hsl.Saturation));
                    temp1 = 2.0 * hsl.Luminance - temp2;

                    double[] t3 = new double[] { hsl.Hue + 1.0 / 3.0, hsl.Hue, hsl.Hue - 1.0 / 3.0 };
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
