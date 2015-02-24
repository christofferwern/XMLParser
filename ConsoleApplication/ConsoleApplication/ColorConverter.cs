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


        public string SetTint(string s, double tint)
        {
            Color c = ColorTranslator.FromHtml("#" + s);

            return SetTint(c,tint);
        }

        public string SetShade(string s, double shade)
        {
            Color c = ColorTranslator.FromHtml("#" + s);

            return SetShade(c, shade);
        }

        public string SetSaturation(string s, double saturation)
        {
            Color c = ColorTranslator.FromHtml("#" + s);

            return SetSaturation(c, saturation);
        }

        public string SetBrightness(string s, double luminance)
        {
            Color c = ColorTranslator.FromHtml("#" + s);

            return SetBrightness(c, luminance);
        }

        public string SetTint(Color c, double tint)
        {
            Color temp = new Color();

            double convTint = tint / 1000;
            double r = 0, g = 0, b = 0;

            r = ((1 - (convTint * 0.01)) * (255 - c.R)) + c.R;
            g = ((1 - (convTint * 0.01)) * (255 - c.G)) + c.G;
            b = ((1 - (convTint * 0.01)) * (255 - c.B)) + c.B;

            temp = Color.FromArgb((int)(r), (int)(g), (int)(b));

            return temp.R.ToString("X2") + temp.G.ToString("X2") + temp.B.ToString("X2");
        }

        public string SetShade(Color c, double tint)
        {
            Color temp = new Color();

            double convTint = tint / 1000;
            double r = 0, g = 0, b = 0;

            r = (1 - (convTint * 0.01)) * c.R;
            g = (1 - (convTint * 0.01)) * c.G;
            b = (1 - (convTint * 0.01)) * c.B;

            temp = Color.FromArgb((int)(r),(int)(g), (int)(b));
            return temp.R.ToString("X2") + temp.G.ToString("X2") + temp.B.ToString("X2");
        }

        public string SetSaturation(Color c, double saturation)
        {

            double convSaturation = saturation / 100000;

            HSLColor hsl = RGB_to_HSL(c);

            hsl.Saturation = convSaturation;
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
