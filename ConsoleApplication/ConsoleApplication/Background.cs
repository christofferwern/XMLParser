using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Background
    {
        private List<KeyValuePair<int, string>> gradientList = null;
        private string bgColor;
        private float alpha;

        

        private string BgImageLocation;

        public Background()
        {
            bgColor = "ffffff";
            BgImageLocation = "No picture";
        }

        public void setGradientList(List<KeyValuePair<int, string>> list)
        {
            if (gradientList == null)
                gradientList = new List<KeyValuePair<int, string>>();

            gradientList = list;
        }

        public void setBgImageLocation(string loc)
        {
            BgImageLocation = loc;
        }

        public string getBgImageLocation()
        {
            return BgImageLocation;
        }

        public List<KeyValuePair<int, string>> getGradientList()
        {
            if (gradientList != null)
                return gradientList;
            else
                return null;
        }

        public float Alpha
        {
            get { return alpha; }
            set { alpha = 1 - ((value / 1000) / 100);}
        }

        public string BgColor
        {
            get { return bgColor; }
            set { bgColor = value; }
        }
    }
}
