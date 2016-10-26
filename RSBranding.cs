using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSRS_Styler
{
    public class RSBranding
    {
        public string name { get; set; }
        public string version { get; set; }

        public RSInterface rsinterface { get; set; }
        public RSTheme theme { get; set; }

        private void InitializeBranding()
        {
            name = "GDS Premium Brand Package";
            version = "1.0";
        }

        public RSBranding()
        {
            rsinterface = new RSInterface();
            theme = new RSTheme();
            InitializeBranding();
        }

        public void Reset()
        {
            theme.Reset();
            rsinterface.Reset();
            InitializeBranding();
        }
    }
}
