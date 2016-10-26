using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSRS_Styler
{
    public class RSTheme
    {
        public string[] dataPoints { get; set; }
        public string good { get; set; }
        public string bad { get; set; }
        public string neutral { get; set; }
        public string none { get; set; }
        public string background { get; set; }
        public string foreground { get; set; }
        public string mapBase { get; set; }
        public string panelBackground { get; set; }
        public string panelForeground { get; set; }
        public string panelAccent { get; set; }
        public string tableAccent { get; set; }
        public string altBackground { get; set; }
        public string altForeground { get; set; }
        public string altMapBase { get; set; }
        public string altPanelBackground { get; set; }
        public string altPanelForeground { get; set; }
        public string altPanelAccent { get; set; }
        public string altTableAccent { get; set; }


        private void InitTheme()
        {
            dataPoints = new String[12] { "#0072c6", "#f68c1f", "#269657", "#dd5900", "#5b3573", "#22bdef", "#b4009e", "#008274", "#fdc336", "#ea3c00", "#00188f", "#9f9f9f" };
            good = "#85ba00";
            bad = "#e90000";
            neutral = "#edb327";
            none = "#333";
            background = "#fff";
            foreground = "#222";
            mapBase = "#00aeef";
            panelBackground = "#f6f6f6";
            panelForeground = "#222";
            panelAccent = "#00aeef";
            tableAccent = "#00aeef";
            altBackground = "#f6f6f6";
            altForeground = "#000";
            altMapBase = "#f68c1f";
            altPanelBackground = "#235378";
            altPanelForeground = "#fff";
            altPanelAccent = "#fdc336";
            altTableAccent = "#fdc336";
        }

        public void Reset()
        {
            InitTheme();
        }

        public RSTheme()
        {
            InitTheme();
        }
    }
}
