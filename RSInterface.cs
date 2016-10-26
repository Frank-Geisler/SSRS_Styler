using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSRS_Styler
{
    public class RSInterface
    {
        public string primary { get; set; }
        public string primaryAlt { get; set; }
        public string primaryAlt2 { get; set; }
        public string primaryAlt3 { get; set; }
        public string primaryAlt4 { get; set; }
        public string primaryContrast { get; set; }
        public string secondary { get; set; }
        public string secondaryAlt { get; set; }
        public string secondaryAlt2 { get; set; }
        public string secondaryAlt3 { get; set; }
        public string secondaryContrast { get; set; }
        public string neutralPrimary { get; set; }
        public string neutralPrimaryAlt { get; set; }
        public string neutralPrimaryAlt2 { get; set; }
        public string neutralPrimaryAlt3 { get; set; }
        public string neutralPrimaryContrast { get; set; }
        public string neutralSecondary { get; set; }
        public string neutralSecondaryAlt { get; set; }
        public string neutralSecondaryAlt2 { get; set; }
        public string neutralSecondaryAlt3 { get; set; }
        public string neutralSecondaryContrast { get; set; }
        public string neutralTertiary { get; set; }
        public string neutralTertiaryAlt { get; set; }
        public string neutralTertiaryAlt2 { get; set; }
        public string neutralTertiaryAlt3 { get; set; }
        public string neutralTertiaryContrast { get; set; }
        public string danger { get; set; }
        public string success { get; set; }
        public string warning { get; set; }
        public string info { get; set; }
        public string dangerContrast { get; set; }
        public string successContrast { get; set; }
        public string warningContrast { get; set; }
        public string infoContrast { get; set; }
        public string kpiGood { get; set; }
        public string kpiBad { get; set; }
        public string kpiNeutral { get; set; }
        public string kpiNone { get; set; }
        public string kpiGoodContrast { get; set; }
        public string kpiBadContrast { get; set; }
        public string kpiNeutralContrast { get; set; }
        public string kpiNoneContrast { get; set; }

        private void InitInterface()
        {
            primary = "#bb2124";
            primaryAlt = "#d31115";
            primaryAlt2 = "#671215";
            primaryAlt3 = "#bb2124";
            primaryAlt4 = "#00abee";
            primaryContrast = "#ffffff";
            secondary = "#000000";
            secondaryAlt = "#444444";
            secondaryAlt2 = "#555555";
            secondaryAlt3 = "#777777";
            secondaryContrast = "#ffffff";
            neutralPrimary = "#ffffff";
            neutralPrimaryAlt = "#f4f4f4";
            neutralPrimaryAlt2 = "#e3e3e3";
            neutralPrimaryAlt3 = "#c8c8c8";
            neutralPrimaryContrast = "#ffffff";
            neutralSecondary = "#ffffff";
            neutralSecondaryAlt = "#eaeaea";
            neutralSecondaryAlt2 = "#b7b7b7";
            neutralSecondaryAlt3 = "#acacac";
            neutralSecondaryContrast = "#000000";
            neutralTertiary = "#b7b7b7";
            neutralTertiaryAlt = "#c8c8c8";
            neutralTertiaryAlt2 = "#eaeaea";
            neutralTertiaryAlt3 = "#ffffff";
            neutralTertiaryContrast = "#222222";
            danger = "#bb2124";
            success = "#2b3";
            warning = "#f0ad4e";
            info = "#5bc0de";
            dangerContrast = "#ffffff";
            successContrast = "#ffffff";
            warningContrast = "#ffffff";
            infoContrast = "#ffffff";
            kpiGood = "#4fb443";
            kpiBad = "#de061a";
            kpiNeutral = "#d9b42c";
            kpiNone = "#333";
            kpiGoodContrast = "#fff";
            kpiBadContrast = "#fff";
            kpiNeutralContrast = "#fff";
            kpiNoneContrast = "#fff";
        }
        public RSInterface()
        {
            InitInterface();
        }

        public void Reset()
        {
            InitInterface();
        }
    }
}
