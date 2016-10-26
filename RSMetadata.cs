using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSRS_Styler
{
    class RSMetadata
    {
        public string GenerateMetadataFile (string pName, string pLogoName, string pColorsName)
        {
            return String.Format("<?xml version=\"1.0\" encoding=\"utf-8\"?><SystemResourcePackage xmlns = \"http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata\" type = \"UniversalBrand\" version = \"2.0.2\" name = \"{0}\"><Contents><Item key = \"colors\" path = \"{1}\"/><Item key = \"logo\" path = \"{2}\" /></Contents></SystemResourcePackage> ", pName, pColorsName, pLogoName);
        }
    }
}
