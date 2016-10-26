using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using System.Diagnostics;
using System.IO.Compression;

namespace SSRS_Styler
{
    public partial class Form1 : Form
    {

        RSBranding rsbrand = new RSBranding();
        RSMetadata rsmeta = new RSMetadata();

        private void InitColorPickers()
        {
            colPrimary.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primary);
            colPrimaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primaryAlt);
            colPrimaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primaryAlt2);
            colPrimaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primaryAlt3);
            colPrimaryAlt4.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primaryAlt4);
            colPrimaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.primaryContrast);

            colSecondary.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.secondary);
            colSecondaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.secondaryAlt);
            colSecondaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.secondaryAlt2);
            colSecondaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.secondaryAlt3);
            colSecondaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.secondaryContrast);

            colNeutralPrimary.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralPrimary);
            colNeutralPrimaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralPrimaryAlt);
            colNeutralPrimaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralPrimaryAlt2);
            colNeutralPrimaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralPrimaryAlt3);
            colNeutralPrimaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralPrimaryContrast);

            colNeutralSecondary.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralSecondary);
            colNeutralSecondaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralSecondaryAlt);
            colNeutralSecondaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralSecondaryAlt2);
            colNeutralSecondaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralSecondaryAlt3);
            colNeutralSecondaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralSecondaryContrast);

            colNeutralTertiary.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralTertiary);
            colNeutralTertiaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralTertiaryAlt);
            colNeutralTertiaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralTertiaryAlt2);
            colNeutralTertiaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralTertiaryAlt3);
            colNeutralTertiaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.neutralTertiaryContrast);

            colDanger.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.danger);
            colDangerContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.dangerContrast);
            colSuccess.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.success);
            colSuccessContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.successContrast);
            colInfo.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.info);
            colInfoContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.infoContrast);
            colWarning.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.warning);
            colWarningContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.warningContrast);

            colKPIGood.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiGood);
            colKPIGoodContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiGoodContrast);
            colKPIBad.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiBad);
            colKPIBadContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiBadContrast);
            colKPINeutral.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiNeutral);
            colKPINeutralContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiNeutralContrast);
            colKPINone.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiNone);
            colKPINoneContrast.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.rsinterface.kpiNoneContrast);

            colDatapoints1.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[0]);
            colDatapoints2.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[1]);
            colDatapoints3.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[2]);
            colDatapoints4.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[3]);
            colDatapoints5.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[4]);
            colDatapoints6.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[5]);
            colDatapoints7.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[6]);
            colDatapoints8.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[7]);
            colDatapoints9.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[8]);
            colDatapoints10.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[9]);
            colDatapoints11.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[10]);
            colDatapoints12.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.dataPoints[11]);

            colBad.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.bad);
            colGood.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.good);
            colNeutral.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.neutral);
            colNone.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.none);

            colBackground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.background);
            colForeground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.foreground);
            colMapBase.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.mapBase);
            colPanelBackground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.panelBackground);
            colPanelForeground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.panelForeground);
            colPanelAccent.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.panelAccent);
            colTableAccent.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.tableAccent);

            colAltBackground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altBackground);
            colAltForeground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altForeground);
            colAltMapBase.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altMapBase);
            colAltPanelBackground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altPanelBackground);
            colAltPanelForeground.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altPanelForeground);
            colAltPanelAccent.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altPanelAccent);
            colAltTableAccent.selectedColor = (Color)ColorTranslator.FromHtml(rsbrand.theme.altTableAccent);

            txtBrandPackageName.Text = rsbrand.name;
            txtBrandPackageVersion.Text = rsbrand.version;
        }

        public Form1()
        {
            InitializeComponent();
            InitColorPickers();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you really want to reset?\nAll unsaved settings will be lost!", "SSRS Styler", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                rsbrand.Reset();
                txtColorsFilename.Text = "colors.json";
                txtLogoFilename.Text = "logo.png";
                InitColorPickers();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            saveFileDialog1.FileName = txtBrandPackageName.Text;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (pbLogo.Image != null)
                {
                    string tempPath = Path.GetTempPath();
                    string zipName = String.Format(@"{0}\{1}.zip", tempPath, txtBrandPackageName.Text);
                    string directoryName = String.Format(@"{0}\{1}", tempPath, txtBrandPackageName.Text);
                    string jsonName = String.Format(@"{0}\{1}\{2}", tempPath, txtBrandPackageName.Text, txtColorsFilename.Text);
                    string metadataName = String.Format(@"{0}\{1}\metadata.xml", tempPath, txtBrandPackageName.Text);
                    string imageName = String.Format(@"{0}\{1}\{2}", tempPath, txtBrandPackageName.Text, txtLogoFilename.Text);
                    string destFileName = saveFileDialog1.FileName;

                    string json = JsonConvert.SerializeObject(rsbrand).Replace("rsinterface", "interface");
                    string metadata = rsmeta.GenerateMetadataFile(txtBrandPackageName.Text, txtLogoFilename.Text, txtColorsFilename.Text);

                    if (!Directory.Exists(directoryName))
                    {
                        Directory.CreateDirectory(String.Format(@"{0}\{1}", tempPath, txtBrandPackageName.Text));
                    }

                    File.WriteAllText(jsonName, json);
                    File.WriteAllText(metadataName, metadata);
                    pbLogo.Image.Save(imageName, System.Drawing.Imaging.ImageFormat.Png);

                    if (File.Exists(zipName))
                    {
                        File.Delete(zipName);
                    }
                    ZipFile.CreateFromDirectory(directoryName, zipName);

                    //FGE: after Zip-File Creation is finished copy to destination folder
                    File.Copy(zipName, destFileName, true);

                    //FGE: cleanup
                    Directory.Delete(directoryName, true);
                    File.Delete(zipName);
                }
                else
                {
                    MessageBox.Show("You have not selected an image as Logo. Please select an image.", "No image", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void colPrimary_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primary = colPrimary.selectedColorHex;
        }

        private void colPrimaryAlt_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primaryAlt = colPrimaryAlt.selectedColorHex;
        }

        private void colPrimaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primaryAlt2 = colPrimaryAlt2.selectedColorHex;
        }

        private void colPrimaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primaryAlt3 = colPrimaryAlt3.selectedColorHex;
        }

        private void colPrimaryAlt4_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primaryAlt4 = colPrimaryAlt4.selectedColorHex;
        }

        private void colPrimaryContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.primaryContrast = colPrimaryContrast.selectedColorHex;
        }

        private void colSecondary_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.secondary = colSecondary.selectedColorHex;
        }

        private void colSecondaryAlt_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.secondaryAlt = colSecondaryAlt.selectedColorHex;
        }

        private void colSecondaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.secondaryAlt2 = colSecondaryAlt2.selectedColorHex;
        }

        private void colSecondaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.secondaryAlt3 = colSecondaryAlt3.selectedColorHex;
        }

        private void colSecondaryContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.secondaryContrast = colSecondaryContrast.selectedColorHex;
        }

        private void colNeutralPrimary_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralPrimary = colNeutralPrimary.selectedColorHex;
        }

        private void colNeutralPrimaryAlt_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralPrimaryAlt = colNeutralPrimaryAlt.selectedColorHex;
        }

        private void colNeutralPrimaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralPrimaryAlt2 = colNeutralPrimaryAlt2.selectedColorHex;
        }

        private void colNeutralPrimaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralPrimaryAlt3 = colNeutralPrimaryAlt3.selectedColorHex;
        }

        private void colNeutralPrimaryContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralPrimaryContrast = colNeutralPrimaryContrast.selectedColorHex;
        }

        private void colNeutralSecondary_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralSecondary = colNeutralSecondary.selectedColorHex;
        }

        private void colNeutralSecondaryAlt_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralSecondaryAlt = colNeutralSecondaryAlt.selectedColorHex;
        }

        private void colNeutralSecondaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralSecondaryAlt2 = colNeutralSecondaryAlt2.selectedColorHex;
        }

        private void colNeutralSecondaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralSecondaryAlt3 = colNeutralSecondaryAlt3.selectedColorHex;
        }

        private void colNeutralSecondaryContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralSecondaryContrast = colNeutralSecondaryContrast.selectedColorHex;
        }

        private void colNeutralTertiary_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralTertiary = colNeutralTertiary.selectedColorHex;
        }

        private void colNeutralTertiaryAlt_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralTertiaryAlt = colNeutralTertiaryAlt.selectedColorHex;
        }

        private void colNeutralTertiaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralTertiaryAlt2 = colNeutralTertiaryAlt2.selectedColorHex;
        }

        private void colNeutralTertiaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralTertiaryAlt3 = colNeutralTertiaryAlt3.selectedColorHex;
        }

        private void colNeutralTertiaryContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.neutralTertiaryContrast = colNeutralTertiaryContrast.selectedColorHex;
        }

        private void colDanger_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.danger = colDanger.selectedColorHex;
        }

        private void colSuccess_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.success = colSuccess.selectedColorHex;
        }

        private void colWarning_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.warning = colWarning.selectedColorHex;
        }

        private void colInfo_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.info = colInfo.selectedColorHex;
        }

        private void colDangerContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.dangerContrast = colDangerContrast.selectedColorHex;
        }

        private void colSuccessContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.successContrast = colSuccessContrast.selectedColorHex;
        }

        private void colWarningContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.warningContrast = colWarningContrast.selectedColorHex;
        }

        private void colInfoContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.infoContrast = colInfoContrast.selectedColorHex;
        }

        private void colKPIGood_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiGood = colKPIGood.selectedColorHex;
        }

        private void colKPIBad_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiBad = colKPIBad.selectedColorHex;
        }

        private void colKPINeutral_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiNeutral = colKPINeutral.selectedColorHex;
        }

        private void colKPINone_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiNone = colKPINone.selectedColorHex;
        }

        private void colKPIGoodContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiGoodContrast = colKPIGoodContrast.selectedColorHex;
        }

        private void colKPIBadContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiBadContrast = colKPIBadContrast.selectedColorHex;
        }

        private void colKPINeutralContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiNeutralContrast = colKPINeutralContrast.selectedColorHex;
        }

        private void colKPINoneContrast_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.rsinterface.kpiNoneContrast = colKPINoneContrast.selectedColorHex;
        }

        private void colDatapoints1_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[0] = colDatapoints1.selectedColorHex;
        }

        private void colDatapoints2_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[1] = colDatapoints2.selectedColorHex;
        }

        private void colDatapoints3_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[2] = colDatapoints3.selectedColorHex;
        }

        private void colDatapoints4_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[3] = colDatapoints4.selectedColorHex;
        }

        private void colDatapoints5_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[4] = colDatapoints5.selectedColorHex;
        }

        private void colDatapoints6_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[5] = colDatapoints6.selectedColorHex;
        }

        private void colDatapoints7_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[6] = colDatapoints7.selectedColorHex;
        }

        private void colDatapoints8_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[7] = colDatapoints8.selectedColorHex;
        }

        private void colDatapoints9_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[8] = colDatapoints9.selectedColorHex;
        }

        private void colDatapoints10_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[9] = colDatapoints10.selectedColorHex;
        }

        private void colDatapoints11_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[10] = colDatapoints11.selectedColorHex;
        }

        private void colDatapoints12_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.dataPoints[11] = colDatapoints12.selectedColorHex;
        }

        private void colGood_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.good = colGood.selectedColorHex;
        }

        private void colBad_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.bad = colBad.selectedColorHex;
        }

        private void colNeutral_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.neutral = colNeutral.selectedColorHex;
        }

        private void colNone_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.none = colNone.selectedColorHex;
        }

        private void colBackground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.background = colBackground.selectedColorHex;
        }

        private void colForeground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.foreground = colForeground.selectedColorHex;
        }

        private void colMapBase_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altMapBase = colMapBase.selectedColorHex;
        }

        private void colPanelBackground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.panelBackground = colPanelBackground.selectedColorHex;
        }

        private void colPanelAccent_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.panelAccent = colPanelAccent.selectedColorHex;
        }

        private void colTableAccent_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.tableAccent = colTableAccent.selectedColorHex;
        }

        private void colAltBackground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altBackground = colAltBackground.selectedColorHex;
        }

        private void colAltForeground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altForeground = colAltForeground.selectedColorHex;
        }

        private void colAltMapBase_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altMapBase = colAltMapBase.selectedColorHex;
        }

        private void colAltPanelBackground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altPanelBackground = colAltPanelBackground.selectedColorHex;
        }

        private void colAltPanelForeground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altPanelForeground = colAltPanelForeground.selectedColorHex;
        }

        private void colAltPanelAccent_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altTableAccent = colAltPanelAccent.selectedColorHex;
        }

        private void colAltTableAccent_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.altTableAccent = colAltTableAccent.selectedColorHex;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                pbLogo.Image = Image.FromFile(openFileDialog1.FileName);
                if (pbLogo.Image.Width != 80 || pbLogo.Image.Height != 80)
                {
                    MessageBox.Show("The size of the image does not fit. Best size is 80x80 px.\nResizing the image.", "Resizing", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    pbLogo.Image = new Bitmap(pbLogo.Image, new Size(80,80));
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo("http://www.gdsbi.com");
            Process.Start(sInfo);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            InitColorPickers();
        }

        private void colPanelForeground_ColorChanged(object sender, EventArgs e)
        {
            rsbrand.theme.panelForeground = colAltPanelForeground.selectedColorHex;
        }

    }
}
