using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using System.Diagnostics;
using System.IO.Compression;
using System.Xml;
using System.Text;

namespace SSRS_Styler
{
    public partial class Form1 : Form
    {

        RSBranding _rsbrand = new RSBranding();
        RSMetadata _rsmeta = new RSMetadata();

        #region UpdateColorInJSonObject
        /*
         * In this region all the glue code can be found that is used to update the JSon-Object each time the
         * color is changed in the Dialog-Control
         * 
         */

        private void colPrimary_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primary = colPrimary.selectedColorHex;
        }

        private void colPrimaryAlt_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primaryAlt = colPrimaryAlt.selectedColorHex;
        }

        private void colPrimaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primaryAlt2 = colPrimaryAlt2.selectedColorHex;
        }

        private void colPrimaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primaryAlt3 = colPrimaryAlt3.selectedColorHex;
        }

        private void colPrimaryAlt4_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primaryAlt4 = colPrimaryAlt4.selectedColorHex;
        }

        private void colPrimaryContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.primaryContrast = colPrimaryContrast.selectedColorHex;
        }

        private void colSecondary_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.secondary = colSecondary.selectedColorHex;
        }

        private void colSecondaryAlt_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.secondaryAlt = colSecondaryAlt.selectedColorHex;
        }

        private void colSecondaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.secondaryAlt2 = colSecondaryAlt2.selectedColorHex;
        }

        private void colSecondaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.secondaryAlt3 = colSecondaryAlt3.selectedColorHex;
        }

        private void colSecondaryContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.secondaryContrast = colSecondaryContrast.selectedColorHex;
        }

        private void colNeutralPrimary_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralPrimary = colNeutralPrimary.selectedColorHex;
        }

        private void colNeutralPrimaryAlt_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralPrimaryAlt = colNeutralPrimaryAlt.selectedColorHex;
        }

        private void colNeutralPrimaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralPrimaryAlt2 = colNeutralPrimaryAlt2.selectedColorHex;
        }

        private void colNeutralPrimaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralPrimaryAlt3 = colNeutralPrimaryAlt3.selectedColorHex;
        }

        private void colNeutralPrimaryContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralPrimaryContrast = colNeutralPrimaryContrast.selectedColorHex;
        }

        private void colNeutralSecondary_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralSecondary = colNeutralSecondary.selectedColorHex;
        }

        private void colNeutralSecondaryAlt_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralSecondaryAlt = colNeutralSecondaryAlt.selectedColorHex;
        }

        private void colNeutralSecondaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralSecondaryAlt2 = colNeutralSecondaryAlt2.selectedColorHex;
        }

        private void colNeutralSecondaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralSecondaryAlt3 = colNeutralSecondaryAlt3.selectedColorHex;
        }

        private void colNeutralSecondaryContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralSecondaryContrast = colNeutralSecondaryContrast.selectedColorHex;
        }

        private void colNeutralTertiary_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralTertiary = colNeutralTertiary.selectedColorHex;
        }

        private void colNeutralTertiaryAlt_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralTertiaryAlt = colNeutralTertiaryAlt.selectedColorHex;
        }

        private void colNeutralTertiaryAlt2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralTertiaryAlt2 = colNeutralTertiaryAlt2.selectedColorHex;
        }

        private void colNeutralTertiaryAlt3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralTertiaryAlt3 = colNeutralTertiaryAlt3.selectedColorHex;
        }

        private void colNeutralTertiaryContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.neutralTertiaryContrast = colNeutralTertiaryContrast.selectedColorHex;
        }

        private void colDanger_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.danger = colDanger.selectedColorHex;
        }

        private void colSuccess_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.success = colSuccess.selectedColorHex;
        }

        private void colWarning_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.warning = colWarning.selectedColorHex;
        }

        private void colInfo_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.info = colInfo.selectedColorHex;
        }

        private void colDangerContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.dangerContrast = colDangerContrast.selectedColorHex;
        }

        private void colSuccessContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.successContrast = colSuccessContrast.selectedColorHex;
        }

        private void colWarningContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.warningContrast = colWarningContrast.selectedColorHex;
        }

        private void colInfoContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.infoContrast = colInfoContrast.selectedColorHex;
        }

        private void colKPIGood_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiGood = colKPIGood.selectedColorHex;
        }

        private void colKPIBad_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiBad = colKPIBad.selectedColorHex;
        }

        private void colKPINeutral_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiNeutral = colKPINeutral.selectedColorHex;
        }

        private void colKPINone_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiNone = colKPINone.selectedColorHex;
        }

        private void colKPIGoodContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiGoodContrast = colKPIGoodContrast.selectedColorHex;
        }

        private void colKPIBadContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiBadContrast = colKPIBadContrast.selectedColorHex;
        }

        private void colKPINeutralContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiNeutralContrast = colKPINeutralContrast.selectedColorHex;
        }

        private void colKPINoneContrast_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.rsinterface.kpiNoneContrast = colKPINoneContrast.selectedColorHex;
        }

        private void colDatapoints1_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[0] = colDatapoints1.selectedColorHex;
        }

        private void colDatapoints2_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[1] = colDatapoints2.selectedColorHex;
        }

        private void colDatapoints3_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[2] = colDatapoints3.selectedColorHex;
        }

        private void colDatapoints4_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[3] = colDatapoints4.selectedColorHex;
        }

        private void colDatapoints5_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[4] = colDatapoints5.selectedColorHex;
        }

        private void colDatapoints6_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[5] = colDatapoints6.selectedColorHex;
        }

        private void colDatapoints7_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[6] = colDatapoints7.selectedColorHex;
        }

        private void colDatapoints8_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[7] = colDatapoints8.selectedColorHex;
        }

        private void colDatapoints9_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[8] = colDatapoints9.selectedColorHex;
        }

        private void colDatapoints10_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[9] = colDatapoints10.selectedColorHex;
        }

        private void colDatapoints11_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[10] = colDatapoints11.selectedColorHex;
        }

        private void colDatapoints12_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.dataPoints[11] = colDatapoints12.selectedColorHex;
        }

        private void colGood_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.good = colGood.selectedColorHex;
        }

        private void colBad_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.bad = colBad.selectedColorHex;
        }

        private void colNeutral_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.neutral = colNeutral.selectedColorHex;
        }

        private void colNone_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.none = colNone.selectedColorHex;
        }

        private void colBackground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.background = colBackground.selectedColorHex;
        }

        private void colForeground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.foreground = colForeground.selectedColorHex;
        }

        private void colMapBase_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altMapBase = colMapBase.selectedColorHex;
        }

        private void colPanelBackground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.panelBackground = colPanelBackground.selectedColorHex;
        }

        private void colPanelAccent_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.panelAccent = colPanelAccent.selectedColorHex;
        }

        private void colTableAccent_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.tableAccent = colTableAccent.selectedColorHex;
        }

        private void colAltBackground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altBackground = colAltBackground.selectedColorHex;
        }

        private void colAltForeground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altForeground = colAltForeground.selectedColorHex;
        }

        private void colAltMapBase_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altMapBase = colAltMapBase.selectedColorHex;
        }

        private void colAltPanelBackground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altPanelBackground = colAltPanelBackground.selectedColorHex;
        }

        private void colAltPanelForeground_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altPanelForeground = colAltPanelForeground.selectedColorHex;
        }

        private void colAltPanelAccent_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altTableAccent = colAltPanelAccent.selectedColorHex;
        }

        private void colAltTableAccent_ColorChanged(object sender, EventArgs e)
        {
            _rsbrand.theme.altTableAccent = colAltTableAccent.selectedColorHex;
        }

        #endregion

        /// <summary>
        /// Returns the temp Path of the current user
        /// </summary>
        public string TempPath
        {
            get { return Path.GetTempPath(); }
        }

        /// <summary>
        /// This procedure updates all Color Controls to the current colors that are set in the JSon Object. It is
        /// used to reset the program and to set the controls colors after the branding file has been loaded. 
        /// </summary>
        private void InitColorPickers()
        {
            colPrimary.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primary);
            colPrimaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primaryAlt);
            colPrimaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primaryAlt2);
            colPrimaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primaryAlt3);
            colPrimaryAlt4.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primaryAlt4);
            colPrimaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.primaryContrast);

            colSecondary.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.secondary);
            colSecondaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.secondaryAlt);
            colSecondaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.secondaryAlt2);
            colSecondaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.secondaryAlt3);
            colSecondaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.secondaryContrast);

            colNeutralPrimary.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralPrimary);
            colNeutralPrimaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralPrimaryAlt);
            colNeutralPrimaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralPrimaryAlt2);
            colNeutralPrimaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralPrimaryAlt3);
            colNeutralPrimaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralPrimaryContrast);

            colNeutralSecondary.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralSecondary);
            colNeutralSecondaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralSecondaryAlt);
            colNeutralSecondaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralSecondaryAlt2);
            colNeutralSecondaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralSecondaryAlt3);
            colNeutralSecondaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralSecondaryContrast);

            colNeutralTertiary.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralTertiary);
            colNeutralTertiaryAlt.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralTertiaryAlt);
            colNeutralTertiaryAlt2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralTertiaryAlt2);
            colNeutralTertiaryAlt3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralTertiaryAlt3);
            colNeutralTertiaryContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.neutralTertiaryContrast);

            colDanger.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.danger);
            colDangerContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.dangerContrast);
            colSuccess.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.success);
            colSuccessContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.successContrast);
            colInfo.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.info);
            colInfoContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.infoContrast);
            colWarning.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.warning);
            colWarningContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.warningContrast);

            colKPIGood.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiGood);
            colKPIGoodContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiGoodContrast);
            colKPIBad.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiBad);
            colKPIBadContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiBadContrast);
            colKPINeutral.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiNeutral);
            colKPINeutralContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiNeutralContrast);
            colKPINone.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiNone);
            colKPINoneContrast.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.rsinterface.kpiNoneContrast);

            colDatapoints1.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[0]);
            colDatapoints2.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[1]);
            colDatapoints3.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[2]);
            colDatapoints4.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[3]);
            colDatapoints5.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[4]);
            colDatapoints6.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[5]);
            colDatapoints7.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[6]);
            colDatapoints8.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[7]);
            colDatapoints9.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[8]);
            colDatapoints10.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[9]);
            colDatapoints11.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[10]);
            colDatapoints12.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.dataPoints[11]);

            colBad.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.bad);
            colGood.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.good);
            colNeutral.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.neutral);
            colNone.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.none);

            colBackground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.background);
            colForeground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.foreground);
            colMapBase.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.mapBase);
            colPanelBackground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.panelBackground);
            colPanelForeground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.panelForeground);
            colPanelAccent.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.panelAccent);
            colTableAccent.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.tableAccent);

            colAltBackground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altBackground);
            colAltForeground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altForeground);
            colAltMapBase.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altMapBase);
            colAltPanelBackground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altPanelBackground);
            colAltPanelForeground.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altPanelForeground);
            colAltPanelAccent.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altPanelAccent);
            colAltTableAccent.selectedColor = (Color)ColorTranslator.FromHtml(_rsbrand.theme.altTableAccent);

            txtBrandPackageName.Text = _rsbrand.name;
            txtBrandPackageVersion.Text = _rsbrand.version;
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
                _rsbrand.Reset();
                txtColorsFilename.Text = "colors.json";
                txtLogoFilename.Text = "logo.png";
                InitColorPickers();
            }
        }

        /// <summary>
        /// This procedure is called when the button "Create Branding" is pressed. It will check if a logo is available.
        /// If not it will show an error message and quit the save process. If a logo is available then the procedure
        /// will create a folder in the Temp-Folder of the current user and will drop the files metadata.xml, colors.json 
        /// and logo.png in this folder. After that it will zip the folder and copy it to the desired output directory.
        /// Last step is deleting the temporary folder and files. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString();
            saveFileDialog1.FileName = txtBrandPackageName.Text;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (pbLogo.Image != null)
                {
                    string zipName = String.Format(@"{0}\{1}.zip", TempPath, txtBrandPackageName.Text);
                    string directoryName = String.Format(@"{0}\{1}", TempPath, txtBrandPackageName.Text);
                    string jsonName = String.Format(@"{0}\{1}\{2}", TempPath, txtBrandPackageName.Text, txtColorsFilename.Text);
                    string metadataName = String.Format(@"{0}\{1}\metadata.xml", TempPath, txtBrandPackageName.Text);
                    string imageName = String.Format(@"{0}\{1}\{2}", TempPath, txtBrandPackageName.Text, txtLogoFilename.Text);
                    string destFileName = saveFileDialog1.FileName;

                    string json = JsonConvert.SerializeObject(_rsbrand).Replace("rsinterface", "interface");
                    string metadata = _rsmeta.GenerateMetadataFile(txtBrandPackageName.Text, txtLogoFilename.Text, txtColorsFilename.Text);

                    CreateTempDirectory(directoryName);

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

        private void button4_Click(object sender, EventArgs e)
        {
            if (openLogoDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openLogoDialog.FileName;
                pbLogo.Image = Image.FromFile(openLogoDialog.FileName);
                if (pbLogo.Image.Width != 80 || pbLogo.Image.Height != 80)
                {
                    MessageBox.Show("The size of the image does not fit. Best size is 80x80 px.\nResizing the image.", "Resizing", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    pbLogo.Image = new Bitmap(pbLogo.Image, new Size(80, 80));
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
            _rsbrand.theme.panelForeground = colAltPanelForeground.selectedColorHex;
        }


        /// <summary>
        /// This procedure tries to create the temporary directory where the files are stored before they are either zipped or loaded into
        /// SSRS Styler
        /// </summary>
        /// <param name="directoryName">Name of the temporary directory</param>
        private void CreateTempDirectory(string directoryName)
        {
            if (!Directory.Exists(directoryName))
            {
                Directory.CreateDirectory(directoryName);
            }
            else
            {
                Directory.Delete(directoryName, true);
                Directory.CreateDirectory(directoryName);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openBrandingDialog.ShowDialog() == DialogResult.OK)
            {
                string zipName = openBrandingDialog.FileName;
                string directoryName = String.Format(@"{0}\{1}", TempPath, Path.GetFileNameWithoutExtension(zipName));
                string metadataName = String.Format(@"{0}\metadata.xml", directoryName);
                XmlDataDocument xmldoc = new XmlDataDocument();
                XmlNodeList xmlnode;
                string jsonName;
                string imageName;
                string destFileName = saveFileDialog1.FileName;
                string jsonfile;

                CreateTempDirectory(directoryName);

                ZipFile.ExtractToDirectory(zipName, directoryName);

                if (!File.Exists(metadataName))
                {
                    MessageBox.Show("No metadata.xml file found!\nAborting opening process.", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                FileStream fs = new FileStream(metadataName, FileMode.Open, FileAccess.Read);
                xmldoc.Load(fs);

                xmlnode = xmldoc.GetElementsByTagName("SystemResourcePackage");

                if (!(xmlnode == null || xmlnode[0].Attributes["type"].Value != "UniversalBrand"))
                {
                    txtBrandPackageName.Text = xmlnode[0].Attributes["name"].Value;
                    txtBrandPackageVersion.Text = xmlnode[0].Attributes["version"].Value;

                    //FGE: Get colors-filename
                    xmlnode = xmldoc.SelectNodes("/*[local-name()='SystemResourcePackage'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][1]/*[local-name()='Contents'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][1]/*[local-name()='Item'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][1]");
                    if (xmlnode != null)
                    {
                        jsonName = String.Format(@"{0}\{1}", directoryName, xmlnode[0].Attributes["path"].Value);
                        using (var streamReader = new StreamReader(jsonName, Encoding.UTF8))
                        {
                            jsonfile = streamReader.ReadToEnd().Replace("interface", "rsinterface"); ;
                            _rsbrand = JsonConvert.DeserializeObject<RSBranding>(jsonfile);
                            InitColorPickers();
                            streamReader.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("XML File has wrong format!\nCannot find node 'colors'", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //FGE: Get image-filename
                    xmlnode = xmldoc.SelectNodes("/*[local-name()='SystemResourcePackage'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][1]/*[local-name()='Contents'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][1]/*[local-name()='Item'][namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/systemresourcepackagemetadata'][2]");
                    if (xmlnode != null && xmlnode.Count == 1)
                    {
                        imageName = String.Format(@"{0}\{1}", directoryName, xmlnode[0].Attributes["path"].Value);
                        using (FileStream stream = new FileStream(imageName, FileMode.Open, FileAccess.Read))
                        {
                            pbLogo.Image = Image.FromStream(stream);
                        }
                    }
                    else if (xmlnode == null)
                    {
                        MessageBox.Show("XML File has wrong format!\nCannot find node 'colors'", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("XML File has wrong format!\nEither no nodes were returned or type was not 'UniversalBrand'", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                fs.Close();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo("https://blogs.msdn.microsoft.com/sqlrsteamblog/2016/03/20/how-to-create-a-custom-brand-package-for-reporting-services-with-sql-server-2016/");
            Process.Start(sInfo);
        }
    }
}
