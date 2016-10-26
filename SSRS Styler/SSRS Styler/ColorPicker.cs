using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SSRS_Styler
{
    public partial class ColorPicker : UserControl
    {
        private Color _selectedColor;
        private string _oldValue;

        public event EventHandler ColorChanged;

        private void HandleColorChanged(object sender, EventArgs e)
        {
            this.OnColorChanged(EventArgs.Empty);
        }

        /// <summary>
        /// This procedure calculates a contrast color for the given color so that
        /// the text on the background can easily be read. The code was adopted from
        // http://stackoverflow.com/questions/1855884/determine-font-color-based-on-background-color
        /// </summary>
        /// <param name="color">color to which a contrast color should be calculated</param>
        /// <returns>Value of the contrast color.</returns>
        private Color ContrastColor(Color color)
        {
            int d = 0;

            // Counting the perceptive luminance - human eye favors green color... 
            double a = 1 - (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255;

            if (a < 0.5)
                d = 0; // bright colors - black font
            else
                d = 255; // dark colors - white font

            return Color.FromArgb(d, d, d);
        }

        protected virtual void OnColorChanged(EventArgs e)
        {
            this.ColorChanged?.Invoke(this, e);
        }

        [Description("Text that is displayed as a label"), Category("Apperance")]
        public string label
        {
            get { return label1.Text; }
            set { label1.Text = value; }
        }

        [Description("Color that is currently selected."), Category("Apperance")]
        public Color selectedColor
        {
            get { return _selectedColor; }
            set
            {
                _selectedColor = value;
                textBox1.BackColor = value;
                textBox1.ForeColor = ContrastColor(value);
                textBox1.Text = String.Format("#{0:X6}", value.ToArgb() & 0x00FFFFFF);
            }
        }

        public string selectedColorHex
        {
            get { return String.Format("#{0:X6}", selectedColor.ToArgb() & 0x00FFFFFF); ; }
        }

        public ColorPicker()
        {
            InitializeComponent();
            textBox1.BackColorChanged += this.HandleColorChanged;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                selectedColor = colorDialog1.Color;
                HandleColorChanged(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// This event handler only allowes hexadecimal letters and numbers 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') ||(e.KeyChar == (char)8)))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// This event handler copies the old value of the component to 
        /// restore it when the user does not enter any value for the field.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Enter(object sender, EventArgs e)
        {
            _oldValue = textBox1.Text.Remove(0,1);
            textBox1.Text = "";
        }

        /// <summary>
        /// if the user did not enter any value the old value is restored. 
        /// Otherwise the text is calculated into a color value.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0) { textBox1.Text = _oldValue; }
            textBox1.Text = String.Format("#{0:X6}", textBox1.Text);
            selectedColor = (Color)ColorTranslator.FromHtml(textBox1.Text);
        }
    }
}
