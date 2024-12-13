using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ShowScripts
{
    public partial class InputForm : Form
    {
        public string ScreenName;
        public InputForm(string prompt)
        {
            InitializeComponent();
            this.TopMost = true;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        public void ShowInvalidScreenError()
        {
            this.labelInvalidScreen.Visible = true;
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            ScreenName = txtScreenName.Text;
            this.DialogResult = DialogResult.OK;
            this.labelInvalidScreen.Visible = false;
            this.Close();
        }
        private void panelBottomGradient_Paint(object sender, PaintEventArgs e)
        {
            // Define the start and end colors of the gradient
            Color startColor = Color.FromArgb(109, 109, 109); // Top color
            Color endColor = Color.FromArgb(28, 28, 28); // Bottom color

            // Create a rectangle that matches the panel's dimensions
            Rectangle panelRect = panelBottomGradient.ClientRectangle;

            // Create a LinearGradientBrush for the gradient effect
            using (LinearGradientBrush brush = new LinearGradientBrush(panelRect, startColor, endColor, LinearGradientMode.Vertical))
            {
                // Fill the panel with the gradient
                e.Graphics.FillRectangle(brush, panelRect);
            }
        }
        private void TitleBar_Paint(object sender, PaintEventArgs e)
        {
            // Define the start and end colors of the gradient
            Color startColor = Color.FromArgb(180, 180, 180); // Top color
            Color endColor = Color.FromArgb(0, 0, 0); // Bottom color

            // Create a rectangle that matches the panel's dimensions
            Rectangle panelRect = TitleBar.ClientRectangle;

            // Create a LinearGradientBrush for the gradient effect
            using (LinearGradientBrush brush = new LinearGradientBrush(panelRect, startColor, endColor, LinearGradientMode.Vertical))
            {
                // Fill the panel with the gradient
                e.Graphics.FillRectangle(brush, panelRect);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            int visibility = 80;
            if (visibility > 100)
            {
                labelInvalidScreen.Visible = true;
            }
            else
            {
                labelInvalidScreen.Visible = false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ScreenName = "";
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
