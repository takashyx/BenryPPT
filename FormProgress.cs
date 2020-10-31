using System.Windows.Forms;

namespace BenryPPT
{
    public partial class FormProgress : Form
    {
        public FormProgress()
        {
            InitializeComponent();
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Value = 0;
        }

        public void setFormTitle(string titleStr)
        {
            this.Text = titleStr;
        }

        public void setProgressBarMessage(string messageStr)
        {
            this.label_Progress.Text = messageStr;
            this.Update();
        }

        public void setProgressBarPercentage(int percentage)
        {
            this.progressBar.Value = percentage;

        }

        public void exitForm()
        {
            this.Close();
        }
    }
}
