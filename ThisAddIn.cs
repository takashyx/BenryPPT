using Office = Microsoft.Office.Core;

namespace BenryPPT
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane benryPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // create pane
            benryPane = this.CustomTaskPanes.Add(new RibbonControl(), "BenryControl");
        }

        public void ShowPanel()
        {
            benryPane.Visible = true;
            benryPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
