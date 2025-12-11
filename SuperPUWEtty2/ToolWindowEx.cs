using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace SuperPUWEtty2
{
    /// <summary>
    /// ToolWindow that supports an MRU tab switching
    /// </summary>
    public partial class ToolWindowDocument : ToolWindow
    {
        public ToolWindowDocument()
        {
            InitializeComponent();
            if (SuperPUWEtty2.MainForm == null) return;
            
            // Insert this panel into the list used for Ctrl-Tab handling.
            if (SuperPUWEtty2.MainForm.CurrentPanel == null)
            {
                // First panel to be created
                SuperPUWEtty2.MainForm.CurrentPanel = this.PreviousPanel = this.NextPanel = this;
            }
            else
            {
                // Other panels exist. Tie ourselves into list ahead of current panel.
                this.PreviousPanel = SuperPUWEtty2.MainForm.CurrentPanel;
                this.NextPanel = SuperPUWEtty2.MainForm.CurrentPanel.NextPanel;
                SuperPUWEtty2.MainForm.CurrentPanel.NextPanel = this;
                this.NextPanel.PreviousPanel = this;

                // We are now the current panel
                SuperPUWEtty2.MainForm.CurrentPanel = this;
            }
        }

        // Make this panel the current one. Remove from previous
        // position in list and re-add in front of current panel
        public void MakePanelCurrent()
        {
            if (SuperPUWEtty2.MainForm.CurrentPanel == this)
                return;

            // Remove ourselves from our position in chain
            this.PreviousPanel.NextPanel = this.NextPanel;
            this.NextPanel.PreviousPanel = this.PreviousPanel;

            this.PreviousPanel = SuperPUWEtty2.MainForm.CurrentPanel;
            this.NextPanel = SuperPUWEtty2.MainForm.CurrentPanel.NextPanel;
            SuperPUWEtty2.MainForm.CurrentPanel.NextPanel = this;
            this.NextPanel.PreviousPanel = this;

            SuperPUWEtty2.MainForm.CurrentPanel = this;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            if (SuperPUWEtty2.MainForm == null) return;

            // only 1 panel
            if (SuperPUWEtty2.MainForm.CurrentPanel == this && this.NextPanel == this && this.PreviousPanel == this)
            {
                SuperPUWEtty2.MainForm.CurrentPanel = null;
                return;
            }

            // Remove ourselves from our position in chain and set last active tab as current
            if (this.PreviousPanel != null)
            {
                this.PreviousPanel.NextPanel = this.NextPanel;
            }
            if (this.NextPanel != null)
            {
                this.NextPanel.PreviousPanel = this.PreviousPanel;
            }
            SuperPUWEtty2.MainForm.CurrentPanel = this.PreviousPanel;

            // manipulate tabs
            if (this.DockHandler.Pane != null)
            {
                int idx = this.DockHandler.Pane.Contents.IndexOf(this);
                if (idx > 0)
                {
                    IDockContent contentToActivate = this.DockHandler.Pane.Contents[idx - 1];
                    contentToActivate.DockHandler.Activate();
                }
            }
        }


        public ToolWindowDocument PreviousPanel { get; set; }
        public ToolWindowDocument NextPanel { get; set; }
    }
}
