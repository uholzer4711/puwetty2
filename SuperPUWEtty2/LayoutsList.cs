using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using SuperPUWEtty2.Data;
using SuperPUWEtty2.Gui;

namespace SuperPUWEtty2
{
    public partial class LayoutsList : ToolWindow
    {
        public LayoutsList()
        {
            InitializeComponent();

            this.listBoxLayouts.DataSource = SuperPUWEtty2.Layouts;
        }

        protected override void OnClosed(EventArgs e)
        {
            this.listBoxLayouts.DataSource = null;
            base.OnClosed(e);
        }

        private void contextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            int idx = IndexAtCursor();
            e.Cancel = idx == -1;

            if (e.Cancel)
                return;
            
            LayoutData layout = (LayoutData) this.listBoxLayouts.Items[idx];

            loadInNewInstanceToolStripMenuItem.Enabled = !SuperPUWEtty2.Settings.SingleInstanceMode;
            renameToolStripMenuItem.Enabled = !layout.IsReadOnly;
            deleteToolStripMenuItem.Enabled = !layout.IsReadOnly;
        }

        private void listBoxLayouts_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // select item under mouse
                int idx = this.listBoxLayouts.IndexFromPoint(e.X, e.Y);
                if (idx != -1)
                {
                    this.listBoxLayouts.SelectedIndex = idx;
                }
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutData layout = (LayoutData) this.listBoxLayouts.SelectedItem;
            if (layout != null)
            {
                SuperPUWEtty2.LoadLayout(layout);
            }
        }

        private void loadInNewInstanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutData layout = (LayoutData)this.listBoxLayouts.SelectedItem;
            if (layout != null)
            {
                SuperPUWEtty2.LoadLayoutInNewInstance(layout);
            }
        }

        private void setAsDefaultLayoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutData layout = (LayoutData)this.listBoxLayouts.SelectedItem;
            if (layout != null)
            {
                SuperPUWEtty2.SetLayoutAsDefault(layout.Name);
            }
        }

        private void listBoxLayouts_DoubleClick(object sender, EventArgs e)
        {
            int idx = IndexAtCursor();
            if (idx != -1)
            {
                LayoutData layout = (LayoutData)this.listBoxLayouts.Items[idx];
                if (layout != null)
                {
                    SuperPUWEtty2.LoadLayout(layout);
                }
            }
            
        }

        int IndexAtCursor()
        {
            Point p = this.listBoxLayouts.PointToClient(Cursor.Position);
            return this.listBoxLayouts.IndexFromPoint(p.X, p.Y);
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutData layout = (LayoutData)this.listBoxLayouts.SelectedItem;
            if (layout != null)
            {
                if (DialogResult.Yes == MessageBox.Show(this, "Delete Layout (" + layout.Name + ")?", "Delete Layout", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    SuperPUWEtty2.RemoveLayout(layout.Name, true);
                }
            }
        }

        private void renameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutData layout = (LayoutData)this.listBoxLayouts.SelectedItem;
            if (layout != null)
            {
                dlgRenameItem renameDialog = new dlgRenameItem
                {
                    DetailName = String.Empty,
                    ItemName = layout.Name,
                    ItemNameValidator = this.ValidateLayoutName
                };
                if (DialogResult.OK == renameDialog.ShowDialog(this))
                {
                    SuperPUWEtty2.RenameLayout(layout, renameDialog.ItemName);
                }
            }
            
        }

        bool ValidateLayoutName(string name, out string error)
        {
            LayoutData layout = SuperPUWEtty2.FindLayout(name);
            if (layout != null)
            {
                error = "Layout exists with same name";
                return false;
            }

            error = null;
            return true;
        }
    }
}
