using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using System.Windows.Forms;
using SuperPUWEtty2.Gui;
using SuperPUWEtty2.Data;
using System.Drawing;

namespace SuperPUWEtty2UnitTests
{
    //[TestFixture]
    public class QuickSelectorTests
    {

        [TestView]
        public void Test()
        {
            List<SessionData> sessions = SessionData.LoadSessionsFromFile("c:/Users/beau/SuperPUWEtty2/sessions.xml");
            QuickSelectorData data = new QuickSelectorData();

            foreach (SessionData sd in sessions)
            {
                data.ItemData.AddItemDataRow(
                    sd.SessionName, 
                    sd.SessionId, 
                    sd.Proto == ConnectionProtocol.Cygterm || sd.Proto == ConnectionProtocol.Mintty ? Color.Blue : Color.Black, null);
            }

            QuickSelectorOptions opt = new QuickSelectorOptions();
            opt.Sort = data.ItemData.DetailColumn.ColumnName;
            opt.BaseText = "Open Session";

            QuickSelector d = new QuickSelector();
            d.ShowDialog(null, data, opt);
        }


    }
}
