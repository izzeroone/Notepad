using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using ScintillaNET;
using System.Windows.Forms;

/// <summary>
/// Tab page Status
/// </summary>
namespace NotePad__
{
    partial class TabControlClass
    {
        //A list to store all the tabpages status
        static List<TabPageStatus> listOfTabPageStatus = new List<TabPageStatus>();

        /// <summary>
        /// Get selected tab page status
        /// </summary>
        public static TabPageStatus CurrentTabPageStatus
        {
            get
            {
                foreach (TabPageStatus tabPageStatus in listOfTabPageStatus)
                {
                    if (tabPageStatus.TabPage == tabControl.SelectedTab)
                    {
                        return tabPageStatus;
                    }
                }
                return null;
            }
        }

        //a class to store tab status
        public class TabPageStatus
        {
            private TabPage tabPage;
            private string language;
            private bool documentMapEnabled;
            private bool canUndo;
            private bool canRedo;

            public TabPage TabPage
            {
                get
                {
                    return tabPage;
                }

                set
                {
                    tabPage = value;
                }
            }

            public string Language
            {
                get
                {
                    return language;
                }

                set
                {
                    language = value;
                }
            }

            public bool CanUndo
            {
                get
                {
                    return canUndo;
                }

                set
                {
                    canUndo = value;
                }
            }

            public bool CanRedo
            {
                get
                {
                    return canRedo;
                }

                set
                {
                    canRedo = value;
                }
            }

            public bool DocumentMapEnabled { get { return documentMapEnabled; } set { documentMapEnabled = value; } }
        };

        /// <summary>
        /// Init tab page status
        /// </summary>
        /// <param name="tabPage"></param>
        private static void InitTabPageStatus(TabPage tabPage)
        {
            TabPageStatus tabPageStatus = new TabPageStatus();
            tabPageStatus.TabPage = tabPage;
            tabPageStatus.Language = "NormalText";
            tabPageStatus.CanUndo = false;
            tabPageStatus.CanRedo = false;
            tabPageStatus.DocumentMapEnabled = false;
            listOfTabPageStatus.Add(tabPageStatus);
        }

        /// <summary>
        /// Delete status of current tabpage
        /// </summary>
        /// <param name="tabPage"></param>
        public static void RemoveTabPageStatus(TabPage tabPage)
        {
            TabPageStatus tabStatusToDelete = null;
            foreach (TabPageStatus tabPageStatus in listOfTabPageStatus)
            {
                if (tabPageStatus.TabPage == tabPage)
                {
                    tabStatusToDelete = tabPageStatus;
                }
            }

            if (tabStatusToDelete != null)
            {
                listOfTabPageStatus.Remove(tabStatusToDelete);
            }
        }

    }
}
