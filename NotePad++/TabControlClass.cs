using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MyTextBox;
using System.Collections;

/// <summary>
/// just a bunch of things to setup tab control look like a notepad
/// </summary>
namespace NotePad__
{
    partial class TabControlClass
    {

        //the "x" image 
        private static Image CloseImage;

        //a variable to easily hold the reference tab control
        private static TabControl tabControl;

        //the "x" highlighted image 
        //private static Image HighLightCloseImage;

        #region Properties

        public static TextArea CurrentTextArea
        {
            get
            {
                if (tabControl.SelectedTab != null)

                    return (tabControl.SelectedTab.Controls[0] as MyRichTextBox).TextArea;
                return null;
            }
        }

        #endregion

        public static void SetupTabControl(TabControl _tabControl)
        {
            //store _tabControl into tabControl to easily do some works 
            tabControl = _tabControl;

            //allow drop
            tabControl.AllowDrop = true;

            //set the draw mode to ownerdrawfixed to be able to manually draw it 
            tabControl.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;

            //set the scale of each tab page
            tabControl.Padding = new Point(10, 3);

            //get the "x" image 
            CloseImage = Properties.Resources.Close;


            //draw the "x" button
            tabControl.DrawItem += tabControl_DrawItem;

            //Drag on Drop tab
            tabControl.DragOver += tabControl_DragOver;

            //Click "x" image
            tabControl.MouseDown += tabControl_MouseDown;

            //allow do drag and drop when mouse move around
            tabControl.MouseMove += tabControl_MouseMove;

            #region hightligh "x" button
            //get the "x" highlighted image 
            //HighLightCloseImage = Properties.Resources.Delete_50;

            ////two line of codes below work right for highlighting "x" image when mouse enter
            ////but cause so many flickings 
            ///so I just leave it here 
            ////paint "x" image when mouse enter
            //tabControl.MouseEnter += tabControl_MouseEnter;

            ////paint "x" image when mouse leave
            //tabControl.MouseLeave += tabControl_MouseLeave;
            #endregion

        }

        #region All tab control Events

        /// <summary>
        /// when user click "x" image
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void tabControl_MouseDown(object sender, MouseEventArgs e)
        {

            //this for loop check if user click mouse at the position of an "x" image
            //loop through all tab page to demonstrate which tab tage is being clicked
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                //get the shape of the tab
                Rectangle tabRect = tabControl.GetTabRect(tabControl.TabPages.IndexOf(tabPage));
                //inflate
                tabRect.Inflate(-2, -2);
                //get the shape of the "x" image
                Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
                                         tabRect.Top,
                                         CloseImage.Width,
                                         CloseImage.Height);
                //if mouse click in the shape of "x" image 
                if (imageRect.Contains(e.Location))
                {
                    if ((DiaLogClass.ShowSafeCloseTabDialog(tabPage) == "Cancel"))
                    {
                        break;
                    }

                    //remove this tab
                    if (tabControl.TabPages.Count > 1) //if this is the last tab page remaining, don't close it
                    {
                        //we use selectedTab variable here because if we click mouse, tabControl will select the clicked tab 
                        //and this if below will always false  
                        //if the tab we are deleting is not the selected tab
                        if (tabPage != (TabPage)tabControl.Tag)
                        {
                            //Important Note: we can change the order of two lines of code below and the main result of us will be the same 
                            //but if we do that, it will cause flinking 
                            tabControl.SelectedTab = (TabPage)tabControl.Tag;

                            //remove tabpage status
                            RemoveTabPageStatus(tabPage);

                            //remove tab page
                            tabControl.TabPages.Remove(tabPage);
                        }
                        else //if the tab we are deleting is the selected tab
                        {
                            //if this isn't the first tab
                            if (tabControl.SelectedIndex != 0)
                            {
                                //set the selectedtab to the closest tab page
                                tabControl.SelectedIndex = tabControl.SelectedIndex - 1;
                            }
                            else //if the tab we are about to remove is the first tab, just simply set selectedtab to 1
                            {
                                tabControl.SelectedIndex = 1;
                            }

                            //remove tabpage status
                            RemoveTabPageStatus((TabPage)tabControl.Tag);

                            //remove tab page
                            tabControl.TabPages.Remove((TabPage)tabControl.Tag);
                        }
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// allow DoDragDrop on Mouse Move
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void tabControl_MouseMove(object sender, MouseEventArgs e)
        {
            // this is necessary for the mousedown event
            tabControl.Tag = tabControl.SelectedTab;

            //if user click left mouse 
            if ((e.Button != MouseButtons.Left)) return;

            // start drag and drop
            tabControl.DoDragDrop(tabControl.SelectedTab, DragDropEffects.All);

        }

        /// <summary>
        /// switch tag when drag around
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void tabControl_DragOver(object sender, DragEventArgs e)
        {

            // if there is no tab being draged
            if (e.Data.GetData(typeof(TabPage)) == null) return;

            //get the tab page being drag
            TabPage dragTab = (TabPage)e.Data.GetData(typeof(TabPage));

            //get the index of the tab being draged
            int dragTab_index = tabControl.TabPages.IndexOf(dragTab);

            // get the index of tab page having the mouse on it 
            int hoverTab_index = getHoverTabIndex();

            //if not, return
            if (hoverTab_index < 0) { e.Effect = DragDropEffects.None; return; }

            //get the tab having the mouse on it
            TabPage hoverTab = tabControl.TabPages[hoverTab_index];

            //set the effect when mouse move around (just to make it look like we are draging something)
            //this is not very important
            e.Effect = DragDropEffects.Move;

            // if drag tab and hover tab are the same, just return
            if (dragTab == hoverTab) return;

            // swap dragTab & hoverTab - avoids toggeling
            Rectangle dragTabRect = tabControl.GetTabRect(dragTab_index);
            Rectangle hoverTabRect = tabControl.GetTabRect(hoverTab_index);

            //some calculations
            if (dragTabRect.Width < hoverTabRect.Width)
            {
                Point tcLocation = tabControl.PointToScreen(tabControl.Location);

                if (dragTab_index < hoverTab_index)
                {
                    if ((e.X - tcLocation.X) > ((hoverTabRect.X + hoverTabRect.Width) - dragTabRect.Width))
                        swapTabPages(dragTab, hoverTab);
                }
                else
                {
                    if ((e.X - tcLocation.X) < (hoverTabRect.X + dragTabRect.Width))
                        swapTabPages(dragTab, hoverTab);
                }
            }
            else swapTabPages(dragTab, hoverTab);

            //set the selected tab by draged tab 
            //this is very important
            tabControl.SelectedIndex = tabControl.TabPages.IndexOf(dragTab);
        }

        /// <summary>
        /// Draw a close image in the right side of a tab 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void tabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

            if (e.Index == tabControl.SelectedIndex)
            {
                //this try catch just draw selected tab
                try
                {
                    //Get the rectangle of the selected tab
                    Rectangle tabRect = tabControl.GetTabRect(tabControl.SelectedIndex);

                    //Set background Color of this tab
                    e.Graphics.FillRectangle(Brushes.White, tabRect);

                    //Inflate it (actually this code deflate this rectangle by -2 both x and y )
                    tabRect.Inflate(-2, -2);

                    //"x" image
                    Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
                                                        tabRect.Top,
                                                        CloseImage.Width,
                                                        CloseImage.Height);

                    //draw the string text 
                    e.Graphics.DrawString(tabControl.TabPages[e.Index].Text,
                                          tabControl.Font, Brushes.Black, tabRect);
                    //draw the "x" image
                    e.Graphics.DrawImage(CloseImage, imageRect.Location);

                }
                catch (Exception) { }
            }
            else
            {
                //this try catch will draw all the tab except the selected tab page
                try
                {
                    //Get the rectangle of the tab being draw
                    Rectangle tabRect = tabControl.GetTabRect(e.Index);

                    //Inflate it (actually this code deflate this rectangle by -2 both x and y )
                    tabRect.Inflate(-2, -2);

                    //"x" image
                    Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
                                                        tabRect.Top,
                                                        CloseImage.Width,
                                                        CloseImage.Height);


                    //draw the string text 
                    e.Graphics.DrawString(tabControl.TabPages[e.Index].Text,
                                          tabControl.Font, Brushes.Black, tabRect);
                    //draw the "x" image
                    e.Graphics.DrawImage(CloseImage, imageRect.Location);

                }
                catch (Exception) { }
            }
        }

        #endregion

        #region Really essential Functions

        /// <summary>
        /// Create new tab page 
        /// </summary>
        /// <param name="tabName">name of the tab</param>
        public static TabPage CreateNewTabPage(string tabName)
        {

            //Create a new Tabpage
            TabPage newTabPage = new TabPage(tabName);
            
            //Init tab page status
            InitTabPageStatus(newTabPage);

            ////Declare a scintilla textBox
            MyRichTextBox textBox = new MyRichTextBox();

            ////Init essential things to make scintilla textbox like a notepad 
            ////these are the most important things in this project
            MyTextBoxClass.InitAllStuffs(textBox, tabControl);

            //Add scintilla TextBox to tabPage
            newTabPage.Controls.Add(textBox);

            //Fill the form by textArea
            textBox.Dock = DockStyle.Fill;

            //Add tabPage to tabControl
            tabControl.TabPages.Add(newTabPage);

            //switch to the new tab
            tabControl.SelectedTab = newTabPage;

            return newTabPage;
        }

        #endregion

        #region Some other function

        /// <summary>
        /// get the index of tab having the mouse on it
        /// </summary>
        /// <param name="tc"></param>
        /// <returns></returns> 
        private static int getHoverTabIndex()
        {
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                if (tabControl.GetTabRect(i).Contains(tabControl.PointToClient(Cursor.Position)))
                    return i;
            }

            return -1;
        }

        /// <summary>
        /// swap tab pages
        /// </summary>
        /// <param name="tc"></param>
        /// <param name="src"></param>
        /// <param name="dst"></param>
        private static void swapTabPages(TabPage src, TabPage dst)
        {
            //swap
            int index_src = tabControl.TabPages.IndexOf(src);
            int index_dst = tabControl.TabPages.IndexOf(dst);
            tabControl.TabPages[index_dst] = src;
            tabControl.TabPages[index_src] = dst;

        }

        #endregion

        #region these are the function essential for highlight "x" image when mouse move on 

        //private static void tabControl_DrawItemWhenEnter(object sender, System.Windows.Forms.DrawItemEventArgs e)
        //{
        //    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

        //    if (e.Index == getHoverTabIndex())
        //    {

        //        //this try catch just draw selected tab
        //        try
        //        {
        //            //Get the rectangle of the selected tab
        //            Rectangle tabRect = tabControl.GetTabRect(e.Index);

        //            //Set background Color of this tab
        //            e.Graphics.FillRectangle(Brushes.White, tabRect);

        //            //Inflate it (actually this code deflate this rectangle by -2 both x and y )
        //            tabRect.Inflate(-2, -2);

        //            //"x" image
        //            Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
        //                                                tabRect.Top,
        //                                                CloseImage.Width,
        //                                                CloseImage.Height);

        //            //draw the string text 
        //            e.Graphics.DrawString(tabControl.TabPages[e.Index].Text,
        //                                  tabControl.Font, Brushes.Black, tabRect);
        //            //draw the "x" image
        //            e.Graphics.DrawImage(HighLightCloseImage, imageRect.Location);

        //        }
        //        catch (Exception) { }

        //    }
        //    else
        //    {
        //        if (e.Index == tabControl.SelectedIndex)
        //        {
        //            //this try catch just draw selected tab
        //            try
        //            {
        //                //Get the rectangle of the selected tab
        //                Rectangle tabRect = tabControl.GetTabRect(tabControl.SelectedIndex);

        //                //Set background Color of this tab
        //                e.Graphics.FillRectangle(Brushes.White, tabRect);

        //                //Inflate it (actually this code deflate this rectangle by -2 both x and y )
        //                tabRect.Inflate(-2, -2);

        //                //"x" image
        //                Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
        //                                                    tabRect.Top,
        //                                                    CloseImage.Width,
        //                                                    CloseImage.Height);

        //                //draw the string text 
        //                e.Graphics.DrawString(tabControl.TabPages[e.Index].Text,
        //                                      tabControl.Font, Brushes.Black, tabRect);
        //                //draw the "x" image
        //                e.Graphics.DrawImage(CloseImage, imageRect.Location);

        //            }
        //            catch (Exception) { }
        //        }
        //        else
        //        {
        //            //this try catch will draw all the tab except the selected tab page
        //            try
        //            {
        //                //Get the rectangle of the tab being draw
        //                Rectangle tabRect = tabControl.GetTabRect(e.Index);

        //                //Inflate it (actually this code deflate this rectangle by -2 both x and y )
        //                tabRect.Inflate(-2, -2);

        //                //"x" image
        //                Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
        //                                                    tabRect.Top,
        //                                                    CloseImage.Width,
        //                                                    CloseImage.Height);


        //                //draw the string text 
        //                e.Graphics.DrawString(tabControl.TabPages[e.Index].Text,
        //                                      tabControl.Font, Brushes.Black, tabRect);
        //                //draw the "x" image
        //                e.Graphics.DrawImage(CloseImage, imageRect.Location);

        //            }
        //            catch (Exception) { }
        //        }
        //    }
        //}

        ///// <summary>
        ///// Mouse Enter
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private static void tabControl_MouseEnter(object sender, EventArgs e)
        //{
        //    //this for loop check if user click mouse at the position of an "x" image
        //    //loop through all tab page to demonstrate which tab tage is being clicked
        //    for (int i = 0; i < tabControl.TabPages.Count; i++)
        //    {
        //        //get the shape of the tab
        //        Rectangle tabRect = tabControl.GetTabRect(i);
        //        //inflate
        //        tabRect.Inflate(-2, -2);
        //        //get the shape of the "x" image
        //        Rectangle imageRect = new Rectangle(tabRect.Right - CloseImage.Width,
        //                                 tabRect.Top,
        //                                 CloseImage.Width,
        //                                 CloseImage.Height);
        //        //if mouse enter in the shape of "x" image 
        //        if (imageRect.Contains(tabControl.PointToClient(Cursor.Position)))
        //        {
        //            tabControl.DrawItem -= tabControl_DrawItem;
        //            tabControl.DrawItem += tabControl_DrawItemWhenEnter;
        //            tabControl.Refresh();
        //        }
        //    }
        //}

        ///// <summary>
        ///// Mouse Enter
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private static void tabControl_MouseLeave(object sender, EventArgs e)
        //{
        //    tabControl.DrawItem -= tabControl_DrawItemWhenEnter;
        //    tabControl.DrawItem += tabControl_DrawItem;
        //    tabControl.Refresh();
        //}
        #endregion

    }
}

