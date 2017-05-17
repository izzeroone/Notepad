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
using System.Collections;
using MyTextBox;

//In this project: 
//automatic syntax highlight: solved (C, C++, C#, VB)
//book mark: solved
//automatic line numbering: solved
//File: save, save as, open, new, close, close all: solved
//Edit: cut, copy, paste,select all, clear: solved
//View: zoom in, zoom out, zoom 100%: solved
//drag and drop file: solved
//drag and drop tab page: solved
//Auto complete box: solved
//
//...
//...
//"x" button on each tab highlighted when mouse move in: haven't solved
//automatic code folding: haven't solved
//WordWrap: disabled to improve performance
//show indent guides: solved
//search text in text area: haven't solved
//show white space: haven't solved
//Brace Matching HighLighting: haven't solved
//Auto indentation: haven't solved
//....
//..... 
namespace NotePad__
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            //Setup Tab Control
            TabControlClass.SetupTabControl(tabControl);

            //just create a new tab page 
            TabControlClass.CreateNewTabPage("New Tab 1");

            SetLanguage("NormalText");

        }


        #region Tool Strip Menu Events

        /// <summary>
        /// Create a new tab by "New" Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Index of the tab page having the name of "New Tab"
            //this is just a variable to check the name of tab page
            int tabIndex = 1;

            //a variable to hold new tab name
            string newTabName = "New Tab " + tabIndex.ToString();

            //loop until we find the "New Tab" tab page has been lacked
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                //get the tabName without "*" (this is considered the real name)
                string tabName = tabControl.TabPages[i].Text;
                tabName = tabName.Replace("*", "");

                if (tabName == newTabName)
                {
                    tabIndex++;
                    newTabName = "New Tab " + tabIndex.ToString();

                    //set this to start over the for loop
                    i = -1;
                }
            }

            //create new tab page
            TabControlClass.CreateNewTabPage(newTabName);

        }

        /// <summary>
        /// copy selected text
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //just a check to ensure no exception occur
            if (TabControlClass.CurrentTextArea != null) TabControlClass.CurrentTextArea.Copy();
        }

        /// <summary>
        /// cut selected text
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea != null) TabControlClass.CurrentTextArea.Cut();
        }

        /// <summary>
        /// paste the text copied or cut at the position of current caret
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea != null) TabControlClass.CurrentTextArea.Paste();
        }

        /// <summary> 
        /// open the choosen file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DiaLogClass.ShowOpenDialog(tabControl);
        }

        /// <summary>
        /// save file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DiaLogClass.ShowSaveDialog(tabControl.SelectedTab);
        }

        /// <summary>
        /// save as file  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DiaLogClass.ShowSaveAsDialog(tabControl.SelectedTab);
        }

        /// <summary>
        /// select all the text in the scintila textBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea != null) TabControlClass.CurrentTextArea.SelectAll();
        }

        /// <summary>
        /// clear 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea != null) TabControlClass.CurrentTextArea.Clear();

        }

        /// <summary>
        /// zoom in
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void zoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabControlClass.CurrentTextArea.ZoomIn();

        }

        /// <summary>
        /// zoom out
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void zoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabControlClass.CurrentTextArea.ZoomOut();

        }

        /// <summary>
        /// Zoom default
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void zoom100ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabControlClass.CurrentTextArea.ZoomFactor = 1;
        }

        /// <summary>
        /// close selected tab
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            if (tabControl.TabPages.Count > 1) //we don't remove tab page when tabControl has only one tab page
            {
                //we can use this one line of code below to easily remove selected tab page but it causes flinking  
                // tabControl.TabPages.Remove(tabControl.SelectedTab);

                //Show SaveDialog
                string result = DiaLogClass.ShowSafeCloseTabDialog(tabControl.SelectedTab);

                if (result == "Cancel") return;

                //Delete the selected tab page status
                TabControlClass.RemoveTabPageStatus(tabControl.SelectedTab);

                //that's the reason why use the number lines of code below 
                //somehow if we set the tab page we are about to remove to another one (in this case we are about to remove selected tab page),
                //it doesn't cause the flinking
                TabPage tabToRemove = tabControl.SelectedTab;
                if (tabControl.SelectedIndex != 0)
                {
                    //set the selectedtab to the first tab page
                    tabControl.SelectedIndex = tabControl.SelectedIndex - 1;
                }
                else //if the tab we are about to remove is the first tab, just simply set selectedtab to 1
                {
                    tabControl.SelectedTab = tabControl.TabPages[1];
                }
                //remove tab 
                tabControl.TabPages.Remove(tabToRemove);
            }

        }

        /// <summary>
        /// close all tap except selected tab
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //remove all the tab except the selected Tab
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage != tabControl.SelectedTab)
                {
                    string result = DiaLogClass.ShowSafeCloseTabDialog(tabPage);
                    if (result == "Cancel") return;

                    //Delete tab status
                    TabControlClass.RemoveTabPageStatus(tabPage);

                    //Remove tab Page
                    tabControl.TabPages.Remove(tabPage);
                }
            }
        }

        /// <summary>
        /// Upper case
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uppercaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextArea textArea = TabControlClass.CurrentTextArea;

            //save the selection
            int selectionStart = textArea.SelectionStart;
            int selectionLength = textArea.SelectionLength;

            //upper text
            textArea.SelectedText = textArea.SelectedText.ToUpper();

            //retrieve selection
            textArea.SelectionStart = selectionStart;
            textArea.SelectionLength = selectionLength;

        }

        /// <summary>
        /// Lower case
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lowercaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextArea textArea = TabControlClass.CurrentTextArea;

            //save the selection
            int selectionStart = textArea.SelectionStart;
            int selectionLength = textArea.SelectionLength;

            //Lower text
            textArea.SelectedText = textArea.SelectedText.ToLower();

            //retrieve selection
            textArea.SelectionStart = selectionStart;
            textArea.SelectionLength = selectionLength;
        }

        /// <summary>
        /// version
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void versionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabPage tabPage = TabControlClass.CreateNewTabPage("ReadMe.txt");
            TextArea textArea = TabControlClass.CurrentTextArea;
            if (File.Exists("../../ReadMe.txt"))
            {
                textArea.Text = File.ReadAllText("../../ReadMe.txt");
            }
            tabPage.Text = tabPage.Text.Replace("*", "");
            textArea.ReadOnly = true;
        }

        /// <summary>
        /// View document map
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void documentMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            documentMapToolStripMenuItem.Checked = !documentMapToolStripMenuItem.Checked;
            TabControlClass.CurrentTabPageStatus.DocumentMapEnabled = documentMapToolStripMenuItem.Checked;
            TabControlClass.CurrentTextArea.DocumentMap.AutoDocumentMap(documentMapToolStripMenuItem.Checked);
        }

        /// <summary>
        /// undo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea.CanUndo())
            {
                TabControlClass.CurrentTextArea.Undo();
            }
        }

        /// <summary>
        /// redo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (TabControlClass.CurrentTextArea.CanRedo())
            {
                TabControlClass.CurrentTextArea.Redo();
            }
        }


        /// <summary>
        /// collapse all 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void collapseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Scintilla textBox = TabControlClass.CurrentTextArea;
            //foreach (Line line in textBox.Lines)
            //{
            //    line.FoldChildren(FoldAction.Contract);
            //}
        }

        /// <summary>
        /// expand All
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TabControlClass.CurrentTextArea.FoldAll(FoldAction.Expand);
        }

        /// <summary>
        /// show Indent Guides
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void showIndentGuidesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // toggle indent guides
            showIndentGuidesToolStripMenuItem.Checked = !showIndentGuidesToolStripMenuItem.Checked;
            //TabControlClass.CurrentTextArea.IndentationGuides = showIndentGuidesToolStripMenuItem.Checked ? IndentView.LookBoth : IndentView.None;
        }


        /// <summary>
        /// show white space
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void showWhiteSpaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showWhiteSpaceToolStripMenuItem.Checked = !showWhiteSpaceToolStripMenuItem.Checked;
            //TabControlClass.CurrentTextArea.ViewWhitespace = showWhiteSpaceToolStripMenuItem.Checked ? WhitespaceMode.VisibleOnlyIndent : WhitespaceMode.Invisible;
        }


        #region Search

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //SearchManagingClass.Setup(TabControlClass.CurrentTextArea);

            ////Open Panel Search
            //PanelSearch.Visible = true;

            ////focus search text box
            //TxtSearch.Focus();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            //SearchManagingClass.HighlightText(TxtSearch.Text, TabControlClass.CurrentTextArea);

        }

        private void BtnNextSearch_Click(object sender, EventArgs e)
        {
            //SearchManagingClass.SearchNext(TxtSearch.Text, TabControlClass.CurrentTextArea);
        }


        private void BtnPrevSearch_Click(object sender, EventArgs e)
        {
            //SearchManagingClass.SearchPrevious(TxtSearch.Text, TabControlClass.CurrentTextArea);
        }

        private void BtnCloseSearch_Click(object sender, EventArgs e)
        {
            PanelSearch.Visible = false;
        }
        #endregion

        #region Languages

        /// <summary>
        /// Normal text editor
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void normalTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("NormalText");

            TabControlClass.CurrentTextArea.TriggerTextAreaStyle("NormalText", true);
        }

        /// <summary>
        /// C# Language
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cSharpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("C#");
            TabControlClass.CurrentTextArea.TriggerTextAreaStyle("C#", true);
        }

        /// <summary>
        /// VB.NET language
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void vBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("VB");
            TabControlClass.CurrentTextArea.TriggerTextAreaStyle("VB", true);
        }

        /// <summary>
        /// C language
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("C");
            TabControlClass.CurrentTextArea.TriggerTextAreaStyle("C", true);
        }

        /// <summary>
        /// C++
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cPlusPlusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("C++");
            TabControlClass.CurrentTextArea.TriggerTextAreaStyle("C++", true);
        }

        #endregion

        #endregion

        #region Some events

        /// <summary>
        /// Show safe dialog when Form is about to close
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DiaLogClass.ShowSafeCloseFormDialog(tabControl) == "Cancel")
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// set the appropriate status for selected tab page when switching tab page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            //get current tab page status
            TabControlClass.TabPageStatus tabPageStatus = TabControlClass.CurrentTabPageStatus;

            //update undo and redo tool strip
            undoToolStripMenuItem.Enabled = tabPageStatus.CanUndo;
            redoToolStripMenuItem.Enabled = tabPageStatus.CanRedo;

            //set up language
            SetLanguage(tabPageStatus.Language);

            //set up document map status
            documentMapToolStripMenuItem.Checked = tabPageStatus.DocumentMapEnabled;

        }

        /// <summary>
        /// set undo and redo status ( this event occur when a tab page is added)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl_ControlAdded(object sender, ControlEventArgs e)
        {
            //variables to store previous undo and redo status
            bool previousUndoStatus = false;
            bool previousRedoStatus = false;

            //note that we can't use TabControlClass.CurrentTextBox here 
            //because this event raise before we set selectedtab by new tab page
            //see TabControlClass.CreateNewTabPage for more details
            MyRichTextBox currentTextBox = tabControl.TabPages[tabControl.TabPages.Count - 1].Controls[0] as MyRichTextBox;

            //get current text area
            TextArea currentTextArea = currentTextBox.TextArea;

            //add text changed event
            currentTextArea.TextChanged += delegate (object sen, EventArgs sube)
                {
                    if ((undoToolStripMenuItem.Enabled & currentTextArea.CanUndo()) == false)
                    {
                        //update undo and redo tool strip
                        undoToolStripMenuItem.Enabled = currentTextArea.CanUndo();
                    }

                    if ((redoToolStripMenuItem.Enabled & currentTextArea.CanRedo()) == false)
                    {
                        //update undo and redo tool strip
                        redoToolStripMenuItem.Enabled = currentTextArea.CanRedo();
                    }

                    if (previousUndoStatus != currentTextArea.CanUndo())
                    {
                        //update the status of current tabpage
                        TabControlClass.CurrentTabPageStatus.CanUndo = currentTextArea.CanUndo();
                        previousUndoStatus = currentTextArea.CanUndo();
                    }

                    if (previousRedoStatus != currentTextArea.CanRedo())
                    {
                        //update the status of current tabpage
                        TabControlClass.CurrentTabPageStatus.CanRedo = currentTextArea.CanRedo();
                        previousRedoStatus = currentTextArea.CanRedo();
                    }
                };
        }

        #endregion

        #region functions

        /// <summary>
        /// set the language item in languageToolStripMenuItem 
        /// </summary>
        private void SetLanguage(string language)
        {
            //remove previous menu item checked
            if (languageToolStripMenuItem.Tag != null)
            {
                ToolStripMenuItem itemToUncheck = (ToolStripMenuItem)languageToolStripMenuItem.Tag;
                itemToUncheck.Checked = false;
            }
            ToolStripMenuItem itemToCheck = null;

            //select langauge
            switch (language)
            {
                case "NormalText":
                    {
                        itemToCheck = normalTextToolStripMenuItem;
                        break;
                    }
                case "C#":
                    {
                        itemToCheck = cSharpToolStripMenuItem;
                        break;
                    }
                case "VB":
                    {
                        itemToCheck = vBToolStripMenuItem;
                        break;
                    }
                case "C":
                    {
                        itemToCheck = cToolStripMenuItem;
                        break;
                    }
                case "C++":
                    {
                        itemToCheck = cPlusPlusToolStripMenuItem;
                        break;
                    }
            }

            itemToCheck.Checked = true;

            //set the tag by clicked item
            //this is a trick to save the current clicked item
            languageToolStripMenuItem.Tag = itemToCheck;

            //update language tab status
            TabControlClass.CurrentTabPageStatus.Language = language;

            //set language
            TabControlClass.CurrentTextArea.Language = language ;

        }

        #endregion
    }

}


