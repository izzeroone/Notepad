using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using MyTextBox;
/// <summary>
/// open dialog, close dialog, ...
/// </summary>
namespace NotePad__
{
    class DiaLogClass
    {

        /// <summary>
        /// Open Dialog
        /// </summary>
        /// <param name="tabControl"></param>
        public static void ShowOpenDialog(TabControl tabControl)
        {
            //create a new file dialog 
            OpenFileDialog openFileDialog = new OpenFileDialog();

            //Pop up the file dialog and check if user press open button 
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //check to see if there is already this tab page
                TabPage targetTabPage = null;
                foreach (TabPage tabPage in tabControl.TabPages)
                {

                    if (tabPage.Name == openFileDialog.FileName)
                    {
                        targetTabPage = tabPage;
                        break;
                    }
                }
                //if this tab page has already opened, just focus this tab and return
                if (targetTabPage != null)
                {
                    tabControl.SelectedTab = targetTabPage;
                    return;
                }

                //Create a new tab page
                TabPage newTabPage = TabControlClass.CreateNewTabPage(openFileDialog.SafeFileName);
                //a variable to hold text box contained in tab page
                TextArea newTextArea = TabControlClass.CurrentTextArea;
                //Get the path of the File
                string filePath = openFileDialog.FileName;
                //Get the text of the file
                string fileText = File.ReadAllText(filePath);
                //Set the text of current text box by file Text 
                //we don't want Undo to record our text here
                newTextArea.StopRecordingUndo();
                newTextArea.Text = fileText;
                newTextArea.ContinueRecordingUndo();

                //when we set the text by the code above, we changed the text in the text area 
                //and accidentally lead to the MarkTabPage event
                //that's the reason why we have to do this to unmark the tabPage initially
                newTabPage.Text = newTabPage.Text.Replace("*", "");
                
                //this is a trick to save the path(FileName) of the saved tab page
                //and the next time if this tab page has already had a name, we shouldn't open the savefiledialog again
                //and just implicitly save 
                tabControl.SelectedTab.Name = openFileDialog.FileName;
            }

        }

        /// <summary>
        /// Open Save file Dialog
        /// </summary>
        /// <param name="tabPage">this variable should equal TabControl.selectedTab</param>
        public static void ShowSaveDialog(TabPage tabPage)
        {
            //get the current text box
            TextArea currentTextArea = (tabPage.Controls[0] as MyRichTextBox).TextArea;

            //if selected tab page has already had a name, just implicitly save it
            if (tabPage.Name != "")
            {
                //implicitly save//
                using (Stream s = File.Open(tabPage.Name, FileMode.Create))
                {
                    //get the streamwriter of the new file 
                    using (StreamWriter sw = new StreamWriter(s))
                    {
                        //Get the text of the current text box and write it to streamwriter
                        sw.Write(currentTextArea.Text);

                        tabPage.Text = Path.GetFileName(tabPage.Name);
                        return;
                    }
                }
            }

            //Create a save file Dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            //Pop up the save File Dialog and check if user press Save button
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Declare a Stream variable to hold the open file to write in
                using (Stream s = File.Open(saveFileDialog.FileName, FileMode.Create))
                {
                    //get the streamwriter of the new file 
                    using (StreamWriter sw = new StreamWriter(s))
                    {
                        //Get the text of the current text box and write it to streamwriter
                        sw.Write(currentTextArea.Text);

                        //change the text title of the tab by file name
                        tabPage.Text = Path.GetFileName(saveFileDialog.FileName);

                        //this is a trick to save the path(FileName) of the saved tab page
                        //and the next time if this tab page has already had a name, we shouldn't open the savefiledialog again
                        //and just implicitly save 
                        tabPage.Name = saveFileDialog.FileName;
                    }
                }
            }
        }

        /// <summary>
        /// Open Save As Dialog
        /// </summary>
        /// <param name="tabPage">this variable should equal tabControl.SelectedTab</param>
        public static void ShowSaveAsDialog(TabPage tabPage)
        {
            //get the current text box
            TextArea currentTextArea = (tabPage.Controls[0] as MyRichTextBox).TextArea;

            //Create a save file Dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            //Pop up the save File Dialog check if user press Save button
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Declare a Stream variable to hold the open file to write in
                using (Stream s = File.Open(saveFileDialog.FileName, FileMode.Create))
                {
                    //Write the text into the new file 
                    using (StreamWriter sw = new StreamWriter(s))
                    {
                        //Get the text of the current text box
                        sw.Write(currentTextArea.Text);

                        //change the name of the tab by file name
                        tabPage.Text = Path.GetFileName(saveFileDialog.FileName);

                        //this is a trick to save the path(FileName) of the saved tab page
                        //and the next time if this tab page has already had a name, we shouldn't open the savefiledialog again
                        //and just implicitly save 
                        tabPage.Name = saveFileDialog.FileName;
                    }
                }
            }
        }

        /// <summary>
        /// Safe Save File Dialog
        /// </summary>
        /// <param name="tabPage"></param>
        /// <returns></returns>
        public static string ShowSafeCloseTabDialog(TabPage tabPage)
        {
            if (tabPage.Text.Contains("*"))
            {
                //get the tabName without "*"  (this is considered the real name)
                string tabName = tabPage.Text;
                tabName = tabName.Replace("*", "");

                string text = "You haven't save " + tabName + "\n Do you want to save it before closing?";
                DialogResult result = MessageBox.Show(text, "Be careful !!!", MessageBoxButtons.YesNoCancel);

                if (result == DialogResult.Yes)
                {
                    ShowSaveDialog(tabPage);
                    return "Yes";
                }
                else
                {
                    if (result == DialogResult.No)
                    {
                        return "No";
                    }
                    else
                    {
                        return "Cancel";
                    }
                }
            }
            return "Nothing";
        }

        /// <summary>
        /// Safe Close Form Dialog
        /// </summary>
        /// <param name="tabPage"></param>
        /// <returns></returns>
        public static string ShowSafeCloseFormDialog(TabControl tabControl)
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                string result = ShowSafeCloseTabDialog(tabPage);
                if (result == "Cancel") return "Cancel";
            }
            return "Nothing";
        }

    }
}
