using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using MyTextBox;
using System.Text.RegularExpressions;

/// <summary>
///In this class:
///Auto numbering
///Drag and Drop file
///Book mark
///Brace Matching highLighting
///Mark tab page on text change
///Note that: these are the things all text area should have 
///that's why we leave these things here
/// </summary>
namespace NotePad__
{
    static class MyTextBoxClass
    {
        /// <summary>
        /// Init all stuffs
        /// </summary>
        /// <param name="textBox"></param>
        public static void InitAllStuffs(MyRichTextBox textBox, TabControl tabControl)
        {
            //init basic text area color    
            //StyleClass.InitDefaultStyle(textBox);

            //init auto numbering
            textBox.NumberMargin.AutoNumbering(true);

            //init book mark
            textBox.BookMarkMargin.AutoBookMark(true);

            textBox.DocumentMap.AutoDocumentMap(false);
            //textBox.DocumentMap.AutoDocumentMap(true);

            //Allow drap and drop file into text area
            InitDragDropFile(textBox, tabControl);

            //mark a "*" in a tab page if the text in text area changed 
            //this will allow us to know whether this tab page is saved or not 
            MarkTabPageOnTextChange(textBox, tabControl);

        }

        /// <summary>
        /// Allow drag and drop file into text area
        /// </summary>
        /// <param name="textBox"></param>
        private static void InitDragDropFile(MyRichTextBox textBox, TabControl tabControl)
        {
            TextArea textArea = textBox.TextArea;
            //allow drop file
            textArea.AllowDrop = true;

            //On Drag Enter event
            textArea.DragEnter += delegate (object sender, DragEventArgs e)
            {
                //if the dropped file can be converted to specified format
                //simply means if the dropped file can be read 
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    //make an effect to look like we are dragging something
                    //this is not really important but we should do it 
                    e.Effect = DragDropEffects.Copy;
                else
                    e.Effect = DragDropEffects.None;
            };

            //On Drop File
            textArea.DragDrop += delegate (object sender, DragEventArgs e)
            {
                //if the dropped file can be read 
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    //get data in this file 
                    //note that because the various type of things can be dropped in the text area
                    //that's why it makes sense that the GetData function return an object
                    //we need to convert this into a String Array (or simply an Array) to get the data stored in the object
                    //it's important to know somehow the data store in the object just the path of the file we are about to read 
                    String[] strArray = (String[])e.Data.GetData(DataFormats.FileDrop);

                    //get the path from String Array 
                    //note that although the String Array has just only one element (the path of the file), 
                    //we can't convert directly the object above into just one string 
                    string path = strArray[0];

                    //just check to make sure the path existing
                    if (File.Exists(path))
                    {
                        //the number lines of code below is just a copy of open function in MainForm
                        //check to see if there is already this tab page being opened
                        TabPage targetTabPage = null;
                        foreach (TabPage tabPage in tabControl.TabPages)
                        {

                            if (tabPage.Name == path)
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


                        //it's better to create a new tab page  to write the stuffs than using the selected tab page,
                        //so I'll create new one
                        //this all the code below just a copy of OpenDialog function

                        //Create a new tab page
                        TabPage newTabPage = TabControlClass.CreateNewTabPage(Path.GetFileName(path));
                        //a variable to hold text box contained in tab page
                        TextArea newTextArea = (newTabPage.Controls[0] as MyRichTextBox).TextArea;

                        //Get the text of the file
                        string fileText = File.ReadAllText(path);
                        //Set the text of current text box by file Text 
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
                        tabControl.SelectedTab.Name = path;
                    }
                }
            };

        }

        /// <summary>
        /// when we type something in text area, we need to make a mark to know 
        /// if we have saved tab page or have't
        /// </summary>
        /// <param name="textBox"></param>
        /// <param name="tabControl"></param>
        private static void MarkTabPageOnTextChange(MyRichTextBox textBox, TabControl tabControl)
        {
            TextArea textArea = textBox.TextArea;
            textArea.TextChanged += delegate (object sender, EventArgs e)
            {
                if (tabControl.SelectedTab.Text.Contains("*") == false)
                {
                    tabControl.SelectedTab.Text = "*" + tabControl.SelectedTab.Text;
                }

            };
        }

        ///// <summary>
        ///// Brace Matching
        ///// </summary>
        ///// <param name="textBox"></param>
        //private static void InitBraceMatching(Scintilla textBox)
        //{
        //    //If we don't set up the color for brace, the brace matching will be invisible
        //    //we don't have to set the color here because we want this to be invisible initially
        //    //set up color
        //    //textBox.Styles[Style.BraceLight].BackColor = Color.LightGray;
        //    //textBox.Styles[Style.BraceLight].ForeColor = Color.BlueViolet;
        //    //textBox.Styles[Style.BraceBad].ForeColor = Color.Red;

        //    //a variable to store the last caret position
        //    //this isn't strictly necessary, but this makes a little more ease for computer 
        //    int lastCaretPos = 0;

        //    //Update UI event
        //    //we use this event because it's fired every time we input something 
        //    //and even when we click mouse at somewhere and making change of the caret position
        //    //if you use some other event like TextChange or KeyDown, 
        //    //the brace matching won't work when we click at any brace again
        //    textBox.UpdateUI += delegate (object sender, UpdateUIEventArgs e)
        //    {

        //        //get the current caret Position
        //        int caretPos = textBox.CurrentPosition;

        //        //if caret changed position, let's do some works
        //        //if not, don't do anything
        //        //again, this "if" line doesn't make any big difference 
        //        //we can just delete this one line  
        //        if (lastCaretPos != caretPos)
        //        {
        //            //set the last caret pos
        //            lastCaretPos = caretPos;

        //            //two variables to store the positions of braces 
        //            int bracePos1 = -1;
        //            int bracePos2 = -1;

        //            //check if there is a brace to the left of caret postion 
        //            if (IsBrace(textBox.GetCharAt(caretPos - 1)))
        //            {
        //                bracePos1 = (caretPos - 1);
        //            }
        //            else
        //            {
        //                //check if there is a brace at the caret position
        //                if (IsBrace(textBox.GetCharAt(caretPos)))
        //                    bracePos1 = caretPos;
        //            }

        //            //if there is a brace around the caret (to the left or at the caret position)
        //            if (bracePos1 >= 0)
        //            {
        //                // Find the matching brace of brace 1
        //                bracePos2 = textBox.BraceMatch(bracePos1);

        //                //if there is no brace matching
        //                if (bracePos2 == -1)
        //                {
        //                    //highlight the brace pos (by the color of bracebad set above) 
        //                    textBox.BraceBadLight(bracePos1);

        //                }
        //                else
        //                {
        //                    //highlight the braces matching
        //                    textBox.BraceHighlight(bracePos1, bracePos2);
        //                }
        //            }
        //            else
        //            {
        //                // Turn off brace matching highlighing by setting the positions by -1
        //                textBox.BraceHighlight(-1, -1);

        //            }
        //        }
        //    };
        //}

        //#region other functions

        ///// <summary>
        ///// just a function to check if a character is a brace
        ///// </summary>
        ///// <param name="c"></param>
        ///// <returns></returns>
        //private static bool IsBrace(int c)
        //{
        //    switch (c)
        //    {
        //        case '(':
        //        case ')':
        //        case '[':
        //        case ']':
        //        case '{':
        //        case '}':
        //        case '<':
        //        case '>':
        //            return true;
        //    }

        //    return false;
        //}
        //#endregion

    }
}
