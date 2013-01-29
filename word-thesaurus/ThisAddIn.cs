using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System;

namespace word_thesaurus
{
    public partial class ThisAddIn
    {
        private string key = ""; // http://www.abbreviations.com/api.php

        private Word.Application app;
        private Word.Selection currentSelection;
        private Office.CommandBar commandBar;
        private Office.CommandBarPopup popup;
        private Office.CommandBarButton button;

        ArrayList menus = new ArrayList();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;
            app.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(app_WindowBeforeRightClick);
            commandBar = app.CommandBars["Text"];
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            app.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(app_WindowBeforeRightClick);
            popup.Delete();
            button.Delete();
            commandBar.Delete();

        }

        public void app_WindowBeforeRightClick(Word.Selection selection, ref bool Cancel)
        {
            if (selection != null && !string.IsNullOrEmpty(selection.Text))
            {
                string text = selection.Text;
                currentSelection = selection;
                ShowMenu(text.Trim());
            }
        }

        public string UriSetup(string query)
        {
            UriBuilder uri = new UriBuilder();
            uri.Scheme = "http";
            uri.Host = "www.stands4.com";
            uri.Path = "services/v2/syno.php";
            uri.Query = "uid=2584&tokenid=" + key +  "&word=" + query; // HIDE TOKENID
            return uri.ToString();
        }

        public string[] Request(string query)
        {
            XmlReader reader = XmlReader.Create(UriSetup(query));

            string content = "", temp = "";
            while (reader.Read())
            {
                if (reader.Name.Equals("synonyms"))
                {
                    temp = reader.ReadElementContentAsString();
                    if (temp.Length.CompareTo(content.Length) > 0)
                    {
                        content = temp;
                    }
                }
            }
            reader.Close();

            string[] words = content.Split(',');
            return words;
        }

        public void RemoveMenu()
        {
            foreach (Office.CommandBarButton btn in menus)
            {
                btn.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(button_Click);
                btn.Visible = false;
            }
            menus.Clear();
            commandBar.Reset();
        }
        public void AddMenu(string[] words)
        {
            popup = (Office.CommandBarPopup)commandBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, 1, false);
            popup.accName = "Theasurus";
            popup.Tag = "Theasurus";
            popup.Visible = true;

            foreach (string word in words)
            {
                string w = word.Trim();
                button = (Office.CommandBarButton)popup.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, popup.Controls.Count + 1, false);
                button.Caption = w;
                button.Tag = w;
                button.Visible = true;
                button.Click += new Office._CommandBarButtonEvents_ClickEventHandler(button_Click);
                menus.Add(button);
            }
        }

        public void ShowMenu(string selectedText)
        {
            if (commandBar.Controls.Count > 0)
            {
                RemoveMenu();
            }
            string[] words = Request(selectedText);
            AddMenu(words);
        }

        private void button_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            currentSelection.TypeText(ctrl.Tag);
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
