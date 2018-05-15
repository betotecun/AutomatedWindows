using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace AutomatedWindows
{
    public partial class FormMailClient : Form
    {
        public FormMailClient()
        {
            InitializeComponent();
        }

        DataTable dt;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook._NameSpace _ns = _app.GetNamespace("MAPI");
                Outlook.MAPIFolder Inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.Items items = Inbox.Items;
                Outlook.MAPIFolder userfolder = Inbox.Folders["AutomateFolder"];
                Outlook.MAPIFolder moveto = Inbox.Folders["AutomateProcessed"];
                _ns.SendAndReceive(true);
            
            
                dt = new DataTable("workBox");
                dt.Columns.Add("From", typeof(string));
                dt.Columns.Add("To", typeof(string));
                dt.Columns.Add("Subject", typeof(string));
                dataGridView1.DataSource = dt;

                foreach (Outlook.MailItem item in userfolder.Items)
                {

                    dt.Rows.Add(new object[]
                        {
                        item.SenderName, item.To, item.Subject
                        });
                    if (item.SenderName.Contains("Bryan"))
                    {
                        MessageBox.Show("Match process based on table rules: " +
                            item.SenderName);
                        item.Move(moveto);
                    }
                }

            }

           catch (Exception ex)
                {
                    MessageBox.Show("OL-ERROR: " + ex.Message);
                }

            //MessageBox.Show("testing");
        }
    }
}
