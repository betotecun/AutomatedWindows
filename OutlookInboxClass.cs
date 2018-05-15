using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
//using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Windows.Forms;
using System.Configuration;
using System.Xml;

namespace AutomatedWindows
{
    class OutlookInboxClass
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        private void OutlookClient(string foldername)
        {
            Outlook._Application _app = new Outlook.Application();
            Outlook._NameSpace _ns = _app.GetNamespace("MAPI");
            Outlook.MAPIFolder Inbox = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items items = Inbox.Items;
            Outlook.MAPIFolder userfolder = Inbox.Folders[foldername];
        }

        private void testXml()
        {
            XmlDocument xd = new XmlDocument();
            xd.AppendChild(xd.CreateXmlDeclaration("1.0", "utf-8", ""));

            XmlNode root = xd.CreateElement("Message");
            xd.AppendChild(root);

            XmlNode result = xd.CreateElement("From");
            result.InnerText = "FromBeto";

            XmlAttribute type = xd.CreateAttribute("Status");
            type.Value = "Done";

            result.Attributes.Append(type);
            root.AppendChild(result);
        }
    }
}
