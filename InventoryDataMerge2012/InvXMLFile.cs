using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace InventoryDataMerge2012
{
    class InvXMLFile
    {
        InventoryWorkBookClass invWb;
        OpenFileDialog fileXml = new OpenFileDialog();
        List<XElement> systems;
        internal InvXMLFile(InventoryWorkBookClass invWbp)
        {
            invWb = invWbp;
            fileXml.Title = "Open IDC Data File - TaxAideInv2012.xml";
            fileXml.Filter = "XML files (*.xml)|*.xml";
            fileXml.FileName = "TaxAideInv2012.xml";
        }

        internal void GetIDCXmlData()
        {
            DialogResult dlg = fileXml.ShowDialog();
            if (dlg == DialogResult.Cancel)
                Dispose();
            XDocument doc = null;
            try
            {
                doc = XDocument.Load(fileXml.FileName);
            }
            catch (Exception)
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("This File is not a correctly formatted XML file. \rExiting!", "IDC Data Merge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Dispose();
            }
            if (doc.Nodes().FirstOrDefault().ToString() != "<!--IDC XML Version 2012.01-->")
            {
                MessageBox.Show("This file is not a 2012 IDC Inventory file.\rExiting!", "IDC Data Merge");
                Dispose();
            }
            systems = doc.Elements("Systems").Elements().ToList();
            if (systems.Count == 0)
            {
                MessageBox.Show("This file contains no system data\rExiting!", "IDC Data Merge");
                Dispose();
            }
            //At this point have a List of Xelements each of which is a system
        }

        private void Dispose()
        {
            invWb.Dispose();
            Environment.Exit(1);
        }

        internal void xmlData2End()
        {
            foreach (XElement  item in systems)
            {
                if (item.HasElements)
                {
                     invWb.XLRowFromXml(item); 
                }
            }
        }
    }
}
