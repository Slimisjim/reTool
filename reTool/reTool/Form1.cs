using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;

namespace reTool
{
    public partial class Form1 : Form
    {
        string path;
        Workbook wb;
        Worksheet excelSheet;

        string CustomerInfoSource = "DrawingProperties";
        string UsageInfoSource = "PreDesign";
        
        //custInfo locations
        int ARr = 4, ARc = 3,
        cNamer = 5, cNamec = 3,
        Streetr =6, Streetc = 3,
        Cityr = 7, Cityc = 3,
        Stater = 8, Statec = 3,
        Zipr = 9, Zipc = 3,
        Utilityr = 10, Utilityc = 3,
        Meterr = 11, Meterc = 3,
        custUsager = 24, custUsagec = 10,
        custTrueUsager = 4, custTrueUsagec = 15,
        estModulesr = 5, estModulesc = 14,

        subStart = 4;

        public Form1()
        {
            InitializeComponent();
        }
        
        public void openTool(string source)
        {
            Microsoft.Office.Interop.Excel.Application excelObj = new Microsoft.Office.Interop.Excel.Application();
            wb = excelObj.Workbooks.Open(source);

            //Worksheet excelSheet = (Worksheet)wb.Worksheets.get_Item(2);
            //excelSheet.Activate();

            //MessageBox.Show(excelSheet.Cells[1, 1].Value.ToString());
            //MessageBox.Show(excelSheet.Name);
        }

        private void getCustomerInfo(string path)
        {

            // set active sheet
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[CustomerInfoSource];

            foreach (Control tb in custInfo.Controls)
            {
                MessageBox.Show(tb.Name.Substring(subStart));
                if (tb is System.Windows.Forms.Button)
                {
                    if (tb.Name.Substring(subStart) == "AR")
                    {
                        string AR = excelSheet.Cells[ARr, ARc].Text;
                        tb.Text = AR.Substring(AR.IndexOf("-")+1);
                    } else if (tb.Name.Substring(subStart) == "Street")
                    {
                        tb.Text = excelSheet.Cells[Streetr, Streetc].Text;
                    } else if (tb.Name.Substring(subStart) == "Name")
                    {
                        //remove first name
                        string lastName = excelSheet.Cells[cNamer, cNamec].Text;

                        tb.Text = lastName.Substring(lastName.IndexOf(" ")+1);
                    }
                    else if (tb.Name.Substring(subStart) == "Utility")
                    {
                        tb.Text = excelSheet.Cells[Utilityr, Utilityc].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "Meter")
                    {
                        tb.Text = excelSheet.Cells[Meterr, Meterc].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "City")
                    {
                        tb.Text = excelSheet.Cells[Cityr, Cityc].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "State")
                    {
                        tb.Text = excelSheet.Cells[Stater, Statec].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "Zip")
                    {
                        tb.Text = excelSheet.Cells[Zipr, Zipc].Text;
                    }
                }
            }
        }

        private void getUsageInfo()
        {
            foreach (Control tb in usageInfo.Controls)
            {
                if(tb is System.Windows.Forms.Button)
                {
                    if(tb.Name.Substring(subStart) == "TrueUsage")
                    {
                        //set source
                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[UsageInfoSource];

                        tb.Text = excelSheet.Cells[custTrueUsager, custTrueUsagec].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "TargetUsage")
                    {
                        //set source
                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[CustomerInfoSource];

                        tb.Text = excelSheet.Cells[custUsager, custUsagec].Text;
                    }
                    else if (tb.Name.Substring(subStart) == "EstModules")
                    {
                        //set source
                        excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[UsageInfoSource];

                        tb.Text = excelSheet.Cells[estModulesr, estModulesc].Text;
                    }

                }
            }
            
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog1.FileName;
            }
            //
            openTool(path);

            getCustomerInfo(path);
            getUsageInfo();
        }


        private void closeTool()
        {
            wb.Close();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            closeTool();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            closeTool();
        }

        private void dispAR_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispAR.Text);
        }

        private void dispName_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispName.Text);
        }

        private void dispStreet_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispStreet.Text);
        }

        private void dispCity_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispCity.Text);
        }

        private void dispState_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispState.Text);
        }

        private void dispZip_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispZip.Text);
        }

        private void dispUtility_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispUtility.Text);
        }

        private void dispMeter_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispMeter.Text);
        }

        private void dispTargetUsage_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispTargetUsage.Text);
        }

        private void dispTrueUsage_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispTrueUsage.Text);
        }

        private void dispAR_Click_1(object sender, EventArgs e)
        {
            Clipboard.SetText(dispAR.Text);
        }

        private void dispEstModules_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(dispEstModules.Text);
        }
    }
}
