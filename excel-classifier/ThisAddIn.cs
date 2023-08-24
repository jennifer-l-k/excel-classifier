using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using static System.Collections.Specialized.BitVector32;
using System.Diagnostics;
using System.ComponentModel.Design;
using System.Data;

namespace excel_classifier
{
    public partial class ThisAddIn
    {
        internal enum Classification
        {
            None,
            White,
            Green,
            Amber,
            Red
        }

        private static string classificationProperty = "Classification";

        // User control
        private UserControl _usr;
        // Custom task pane
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;

        private static readonly Dictionary<Classification, string> classificationHeader = new Dictionary<Classification, string> {
            { Classification.White, "Classified: TLP White" },
            { Classification.Green, "Classified: TLP Green" },
            { Classification.Amber, "Classified: TLP Amber" },
            { Classification.Red, "Classified: TLP Red" },
        };

        private static readonly Dictionary<Classification, Microsoft.Office.Interop.Excel.XlRgbColor> classificationTextColor = new Dictionary<Classification, Microsoft.Office.Interop.Excel.XlRgbColor> {
             
            { Classification.White, Excel.XlRgbColor.rgbDarkGray },
            { Classification.Green, Excel.XlRgbColor.rgbGreen },
            { Classification.Amber, Excel.XlRgbColor.rgbOrange },
            { Classification.Red, Excel.XlRgbColor.rgbRed},
        };

        private static readonly Dictionary<Classification, string> classificationProperties = new Dictionary<Classification, string> {
            { Classification.White, "TLP:WHITE" },
            { Classification.Green, "TLP:GREEN" },
            { Classification.Amber, "TLP:AMBER" },
            { Classification.Red, "TLP:RED" },
        };

        internal void ToggleTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);


            //Create an instance of the user control
            _usr = new TaskPane();
            // Connect the user control and the custom task pane 
            _taskPane = CustomTaskPanes.Add(_usr, "Classifier TLP");
            _taskPane.Width = 300;
            _taskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Classification classification = GetClassification();
            if (classification == Classification.None)
            {
                Cancel = true;
                MessageBox.Show("Please classify document before saving. See the Add-In / Classifier Toolbar Menu.", "Classifier prevented saving", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // re-apply changes as defense in depth
                Classify(classification);
            }
        }

        Classification GetClassification()
        {
            string classificationString = ReadDocumentProperty(classificationProperty);
            if (!classificationProperties.ContainsValue(classificationString))
            {
                return Classification.None;
            }
            else
            {
                Classification key = classificationProperties.FirstOrDefault(x => x.Value == classificationString).Key;
                return key;
            }
        }

        internal void Classify(Classification classification)
        {

            Debug.Print(classification.ToString());

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            if (Application.Interactive)
            {
                // Exit cell-edit mode for Excel (no API), otherwise Range.Insert and headerRange.Value will fail
                Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;
                Globals.ThisAddIn.Application.ActiveWindow.Activate();
                SendKeys.Flush();
                SendKeys.SendWait("{ENTER}");
                r.Select(); // restore selection
            }

            if (GetClassification() == Classification.None) {
                // Shift down everything by one row to make space for header
                activeWorksheet.Rows[1].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }

            // Set custom document property
            SetDocumentProperty(classificationProperty, classificationProperties[classification]);


            // Set classification notice
            activeWorksheet.Rows[1].Cells.Clear();
            Range headerRange = activeWorksheet.Range["A1"];
            headerRange.Font.Color = classificationTextColor[classification];
            headerRange.Font.Size = 16;
            //headerRange.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            headerRange.Value = classificationHeader[classification];
        }

        private string ReadDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)this.Application.ActiveWorkbook.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        private void SetDocumentProperty(string propertyName, string value)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)this.Application.ActiveWorkbook.CustomDocumentProperties;

            if (ReadDocumentProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }

            properties.Add(propertyName, false,
                Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                value);
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
 