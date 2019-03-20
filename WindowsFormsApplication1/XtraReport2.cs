using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using DevExpress.XtraReports.UI.PivotGrid;

namespace WindowsFormsApplication1
{
    public partial class XtraReport2 : DevExpress.XtraReports.UI.XtraReport
    {
        public XtraReport2()
        {
            InitializeComponent();
        }

        private void xrPivotGrid1_CustomFieldValueCells(object sender, DevExpress.XtraReports.UI.PivotGrid.PivotCustomFieldValueCellsEventArgs e)
        {
            bool isColumn = true;
            for (int i = e.GetCellCount(isColumn) - 1; i >= 0; i--)
            {
                FieldValueCell cell = e.GetCell(isColumn, i);
                if (cell != null)
                    e.Remove(cell);

            }
        }
    }
}
