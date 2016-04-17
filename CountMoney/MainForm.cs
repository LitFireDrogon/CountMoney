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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CountMoney
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            DataGridView dataGridView1 = new DataGridView();
            DataGridViewRolloverCellColumn col = new DataGridViewRolloverCellColumn();
            dataGridView1.Columns.Add(col);
            dataGridView1.Rows.Add(new string[] { "" });
            dataGridView1.Rows.Add(new string[] { "" });
            dataGridView1.Rows.Add(new string[] { "" });
            dataGridView1.Rows.Add(new string[] { "" });
            this.Controls.Add(dataGridView1);
            this.button_OK.Enabled = false;
            this.button_Calculate.Enabled = false;
        }

        #region DataGridViewRolloverCell
        public class DataGridViewRolloverCell : DataGridViewTextBoxCell
        {
            protected override void Paint(
                Graphics graphics,
                Rectangle clipBounds,
                Rectangle cellBounds,
                int rowIndex,
                DataGridViewElementStates cellState,
                object value,
                object formattedValue,
                string errorText,
                DataGridViewCellStyle cellStyle,
                DataGridViewAdvancedBorderStyle advancedBorderStyle,
                DataGridViewPaintParts paintParts)
            {
                // Call the base class method to paint the default cell appearance.
                base.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState,
                    value, formattedValue, errorText, cellStyle,
                    advancedBorderStyle, paintParts);

                // Retrieve the client location of the mouse pointer.
                Point cursorPosition =
                    this.DataGridView.PointToClient(Cursor.Position);

                // If the mouse pointer is over the current cell, draw a custom border.
                if (cellBounds.Contains(cursorPosition))
                {
                    Rectangle newRect = new Rectangle(cellBounds.X + 1,
                        cellBounds.Y + 1, cellBounds.Width - 4,
                        cellBounds.Height - 4);
                    graphics.DrawRectangle(Pens.Red, newRect);
                }
            }

            // Force the cell to repaint itself when the mouse pointer enters it.
            protected override void OnMouseEnter(int rowIndex)
            {
                this.DataGridView.InvalidateCell(this);
            }

            // Force the cell to repaint itself when the mouse pointer leaves it.
            protected override void OnMouseLeave(int rowIndex)
            {
                this.DataGridView.InvalidateCell(this);
            }

        }

        public class DataGridViewRolloverCellColumn : DataGridViewColumn
        {
            public DataGridViewRolloverCellColumn()
            {
                this.CellTemplate = new DataGridViewRolloverCell();
            }
        } 
        #endregion

        #region DataGridViewDateTimePicker
        public class DataGridViewCalendarColumn : DataGridViewColumn
        {
            public  DataGridViewCalendarColumn() : base(new DataGridViewCalendarCell())
            {
            }

            public override DataGridViewCell CellTemplate
            {
                get
                {
                    return base.CellTemplate;
                }
                set
                {
                    // Ensure that the cell used for the template is a CalendarCell.
                    if (value != null &&
                        !value.GetType().IsAssignableFrom(typeof(DataGridViewCalendarCell)))
                    {
                        throw new InvalidCastException("Must be a DataGridViewCalendarColumn");
                    }
                    base.CellTemplate = value;
                }
            }
        }

        public class DataGridViewCalendarCell : DataGridViewTextBoxCell
        {

            public DataGridViewCalendarCell()
                : base()
            {
                // Use the short date format.
                this.Style.Format = "d";
            }

            public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
            {
                // Set the value of the editing control to the current cell value.
                base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
                DataGridViewCalendarEditingControl ctl = DataGridView.EditingControl as DataGridViewCalendarEditingControl;
                // Use the default row value when Value property is null.
                if (this.Value == null)
                {
                    ctl.Value = (DateTime)this.DefaultNewRowValue;
                }
                else
                {
                    ctl.Value = (DateTime)this.Value;
                }
            }

            public override Type EditType
            {
                get
                {
                    // Return the type of the editing control that CalendarCell uses.
                    return typeof(DataGridViewCalendarEditingControl);
                }
            }

            public override Type ValueType
            {
                get
                {
                    // Return the type of the value that CalendarCell contains.

                    return typeof(DateTime);
                }
            }

            public override object DefaultNewRowValue
            {
                get
                {
                    // Use the current date and time as the default value.
                    return DateTime.Now;
                }
            }
        }

        public class DataGridViewCalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
        {
            DataGridView dataGridView;
            private bool valueChanged = false;
            int rowIndex;

            public DataGridViewCalendarEditingControl()
            {
                this.Format = DateTimePickerFormat.Short;
            }

            // Implements the IDataGridViewEditingControl.EditingControlFormattedValue 
            // property.
            public object EditingControlFormattedValue
            {
                get
                {
                    return this.Value.ToShortDateString();
                }
                set
                {
                    if (value is String)
                    {
                        try
                        {
                            // This will throw an exception of the string is 
                            // null, empty, or not in the format of a date.
                            this.Value = DateTime.Parse((String)value);
                        }
                        catch
                        {
                            // In the case of an exception, just use the 
                            // default value so we're not left with a null
                            // value.
                            this.Value = DateTime.Now;
                        }
                    }
                }
            }

            // Implements the 
            // IDataGridViewEditingControl.GetEditingControlFormattedValue method.
            public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context)
            {
                return EditingControlFormattedValue;
            }

            // Implements the 
            // IDataGridViewEditingControl.ApplyCellStyleToEditingControl method.
            public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle)
            {
                this.Font = dataGridViewCellStyle.Font;
                this.CalendarForeColor = dataGridViewCellStyle.ForeColor;
                this.CalendarMonthBackground = dataGridViewCellStyle.BackColor;
            }

            // Implements the IDataGridViewEditingControl.EditingControlRowIndex 
            // property.
            public int EditingControlRowIndex
            {
                get
                {
                    return rowIndex;
                }
                set
                {
                    rowIndex = value;
                }
            }

            // Implements the IDataGridViewEditingControl.EditingControlWantsInputKey 
            // method.
            public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey)
            {
                // Let the DateTimePicker handle the keys listed.
                switch (key & Keys.KeyCode)
                {
                    case Keys.Left:
                    case Keys.Up:
                    case Keys.Down:
                    case Keys.Right:
                    case Keys.Home:
                    case Keys.End:
                    case Keys.PageDown:
                    case Keys.PageUp:
                        return true;
                    default:
                        return !dataGridViewWantsInputKey;
                }
            }

            // Implements the IDataGridViewEditingControl.PrepareEditingControlForEdit 
            // method.
            public void PrepareEditingControlForEdit(bool selectAll)
            {
                // No preparation needs to be done.
            }

            // Implements the IDataGridViewEditingControl
            // .RepositionEditingControlOnValueChange property.
            public bool RepositionEditingControlOnValueChange
            {
                get
                {
                    return false;
                }
            }

            // Implements the IDataGridViewEditingControl
            // .EditingControlDataGridView property.
            public DataGridView EditingControlDataGridView
            {
                get
                {
                    return dataGridView;
                }
                set
                {
                    dataGridView = value;
                }
            }

            // Implements the IDataGridViewEditingControl
            // .EditingControlValueChanged property.
            public bool EditingControlValueChanged
            {
                get
                {
                    return valueChanged;
                }
                set
                {
                    valueChanged = value;
                }
            }

            // Implements the IDataGridViewEditingControl
            // .EditingPanelCursor property.
            public Cursor EditingPanelCursor
            {
                get
                {
                    return base.Cursor;
                }
            }

            protected override void OnValueChanged(EventArgs eventargs)
            {
                // Notify the DataGridView that the contents of the cell
                // have changed.
                valueChanged = true;
                this.EditingControlDataGridView.NotifyCurrentCellDirty(true);
                base.OnValueChanged(eventargs);
            }
        }
        #endregion

        private void button_ChooseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox_Excel.Text = openFileDialog.FileName;
                GetSheet(this.textBox_Excel.Text);
                if(this.comboBox_SheetName.Items.Count>0)
                {
                    this.comboBox_SheetName.Text = this.comboBox_SheetName.Items[0].ToString();
                }
            }

            if(this.textBox_Excel.Text != "" && this.comboBox_SheetName.Text !="" )
            {
                this.button_OK.Enabled = true;
            }
        }
        
        private void button_OK_Click(object sender, EventArgs e)
        {
            //設定Excel路徑,使用OleDB獲取Excel資料
            string pathConnect = string.Empty;
            bool isXls = this.textBox_Excel.Text.EndsWith(".xlsx");//isXls = true 表示EXCEL版本為2007以上
            if (isXls)
            {
                pathConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + this.textBox_Excel.Text + ";Extended Properties =\"Excel 12.0;HDR=Yes;IMEX=2\";";
            }
            else
            {
                pathConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.textBox_Excel.Text + ";Extended Properties =\"Excel 4.0;HDR=Yes;IMEX=2\";";
            }
            OleDbConnection myConn = new OleDbConnection(pathConnect);
            OleDbDataAdapter myDataAdpter = new OleDbDataAdapter("Select * from [" + this.comboBox_SheetName.Text + "$]", myConn);

            DataTable dt = new DataTable();
            myDataAdpter.Fill(dt);

            //將Excel資料塞入GRID
            if (dt.Rows.Count==0)
            {
                MessageBox.Show("請檢察Excel是否有資料", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                this.dataGridView1.Rows.Clear();
                this.dataGridView1.Rows.Add(dt.Rows.Count);
                for (int i = 0; i < dt.Columns.Count - 1; i++)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        this.dataGridView1.Rows[j].Cells[i].Value = dt.Rows[j][i];
                    }
                }
                this.button_Calculate.Enabled = true;
            }
        }

        private void button_Calculate_Click(object sender, EventArgs e)
        {
            if(this.dataGridView1.RowCount >1)
            {
                List<object> gridViewData = new List<object>();
                decimal totalAmount = 0M;
                for(int i=0;i<this.dataGridView1.RowCount;i++)
                {
                    totalAmount = totalAmount + ToDecimal(this.dataGridView1.Rows[i].Cells[1].Value);
                }
                this.textBox_TotalAmount.Text = totalAmount.ToString();
            }
            else
            {
                this.textBox_TotalAmount.Text = "0";
            }
        }
        /// <summary>
        /// 獲取Excel所有的工作表名稱
        /// </summary>
        /// <param name="filePath">Excel路徑</param>
        private void GetSheet(string filePath)
        {
            Excel.Application myExcel = new Excel.Application();
            Excel.Workbook myWB;
            myWB = myExcel.Workbooks.Open(filePath) as Excel.Workbook;
            string[] sheets = new string[myWB.Worksheets.Count];
            int i = 0;
            foreach (Excel.Worksheet sheet in myWB.Worksheets)
            {
                sheets[i] = sheet.Name;
                i++;
            }
            //將工作表名稱加入comboBox_SheetName此控件
            this.comboBox_SheetName.Text = string.Empty;
            this.comboBox_SheetName.Items.Clear();
            this.comboBox_SheetName.Items.AddRange(sheets);

            myWB.Close();
            myExcel.Quit();
        }
        /// <summary>
        /// 型態轉為Decimal
        /// </summary>
        /// <param name="sth"></param>
        /// <returns></returns>
        private decimal ToDecimal(object sth)
        {
            decimal decimalVal = 0m;
            decimalVal = System.Convert.ToDecimal(sth);
            return decimalVal;
        }
    }
}
