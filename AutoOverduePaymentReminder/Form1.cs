using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace AutoOverduePaymentReminder
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Toprint();
        }
        #region
        public void Toprint()
        {
            WriteIntoWord wiw = new WriteIntoWord();
            string FilePath = System.Windows.Forms.Application.StartupPath+ @"\OverduePaymentReminder.dot";     //模板路径 
            string BookmarkABSdate = "ABSdate";
            string FillABSdate = "2018-03-06";
            string BookmarkCustomer = "Customer";
            string FillCustomer = "L&V";
            string BookmarkInvoiceNo = "InvoiceNo";
            string FillInvoiceNo = "IN10001";
            string BookmarkInvoiceDate = "InvoiceDate";
            string FillInvoiceDate = "2018-02-06";
            string BookmarkDueDate = "DueDate";
            string FillDueDate = "2018-02-10";
            string BookmarkOutstandingAmount = "OutstandingAmount";
            string FillOutstandingAmount = "500000";
            string BookmarkEndDate = "EndDate";
            string FillEndDate = "2018-03-10";
            string BookmarkContactNo = "ContactNo";
            string FillContactNo = "13400000000";
            string BookmarkSalesname = "Salesname";
            string FillSalesname = "Justin";
            string SaveDocPath = System.Windows.Forms.Application.StartupPath + @"\OverduePaymentReminder.doc"; 
            wiw.OpenDocument(FilePath);
            wiw.WriteIntoDocument(BookmarkABSdate, FillABSdate);
            wiw.WriteIntoDocument(BookmarkCustomer, FillCustomer);
            wiw.WriteIntoDocument(BookmarkInvoiceNo, FillInvoiceNo);
            wiw.WriteIntoDocument(BookmarkInvoiceDate, FillInvoiceDate);
            wiw.WriteIntoDocument(BookmarkDueDate, FillDueDate);
            wiw.WriteIntoDocument(BookmarkOutstandingAmount, FillOutstandingAmount);
            wiw.WriteIntoDocument(BookmarkEndDate, FillEndDate);
            wiw.WriteIntoDocument(BookmarkContactNo, FillContactNo);
            wiw.WriteIntoDocument(BookmarkSalesname, FillSalesname);
            wiw.Save_CloseDocument(SaveDocPath);
            MessageBox.Show("OK");
            this.Dispose();
            this.Close();

        }
        #endregion
        public class WriteIntoWord
        {
            ApplicationClass app = null;   //定义应用程序对象 
            Document doc = null;
            Object missing = System.Reflection.Missing.Value; //定义空变量
            Object isReadOnly = false;
            // 向 word 文档写入数据 
            public void OpenDocument(string FilePath)
            {
                object filePath = FilePath;//文档路径
                app = new ApplicationClass(); //打开文档 
                doc = app.Documents.Open(ref filePath, ref missing, ref missing, ref missing,
ref missing, ref missing, ref missing, ref missing);
                doc.Activate();//激活文档
            }
            /// <summary>
            /// </summary>
            ///<param name="parLableName">域标签</param> 
            /// <param name="parFillName">写入域中的内容</param> 
            ///
            //打开word，将对应数据写入word里对应书签域 
            public void WriteIntoDocument(string BookmarkName, string FillName)
            {
                object bookmarkName = BookmarkName;
                Bookmark bm = doc.Bookmarks.get_Item(ref bookmarkName);//返回书签
                bm.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                bm.Range.Text = FillName;//设置书签域的内容                
                //bm.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineThick;
            }
            /// <summary> 
            /// 保存并关闭 
            /// </summary> 
            /// <param name="parSaveDocPath">文档另存为的路径</param>
            /// 
            public void Save_CloseDocument(string SaveDocPath)
            {
                string _SaveDocPath = SaveDocPath.Replace(".pdf", ".doc");
                object savePath = _SaveDocPath;  //文档另存为的路径 
                Object saveChanges = app.Options.BackgroundSave;//文档另存为 
                Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                Document document = null;
                doc.SaveAs(ref savePath, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing);                
                doc.Close(ref saveChanges, ref missing, ref missing);//关闭文档
                app.Quit(ref missing, ref missing, ref missing);//关闭应用程序
                application.Visible = false;
                document = application.Documents.Open(savePath);
                document.ExportAsFixedFormat(SaveDocPath, WdExportFormat.wdExportFormatPDF);
                document.Close();
                application.Quit();

            }
            public void WriteIntoTable(string BookmarkName, System.Data.DataTable _dt)
            {
                Microsoft.Office.Interop.Word.Range tmpRange, srcRange, desRange;
                Microsoft.Office.Interop.Word.Table operTable;
                int Trow = _dt.Rows.Count;
                int cur_row = 1;
                try
                {//使用一个特殊的标签 table_bookmark_template_
                    tmpRange = doc.Bookmarks.get_Item(BookmarkName).Range;//
                    operTable = tmpRange.Tables[1];//得到该书签所在的表，以它为报表的循环模板
                    for (int j = 0; j < Trow; j++)
                    {//插入需要重复循环的行数loopRow的空行，一行一行的复制粘贴
                     //复制模板行
                        srcRange = operTable.Rows[2].Range;// 所以，粘贴几行，就要多加几行，j+j
                        srcRange.Copy();
                        //粘贴到目标行
                        desRange = operTable.Rows[cur_row + j+1].Range;//因为，新粘贴的行 在原来模板行的上面
                        desRange.Paste();
                        //operTable.Cell(nameCellRow[cur_row + j + 1], nameCellColumn[name_cell_idx]).Range.Text = str;
                        operTable.Cell(cur_row + j + 1, 1).Range.Text = _dt.Rows[j][0].ToString();
                        operTable.Cell(cur_row + j + 1, 2).Range.Text = _dt.Rows[j][1].ToString();
                        operTable.Cell(cur_row + j + 1, 3).Range.Text = _dt.Rows[j][2].ToString();
                        operTable.Cell(cur_row + j + 1, 4).Range.Text = _dt.Rows[j][3].ToString();
                        operTable.Cell(cur_row + j + 1, 5).Range.Text = _dt.Rows[j][4].ToString();
                        //string _curr = _dt.Rows[j][5].ToString().Substring(0, 3);
                        string _number = _dt.Rows[j][5].ToString();                        
                        decimal _Fmoney = Convert.ToDecimal(_number);
                        //MessageBox.Show(string.Format("{0:N}", _Fmoney));
                        operTable.Cell(cur_row + j + 1, 6).Range.Text = string.Format("{0:N}", _Fmoney);
                    }
                    cur_row += Trow;                   
                }
                catch (System.Exception ex)
                {
                    Utils.Writelogstream(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " "+ex.ToString());
                }
            }
        }
        public System.Data.DataTable GetDTbyCust()
        {
            string conn = "server = 192.168.19.3; database = AErp_Sinno; uid = Justin; pwd = 5t6y&U*I";
            System.Data.DataTable _dt = new System.Data.DataTable();
            _dt = Utils.executeQueryT("Pro_GetOpenSAInv", conn);
            return _dt;
        }
        
        public void RedueAutoReminder()
        {
            #region
            string conn = "server = 192.168.19.3; database = AErp_Sinno; uid = Justin; pwd = 5t6y&U*I";
            System.Data.DataTable _dt = new System.Data.DataTable();
            _dt = GetDTbyCust();
            if (_dt != null && _dt.Rows.Count > 0)
            {
                for (int ii = 0; ii < _dt.Rows.Count - 1; ii++)
                {
                    WriteIntoWord wiw = new WriteIntoWord();
                    string FilePath = System.Windows.Forms.Application.StartupPath + @"\OverduePaymentReminder.dot";     //模板路径 
                    string CrdName = "CrdName";
                    string _CrdName = _dt.Rows[ii][0].ToString();
                    string CurrID = "CurrID";
                    string _CurrID = _dt.Rows[ii][1].ToString();
                    string EmpName = "EmpName";
                    string _EmpName = _dt.Rows[ii][2].ToString();
                    string OverOpenTotal = "OverOpenTotal";
                    string _curr = _dt.Rows[ii][3].ToString().Substring(0, 3);
                    string _number = _dt.Rows[ii][3].ToString().Substring(3, _dt.Rows[ii][3].ToString().Length-3);
                    decimal _Fmoney = Convert.ToDecimal(_number);
                    string _OverOpenTotal = _curr+string.Format("{0:N}", _Fmoney);
                    string OverOpenTotal1 = "OverOpenTotal1";
                    string NowDate = "NowDate";
                    string _NowDate = _dt.Rows[ii][4].ToString();
                    string _NowDate1 = "NowDate1";
                    string DueDate = "DueDate";
                    string _DueDate = _dt.Rows[ii][5].ToString();
                    string CNDueDate = "CNDueDate";
                    string _CNDueDate = _dt.Rows[ii][6].ToString();
                    string _empid = _dt.Rows[ii][7].ToString();
                    string Mobile = "Mobile";
                    String Mobile1 = "Mobile1";
                    string _Mobile = _dt.Rows[ii][8].ToString();
                    string _SMail = _dt.Rows[ii][9].ToString();
                    string _path = System.Windows.Forms.Application.StartupPath + @"\Export\" + DateTime.Now.ToString("yyyyMMdd") + @"\";
                    if (!Directory.Exists(_path))
                    {
                        Directory.CreateDirectory(_path);
                    }
                    string SaveDocPath = _path + _CrdName + _CurrID + "_OverduePaymentReminder.pdf";
                    wiw.OpenDocument(FilePath);
                    //wiw.WriteIntoDocument(CrdName, _CrdName);
                    // wiw.WriteIntoDocument(CurrID, _CurrID);
                    wiw.WriteIntoDocument(EmpName, _EmpName);
                    wiw.WriteIntoDocument(OverOpenTotal, _OverOpenTotal);
                    wiw.WriteIntoDocument(OverOpenTotal1, _OverOpenTotal);
                    wiw.WriteIntoDocument(NowDate, _NowDate);
                    wiw.WriteIntoDocument(_NowDate1, _NowDate);
                    wiw.WriteIntoDocument(DueDate, _DueDate);
                    wiw.WriteIntoDocument(CNDueDate, _CNDueDate);
                    wiw.WriteIntoDocument(Mobile, _Mobile);
                    wiw.WriteIntoDocument(Mobile1, _Mobile);
                    System.Data.DataTable _dts = new System.Data.DataTable();
                    _dts = Utils.executeQueryT("Pro_GetOpenSAInvList '" + _CrdName + "','" + _CurrID + "','" + _empid + "'", conn);
                    if (_dts != null && _dts.Rows.Count > 0)
                    {
                        wiw.WriteIntoTable("Report__", _dts);
                    }
                    wiw.Save_CloseDocument(SaveDocPath);
                    Utils.sendMail("Finance", _SMail, "", SaveDocPath, _EmpName, _CrdName, _OverOpenTotal);
                }
            }
            #endregion
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            RedueAutoReminder();
            this.Close();
            this.Dispose();
        }
        
    }
}
