using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using MOIE = Microsoft.Office.Interop.Excel;

namespace SunnyExcel
{
    public partial class MainForm : Form
    {
        private static CLayout __ilayout = null;
        public static CLayout __clayout
        {
            get
            {
                if (__ilayout == null)
                    __ilayout = new CLayout();
                return __ilayout;
            }
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void miExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            __clayout.SaveFormLayout(this);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            __clayout.RestoreFormLayout(this);
        }

        private void miAbout_Click(object sender, EventArgs e)
        {
            var _dialog = new AboutDialog(this);
            _dialog.Show();
        }

        public void RestoreFormLayout(Form p_child_form)
        {
            __clayout.RestoreFormLayout(p_child_form, this.GetType().GUID.ToString());
        }

        public void SaveFormLayout(Form p_child_form)
        {
            __clayout.SaveFormLayout(p_child_form, this.GetType().GUID.ToString());
        }

        private void btBefore_Click(object sender, EventArgs e)
        {
            btBefore.Enabled = btAfter.Enabled = false;

            try
            {
                status_message.Text = "변환 할 엑셀 파일을 선택 하세요.";

                if (ofBeforeExcel.ShowDialog() == DialogResult.OK)
                {
                    status_message.Text = "변환 할 엑셀 파일이 선택 되었습니다.";

                    var _before_name = tbBefore.Text = ofBeforeExcel.FileName;
                    tbAfter.Text = "";

                    var _extension = Path.GetExtension(_before_name).ToLower();
                    if (_extension != ".xls" && _extension != ".xlsx")
                    {
                        status_message.Text = String.Format("{0}는 엑셀파일이 아닙니다.", Path.GetFileName(_before_name));
                        tbBefore.Text = tbAfter.Text = "";

                        return;
                    }

                    if (_extension.ToLower() == ".xls")
                        _before_name = ConvertXlsToXlsx(_before_name);

                    var _wb = new XLWorkbook(_before_name);
                    {
                        tbAuthor.Text = _wb.Author;

                        tbModifier.Text = _wb.Properties.LastModifiedBy;
                        tbCreated.Text = _wb.Properties.Created.ToShortDateString();
                    }

                    tbBefore.Text = _before_name;
                    tbAfter.Text = "";

                    status_message.Text = "변환 할 엑셀 파일이 정상적으로 선택 되었습니다.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                status_message.Text = ex.Message;
                tbBefore.Text = tbAfter.Text = "";
            }
            finally
            {
                btBefore.Enabled = btAfter.Enabled = true;
            }
        }

        private static CancellationTokenSource __cancel_token_source = null;

        private async void btAfter_Click(object sender, EventArgs e)
        {
            if (__cancel_token_source != null)
            {
                btAfter.Enabled = false;
                __cancel_token_source.Cancel();

                Thread.Sleep(1000);
                return;
            }

            btBefore.Enabled = false;
            btAfter.Text = "변환중지";

            var _save_max = tsProgressBar.Maximum;
            var _save_step = tsProgressBar.Step;

            try
            {
                tsProgressBar.Value = tsProgressBar.Minimum;
                status_message.Text = "파일 변환 작업이 시작 되었습니다.";

                var _before_name = tbBefore.Text;

                var _extension = Path.GetExtension(_before_name).ToLower();
                if (_extension != ".xls" && _extension != ".xlsx")
                {
                    status_message.Text = "엑셀파일을 먼저 선택 해 주세요.";
                    tbBefore.Text = tbAfter.Text = "";

                    return;
                }

                var _after_name = Path.Combine
                            (
                                Path.GetDirectoryName(_before_name),
                                Path.GetFileNameWithoutExtension(_before_name)
                                + "_result"
                                + Path.GetExtension(_before_name)
                            );

                if (File.Exists(_after_name) == true)
                {
                    var _action = MessageBox.Show("변환 된 동일한 파일이 존재 합니다. 덮어쓰기를 하시겠습니까?", "작업선택", MessageBoxButtons.YesNo);
                    if (_action == DialogResult.No)
                    {
                        status_message.Text = "작업을 취소 했습니다.";
                        return;
                    }

                    File.Delete(_after_name);
                }

                status_message.Text = "엑셀파일을 변환 합니다.";

                tbAfter.Text = _after_name;

                var _wb = new XLWorkbook(_before_name);
                {
                    __cancel_token_source = new CancellationTokenSource();

                    var _total_pages = await CalculateTotalPages(_wb, __cancel_token_source.Token);

                    tsProgressBar.Maximum = _total_pages;
                    tsProgressBar.Step = 1;

                    var _progress = new Progress<string>(s =>
                    {
                        status_message.Text = s;

                        if (tsProgressBar.Value >= tsProgressBar.Maximum)
                            tsProgressBar.Value = tsProgressBar.Minimum;

                        tsProgressBar.PerformStep();

                    });

                    var _success = await TransferOldToNew(_wb, _total_pages, _progress, __cancel_token_source.Token);
                    if (_success == true)
                    {
                        status_message.Text = "변환 된 엑셀파일을 저장 중 입니다.";
                        _wb.SaveAs(_after_name);
                    }
                    else
                        status_message.Text = "변환 작업 중 오류가 발생 하였습니다.";
                }

                if (cbAfterOpen.Checked == true)
                {
                    status_message.Text = "변환 된 엑셀 파일을 실행(EXEC) 합니다.";

                    await OpenExcelFile(_after_name);
                }
                else
                {
                    status_message.Text = "변환 후 엑셀파일을 저장 하였습니다.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                status_message.Text = ex.Message;
            }
            finally
            {
                await ClearStatusMessage();

                btAfter.Text = "변환시작";
                btBefore.Enabled = btAfter.Enabled = true;

                if (__cancel_token_source != null)
                {
                    __cancel_token_source.Dispose();
                    __cancel_token_source = null;
                }

                tsProgressBar.Maximum = _save_max;
                tsProgressBar.Step = _save_step;
            }
        }

        private string ConvertXlsToXlsx(string p_file_name)
        {
            status_message.Text = String.Format("{0} 파일을 xlsx 파일로 업그레이드 중 입니다.", Path.GetFileName(p_file_name));

            var _xlsx_name = Path.ChangeExtension(p_file_name, ".xlsx");
            if (File.Exists(_xlsx_name))
                File.Delete(_xlsx_name);

            // Interop.Excel
            {
                MOIE.Application _xls_app = new MOIE.Application();
                _xls_app.Visible = false;

                MOIE.Workbook _workbook = _xls_app.Workbooks.Open(p_file_name);
                _workbook.SaveAs(_xlsx_name, MOIE.XlFileFormat.xlOpenXMLWorkbook, AccessMode: MOIE.XlSaveAsAccessMode.xlNoChange);

                _workbook.Close(false, Type.Missing, Type.Missing);
            }

            return _xlsx_name;
        }

        private async Task<int> CalculateTotalPages(XLWorkbook p_before_book, CancellationToken cancellationToken)
        {
            var _result = 0;

            await Task.Run(() =>
            {
                var _start_row_number = Convert.ToInt32(tbStartRowNumber.Text);
                var _number_of_row_per_page = Convert.ToInt32(tbNumberOfRowPerPage.Text);

                foreach (IXLWorksheet _ws in p_before_book.Worksheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var _last_row_number = _ws.LastRowUsed().RowNumber();
                    _result += (int)Decimal.Round((_last_row_number - (_start_row_number - 1)) / _number_of_row_per_page);
                }
            });

            return _result;
        }

        private async Task<bool> TransferOldToNew(XLWorkbook p_before_book, int p_total_pages, IProgress<string> p_progress, CancellationToken p_cance_token)
        {
            var _result = true;

            await Task.Run(() =>
            {
                var _start_row_number = Convert.ToInt32(tbStartRowNumber.Text);
                var _number_of_row_per_page = Convert.ToInt32(tbNumberOfRowPerPage.Text);
                var _height_of_row = Convert.ToDouble(tbHeightOfRow.Text);

                var _page_offset = 1;

                foreach (IXLWorksheet _ws in p_before_book.Worksheets)
                {
                    var _last_row_number = _ws.LastRowUsed().RowNumber();
                    var _number_of_pages = (int)Decimal.Round((_last_row_number - (_start_row_number - 1)) / _number_of_row_per_page);

                    for (int _page_number = _number_of_pages; _page_number > 0; _page_number--, _page_offset++)
                    {
                        p_cance_token.ThrowIfCancellationRequested();

                        p_progress.Report(String.Format("sheet-name: {0}, page:{1}/{2} 변환 중...", _ws.Name, _page_offset, p_total_pages));

                        var _page_first_row = (_page_number - 1) * _number_of_row_per_page + _start_row_number;
                        if (String.IsNullOrEmpty(_ws.Row(_page_first_row).Cell(1).GetValue<string>()))
                            continue;

                        var _row_position = _page_first_row + _number_of_row_per_page - 1;

                        for (int i = 1; i <= _number_of_row_per_page; i++)
                        {
                            _ws.Row(_row_position).InsertRowsBelow(1);
                            _row_position--;
                        }

                        _row_position = _page_first_row;

                        for (int i = 1; i <= _number_of_row_per_page; i++)
                        {
                            _ws.Range(String.Format("A{0}:A{1}", _row_position, _row_position + 1)).Column(1).Merge();
                            _ws.Range(String.Format("B{0}:B{1}", _row_position, _row_position + 1)).Column(1).Merge();
                            _ws.Range(String.Format("C{0}:C{1}", _row_position, _row_position + 1)).Column(1).Merge();
                            _ws.Range(String.Format("M{0}:M{1}", _row_position, _row_position + 1)).Column(1).Merge();

                            var _dcell = _ws.Row(_row_position).Cell("D");
                            var _ecell = _ws.Row(_row_position).Cell("E");
                            var _fcell = _ws.Row(_row_position).Cell("F");
                            var _gcell = _ws.Row(_row_position).Cell("G");
                            var _hcell = _ws.Row(_row_position).Cell("H");
                            var _icell = _ws.Row(_row_position).Cell("I");
                            var _jcell = _ws.Row(_row_position).Cell("J");
                            var _kcell = _ws.Row(_row_position).Cell("K");
                            var _lcell = _ws.Row(_row_position).Cell("L");

                            _row_position++;

                            _ws.Row(_row_position).Cell("D").Value = _dcell;
                            _ws.Row(_row_position).Cell("E").Value = _ecell;
                            _ws.Row(_row_position).Cell("F").Value = _fcell;
                            _ws.Row(_row_position).Cell("G").Value = _gcell;
                            _ws.Row(_row_position).Cell("H").Value = _hcell;
                            _ws.Row(_row_position).Cell("I").Value = _icell;
                            _ws.Row(_row_position).Cell("J").Value = _jcell;
                            _ws.Row(_row_position).Cell("K").Value = _kcell;
                            _ws.Row(_row_position).Cell("L").Value = _lcell;

                            _row_position++;
                        }

                        _row_position = _page_first_row;

                        for (int i = 1; i <= _number_of_row_per_page; i++)
                        {
                            var _upper = _ws.Range(String.Format("D{0}:L{1}", _row_position, _row_position)).Cells();
                            _upper.Style.Font.SetFontColor(XLColor.Red);

                            _row_position++;

                            var _bottom = _ws.Range(String.Format("D{0}:L{1}", _row_position, _row_position)).Cells();
                            _upper.Style.Border.SetBottomBorder(XLBorderStyleValues.None);
                            _bottom.Style.Border.SetTopBorder(XLBorderStyleValues.None);

                            _row_position++;
                        }

                        var _page_last_row = _page_first_row + _number_of_row_per_page * 2 - 1;
                        _ws.Rows(_page_first_row, _page_last_row).Height = _height_of_row;
                    }
                }
            },            
            p_cance_token);

            return _result;
        }

        private async Task ClearStatusMessage()
        {
            await Task.Run(() =>
            {
                this.InvokeEx(f =>
                {
                    tsProgressBar.Value = tsProgressBar.Maximum;
                });

                Thread.Sleep(3000);

                this.InvokeEx(f =>
                {
                    status_message.Text = "Ready";
                    tsProgressBar.Value = tsProgressBar.Minimum;
                });
            });
        }

        private async Task OpenExcelFile(string p_xlsx_name)
        {
            await Task.Run(() =>
            {
                Thread.Sleep(1500);

                System.Diagnostics.Process.Start(p_xlsx_name);
            });
        }
    }
}