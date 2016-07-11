using System;
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

            if (cbKindOfSheet.SelectedIndex == -1)
                cbKindOfSheet.SelectedIndex = 0;
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

                    // 집계표 변환 작업
                    if (cbKindOfSheet.SelectedIndex == 0)
                    {
                        var _worker = new TransferSheetType01();

                        var _start_row_number = Convert.ToInt32(tbStartRowNumber.Text);
                        var _number_of_row_per_page = Convert.ToInt32(tbNumberOfRowPerPage.Text);
                        var _height_of_row = Convert.ToDouble(tbHeightOfRow.Text);

                        var _total_pages = await _worker.GetPageCount(_wb, _start_row_number, _number_of_row_per_page, __cancel_token_source.Token);

                        tsProgressBar.Maximum = _total_pages;
                        tsProgressBar.Step = 1;

                        var _progress = new Progress<string>(s =>
                        {
                            status_message.Text = s;

                            if (tsProgressBar.Value >= tsProgressBar.Maximum)
                                tsProgressBar.Value = tsProgressBar.Minimum;

                            tsProgressBar.PerformStep();
                        });

                        var _success = await _worker.DoTransfer(_wb, _total_pages, _start_row_number, _number_of_row_per_page, _height_of_row, _progress, __cancel_token_source.Token);
                        if (_success == true)
                        {
                            status_message.Text = "변환 된 엑셀파일을 저장 중 입니다.";
                            _wb.SaveAs(_after_name);
                        }
                        else
                            status_message.Text = "변환 작업 중 오류가 발생 하였습니다.";
                    }
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
                if (File.Exists(p_xlsx_name) == false)
                {
                    MessageBox.Show($"파일({Path.GetFileName(p_xlsx_name)})을 찾을 수 없습니다.");
                    return;
                }

                Thread.Sleep(1500);

                System.Diagnostics.Process.Start(p_xlsx_name);
            });
        }
    }
}