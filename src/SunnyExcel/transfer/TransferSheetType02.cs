using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace SunnyExcel
{
    public class TransferSheetType02
    {
        /// <summary>
        /// 진행율을 보여 주기 위해
        /// 전체 페이지 수가 몇개인지 변환 작업 전에 계산 합니다.
        /// </summary>
        /// <param name="p_before_book"></param>
        /// <param name="p_start_row_number"></param>
        /// <param name="p_number_of_row_per_page"></param>
        /// <param name="p_cancel_token"></param>
        /// <returns></returns>
        public async Task<int> GetPageCount(XLWorkbook p_before_book, int p_start_row_number, int p_number_of_row_per_page, CancellationToken p_cancel_token)
        {
            var _result = 0;

            await Task.Run(() =>
            {
                foreach (IXLWorksheet _ws in p_before_book.Worksheets)
                {
                    p_cancel_token.ThrowIfCancellationRequested();

                    var _last_row_number = _ws.LastRowUsed().RowNumber();
                    _result += (int)Decimal.Round((_last_row_number - (p_start_row_number - 1)) / p_number_of_row_per_page);
                }
            });

            return _result;
        }

        /// <summary>
        /// 전기공사 집계표 변환 작업을 실행 합니다.
        /// </summary>
        /// <param name="p_before_book"></param>
        /// <param name="p_total_pages"></param>
        /// <param name="p_start_row_number"></param>
        /// <param name="p_number_of_row_per_page"></param>
        /// <param name="p_height_of_row"></param>
        /// <param name="p_progress"></param>
        /// <param name="p_cancel_token"></param>
        /// <returns></returns>
        public async Task<bool> DoTransfer
            (
                XLWorkbook p_before_book, 
                int p_total_pages, int p_start_row_number, int p_number_of_row_per_page, double p_height_of_row, 
                IProgress<string> p_progress, CancellationToken p_cancel_token
            )
        {
            var _result = true;

            await Task.Run(() =>
            {
                var _page_offset = 1;

                foreach (IXLWorksheet _ws in p_before_book.Worksheets)
                {
                    var _last_row_number = _ws.LastRowUsed().RowNumber();
                    var _number_of_pages = (int)Decimal.Round((_last_row_number - (p_start_row_number - 1)) / p_number_of_row_per_page);

                    for (int _page_number = _number_of_pages; _page_number > 0; _page_number--, _page_offset++)
                    {
                        p_cancel_token.ThrowIfCancellationRequested();

                        p_progress.Report(String.Format("sheet-name: {0}, page:{1}/{2} 변환 중...", _ws.Name, _page_offset, p_total_pages));

                        var _page_first_row = (_page_number - 1) * p_number_of_row_per_page + p_start_row_number;
                        if (String.IsNullOrEmpty(_ws.Row(_page_first_row).Cell(1).GetValue<string>()))
                            continue;

                        var _row_position = _page_first_row + p_number_of_row_per_page - 1;

                        // 각각의 row 밑에 한줄을 삽입 합니다.
                        for (int i = 1; i <= p_number_of_row_per_page; i++)
                        {
                            _ws.Row(_row_position).InsertRowsBelow(1);
                            _row_position--;
                        }

                        _row_position = _page_first_row;

                        for (int i = 1; i <= p_number_of_row_per_page; i++)
                        {
                            // C 컬럼을 2줄씩 병합 합니다.
                            {
                                _ws.Range(String.Format("C{0}:C{1}", _row_position, _row_position + 1)).Column(1).Merge();
                            }

                            // 윗줄을 아랫줄로 동일하게 복사 합니다.
                            {
                                var _dcell = _ws.Row(_row_position).Cell("D");
                                var _ecell = _ws.Row(_row_position).Cell("E");
                                var _fcell = _ws.Row(_row_position).Cell("F");
                                var _gcell = _ws.Row(_row_position).Cell("G");

                                _row_position++;

                                _ws.Row(_row_position).Cell("D").Value = _dcell;
                                _ws.Row(_row_position).Cell("E").Value = _ecell;
                                _ws.Row(_row_position).Cell("F").Value = _fcell;
                                _ws.Row(_row_position).Cell("G").Value = _gcell;

                                _row_position++;
                            }
                        }

                        _row_position = _page_first_row;

                        // 첫 줄 병합
                        _ws.Range(String.Format("A{0}:A{1}", _row_position + 0, _row_position + 35)).Merge();
                        _ws.Range(String.Format("B{0}:B{1}", _row_position + 0, _row_position + 7)).Merge();
                        _ws.Range(String.Format("B{0}:B{1}", _row_position + 8, _row_position + 13)).Merge();
                        _ws.Range(String.Format("B{0}:B{1}", _row_position + 14, _row_position + 35)).Merge();

                        for (int r = 0; r < 6; r++)
                            _ws.Range(String.Format("A{0}:C{1}", _row_position + 36 + r * 2, _row_position + 37 + r * 2)).Merge();

                        _ws.Range(String.Format("D{0}:E{1}", _row_position + 37, _row_position + 37)).Merge();
                        _ws.Range(String.Format("D{0}:E{1}", _row_position + 43, _row_position + 43)).Merge();

                        // 첫줄의 글자색을 red-color로 변경 하고, 하단 테두리를 삭제 합니다.
                        for (int i = 1; i <= p_number_of_row_per_page; i++)
                        {
                            var _upper = _ws.Range(String.Format("D{0}:G{1}", _row_position, _row_position)).Cells();
                            _upper.Style.Font.SetFontColor(XLColor.Red);

                            var _cell4 = _ws.Cell(_row_position, 4);
                            _cell4.RichText.SetFontColor(XLColor.Red);

                            _row_position++;

                            var _bottom = _ws.Range(String.Format("D{0}:G{1}", _row_position, _row_position)).Cells();
                            _upper.Style.Border.SetBottomBorder(XLBorderStyleValues.None);
                            _bottom.Style.Border.SetTopBorder(XLBorderStyleValues.None);

                            _row_position++;
                        }

                        // 각줄의 높이를 동일하게 변경 합니다.
                        var _page_last_row = _page_first_row + p_number_of_row_per_page * 2 - 1;
                        _ws.Rows(_page_first_row, _page_last_row).Height = p_height_of_row;
                    }
                }
            },
            p_cancel_token);

            return _result;
        }
    }
}