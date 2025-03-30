using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using Shapes = Microsoft.Office.Interop.Excel.Shapes;

namespace ExcelAddIn
{
    public class MyLookupFunction
    {
        #region セル検索

        /// <summary>
        /// 複数行・列の範囲でRegexに合致した行の範囲を返す。
        /// </summary>
        /// <param name="searchPattern"></param>
        /// <param name="searchRange"></param>
        /// <param name="searchColIndex"></param>
        /// <param name="startColIndex"></param>
        /// <param name="endColIndex"></param>
        /// <param name="removeSpaces"></param>
        /// <returns></returns>
        [ExcelFunction(Name = "RxMultiLOOKUP",
         Description = "正規表現を使用して検索し、複数行/列の範囲を抽出します。",
         Category = "検索/参照")]
        public static object[,] RxMultiLOOKUP(
            [ExcelArgument(Name = "検索値", Description = "検索する正規表現パターン")] object searchPattern,
            [ExcelArgument(Name = "検索範囲", Description = "パターンを検索する範囲")] object[,] searchRange,
            [ExcelArgument(Name = "検索列", Description = "検索範囲の何列目を検索対象とするか (0から始まる)")] int searchColIndex = 0,
            [ExcelArgument(Name = "開始列", Description = "検索範囲の何列目から抽出するか (0から始まる)")] int startColIndex = 0,
            [ExcelArgument(Name = "終了列", Description = "検索範囲の何列目まで抽出するか (0から始まる)")] int endColIndex = 0,
            [ExcelArgument(Name = "スペース除去", Description = "検索時に半角/全角スペースを除去するか")] bool removeSpaces = false)
        {
            try
            {
                // 検索パターンを文字列に変換
                string searchPatternStr = searchPattern?.ToString() ?? "";

                // 検索する値がない場合はエラー
                if ( string.IsNullOrEmpty(searchPatternStr) )
                {
                    return new object[1, 1] { { "#VALUE!" } }; // ExcelErrorを避ける
                }

                // 範囲の行数と列数を取得
                int rows = searchRange.GetLength(0);
                int cols = searchRange.GetLength(1);

                // 検索列インデックスが範囲外の場合はエラー
                if ( searchColIndex < 0 || searchColIndex >= cols )
                {
                    return new object[1, 1] { { $"#エラー: 検索列インデックスが範囲外です (0-{cols - 1})" } };
                }

                // 終了列が指定されていない場合は、検索範囲の最右列を使用
                if ( endColIndex == 0 && startColIndex == 0 )
                {
                    endColIndex = cols - 1;
                }
                else if ( endColIndex < startColIndex )
                {
                    // 終了列が開始列より小さい場合は開始列に合わせる
                    endColIndex = startColIndex;
                }

                // インデックスが範囲外にならないよう調整
                startColIndex = Math.Max(0, Math.Min(cols - 1, startColIndex));
                endColIndex = Math.Max(0, Math.Min(cols - 1, endColIndex));

                // 正規表現オブジェクトを作成
                Regex regex;
                try
                {
                    regex = new Regex(searchPatternStr);
                }
                catch ( Exception ex )
                {
                    return new object[1, 1] { { $"#正規表現エラー: {ex.Message}" } };
                }

                // 指定された列で正規表現に一致する行を検索
                List<int> matchingRows = new List<int>();
                for ( int row = 0; row < rows; row++ )
                {
                    object cellValue = searchRange[row, searchColIndex];
                    string cellText = cellValue?.ToString() ?? "";

                    // スペース除去オプションがTRUEの場合
                    if ( removeSpaces )
                    {
                        // 半角スペースと全角スペースを除去
                        cellText = cellText.Replace(" ", "").Replace("　", "");
                    }

                    if ( regex.IsMatch(cellText) )
                    {
                        matchingRows.Add(row);
                    }
                }

                // 一致する行がない場合はN/Aエラーを返す
                if ( matchingRows.Count == 0 )
                {
                    return new object[1, 1] { { "#N/A" } }; // ExcelErrorを避ける
                }

                // 抽出する列数を計算
                int resultColumnsCount = endColIndex - startColIndex + 1;

                // 結果配列を作成
                object[,] result = new object[matchingRows.Count, resultColumnsCount];

                // 一致した行ごとに結果を設定
                for ( int i = 0; i < matchingRows.Count; i++ )
                {
                    int matchedRow = matchingRows[i];

                    // 指定された列範囲のデータを取得
                    for ( int j = 0; j < resultColumnsCount; j++ )
                    {
                        int sourceColIndex = startColIndex + j;
                        result[i, j] = searchRange[matchedRow, sourceColIndex];
                    }
                }

                // 常に配列として返す - 単一セルの場合の特別処理を削除
                return result;
            }
            catch ( Exception ex )
            {
                return new object[1, 1] { { $"#エラー: {ex.Message}" } };
            }
        }

        /// <summary>
        /// 一行だけ結果を返すタイプのRegex数式。
        /// </summary>
        /// <param name="searchPattern"></param>
        /// <param name="searchRange"></param>
        /// <param name="searchColIndex"></param>
        /// <param name="startColIndex"></param>
        /// <param name="endColIndex"></param>
        /// <param name="removeSpaces"></param>
        /// <returns></returns>
        [ExcelFunction(Name = "RxLOOKUP",
 Description = "正規表現を使用して検索し、単一行/複数列の範囲を抽出します。",
 Category = "検索/参照")]
        public static object[,] RxLOOKUP(
            [ExcelArgument(Name = "検索値", Description = "検索する正規表現パターン")] object searchPattern,
            [ExcelArgument(Name = "検索範囲", Description = "パターンを検索する範囲")] object[,] searchRange,
            [ExcelArgument(Name = "検索列", Description = "検索範囲の何列目を検索対象とするか (0から始まる)")] int searchColIndex = 0,
            [ExcelArgument(Name = "開始列", Description = "検索範囲の何列目から抽出するか (0から始まる)")] int startColIndex = 0,
            [ExcelArgument(Name = "終了列", Description = "検索範囲の何列目まで抽出するか (0から始まる)")] int endColIndex = 0,
            [ExcelArgument(Name = "スペース除去", Description = "検索時に半角/全角スペースを除去するか")] bool removeSpaces = false)
        {
            try
            {
                // 検索パターンを文字列に変換
                string searchPatternStr = searchPattern?.ToString() ?? "";

                // 検索する値がない場合はエラー
                if ( string.IsNullOrEmpty(searchPatternStr) )
                {
                    return new object[1, 1] { { "#VALUE!" } }; // ExcelErrorを避ける
                }

                // 範囲の行数と列数を取得
                int rows = searchRange.GetLength(0);
                int cols = searchRange.GetLength(1);

                // 検索列インデックスが範囲外の場合はエラー
                if ( searchColIndex < 0 || searchColIndex >= cols )
                {
                    return new object[1, 1] { { $"#エラー: 検索列インデックスが範囲外です (0-{cols - 1})" } };
                }

                // 終了列が指定されていない場合は、検索範囲の最右列を使用
                if ( endColIndex == 0 && startColIndex == 0 )
                {
                    endColIndex = cols - 1;
                }
                else if ( endColIndex < startColIndex )
                {
                    // 終了列が開始列より小さい場合は開始列に合わせる
                    endColIndex = startColIndex;
                }

                // インデックスが範囲外にならないよう調整
                startColIndex = Math.Max(0, Math.Min(cols - 1, startColIndex));
                endColIndex = Math.Max(0, Math.Min(cols - 1, endColIndex));

                // 正規表現オブジェクトを作成
                Regex regex;
                try
                {
                    regex = new Regex(searchPatternStr);
                }
                catch ( Exception ex )
                {
                    return new object[1, 1] { { $"#正規表現エラー: {ex.Message}" } };
                }

                // 指定された列で正規表現に一致する行を検索
                int matchingRow = -1;
                for ( int row = 0; row < rows; row++ )
                {
                    object cellValue = searchRange[row, searchColIndex];
                    string cellText = cellValue?.ToString() ?? "";

                    // スペース除去オプションがTRUEの場合
                    if ( removeSpaces )
                    {
                        // 半角スペースと全角スペースを除去
                        cellText = cellText.Replace(" ", "").Replace("　", "");
                    }

                    if ( regex.IsMatch(cellText) )
                    {
                        matchingRow = row;
                        break;
                    }
                }

                // 一致する行がない場合はN/Aエラーを返す
                if ( matchingRow == -1 )
                {
                    return new object[1, 1] { { "#N/A" } }; // ExcelErrorを避ける
                }

                // 抽出する列数を計算
                int resultColumnsCount = endColIndex - startColIndex + 1;

                // 結果配列を作成
                object[,] result = new object[1, resultColumnsCount];

                // 指定された列範囲のデータを取得
                for ( int j = 0; j < resultColumnsCount; j++ )
                {
                    int sourceColIndex = startColIndex + j;
                    result[0, j] = searchRange[matchingRow, sourceColIndex];
                }

                // 常に配列として返す - 単一セルの場合の特別処理を削除
                return result;
            }
            catch ( Exception ex )
            {
                return new object[1, 1] { { $"#エラー: {ex.Message}" } };
            }
        }

        /// <summary>
        /// 複数行・列の範囲で検索文字列に合致した行の範囲を返す。
        /// </summary>
        /// <param name="searchPattern"></param>
        /// <param name="searchRange"></param>
        /// <param name="searchColIndex"></param>
        /// <param name="startColIndex"></param>
        /// <param name="endColIndex"></param>
        /// <param name="removeSpaces"></param>
        /// <returns></returns>
        [ExcelFunction(Name = "ExMultiLOOKUP",
         Description = "通常文字列を使用して検索し、複数行/列の範囲を抽出します。",
         Category = "検索/参照")]
        public static object[,] ExMultiLOOKUP(
            [ExcelArgument(Name = "検索値", Description = "検索する正規表現パターン")] object searchPattern,
            [ExcelArgument(Name = "検索範囲", Description = "パターンを検索する範囲")] object[,] searchRange,
            [ExcelArgument(Name = "検索列", Description = "検索範囲の何列目を検索対象とするか (0から始まる)")] int searchColIndex = 0,
            [ExcelArgument(Name = "開始列", Description = "検索範囲の何列目から抽出するか (0から始まる)")] int startColIndex = 0,
            [ExcelArgument(Name = "終了列", Description = "検索範囲の何列目まで抽出するか (0から始まる)")] int endColIndex = 0,
            [ExcelArgument(Name = "スペース除去", Description = "検索時に半角/全角スペースを除去するか")] bool removeSpaces = false,
            [ExcelArgument(Name = "完全一致", Description = "完全一致のみ取得するか")] bool fullMatch = false)
        {
            try
            {
                // 検索パターンを文字列に変換
                string searchPatternStr = searchPattern?.ToString() ?? "";

                // 検索する値がない場合はエラー
                if ( string.IsNullOrEmpty(searchPatternStr) )
                {
                    return new object[1, 1] { { "#VALUE!" } }; // ExcelErrorを避ける
                }

                // 範囲の行数と列数を取得
                int rows = searchRange.GetLength(0);
                int cols = searchRange.GetLength(1);

                // 検索列インデックスが範囲外の場合はエラー
                if ( searchColIndex < 0 || searchColIndex >= cols )
                {
                    return new object[1, 1] { { $"#エラー: 検索列インデックスが範囲外です (0-{cols - 1})" } };
                }

                // 終了列が指定されていない場合は、検索範囲の最右列を使用
                if ( endColIndex == 0 && startColIndex == 0 )
                {
                    endColIndex = cols - 1;
                }
                else if ( endColIndex < startColIndex )
                {
                    // 終了列が開始列より小さい場合は開始列に合わせる
                    endColIndex = startColIndex;
                }

                // インデックスが範囲外にならないよう調整
                startColIndex = Math.Max(0, Math.Min(cols - 1, startColIndex));
                endColIndex = Math.Max(0, Math.Min(cols - 1, endColIndex));

                // 指定された列で正規表現に一致する行を検索
                List<int> matchingRows = new List<int>();
                for ( int row = 0; row < rows; row++ )
                {
                    object cellValue = searchRange[row, searchColIndex];
                    string cellText = cellValue?.ToString() ?? "";

                    // スペース除去オプションがTRUEの場合
                    if ( removeSpaces )
                    {
                        // 半角スペースと全角スペースを除去
                        cellText = cellText.Replace(" ", "").Replace("　", "");
                    }

                    // 検索文字列を含むかの判断。
                    // 完全マッチするかで判断方法を変える。
                    if ( fullMatch )
                    {
                        if ( searchPatternStr == cellText )
                        {
                            matchingRows.Add(row);
                        }
                    }
                    else
                    {
                        if ( cellText.Contains(searchPatternStr) )
                        {
                            matchingRows.Add(row);
                        }
                    }
                }

                // 一致する行がない場合はN/Aエラーを返す
                if ( matchingRows.Count == 0 )
                {
                    return new object[1, 1] { { "#N/A" } }; // ExcelErrorを避ける
                }

                // 抽出する列数を計算
                int resultColumnsCount = endColIndex - startColIndex + 1;

                // 結果配列を作成
                object[,] result = new object[matchingRows.Count, resultColumnsCount];

                // 一致した行ごとに結果を設定
                for ( int i = 0; i < matchingRows.Count; i++ )
                {
                    int matchedRow = matchingRows[i];

                    // 指定された列範囲のデータを取得
                    for ( int j = 0; j < resultColumnsCount; j++ )
                    {
                        int sourceColIndex = startColIndex + j;
                        result[i, j] = searchRange[matchedRow, sourceColIndex];
                    }
                }

                // 常に配列として返す - 単一セルの場合の特別処理を削除
                return result;
            }
            catch ( Exception ex )
            {
                return new object[1, 1] { { $"#エラー: {ex.Message}" } };
            }
        }

        /// <summary>
        /// 一行だけ結果を返すタイプのLookUp数式。
        /// </summary>
        /// <param name="searchPattern"></param>
        /// <param name="searchRange"></param>
        /// <param name="searchColIndex"></param>
        /// <param name="startColIndex"></param>
        /// <param name="endColIndex"></param>
        /// <param name="removeSpaces"></param>
        /// <returns></returns>
        [ExcelFunction(Name = "ExLOOKUP",
 Description = "通常文字列を使用して検索し、単一行/複数列の範囲を抽出します。",
 Category = "検索/参照")]
        public static object[,] ExLOOKUP(
            [ExcelArgument(Name = "検索値", Description = "検索する正規表現パターン")] object searchPattern,
            [ExcelArgument(Name = "検索範囲", Description = "パターンを検索する範囲")] object[,] searchRange,
            [ExcelArgument(Name = "検索列", Description = "検索範囲の何列目を検索対象とするか (0から始まる)")] int searchColIndex = 0,
            [ExcelArgument(Name = "開始列", Description = "検索範囲の何列目から抽出するか (0から始まる)")] int startColIndex = 0,
            [ExcelArgument(Name = "終了列", Description = "検索範囲の何列目まで抽出するか (0から始まる)")] int endColIndex = 0,
            [ExcelArgument(Name = "スペース除去", Description = "検索時に半角/全角スペースを除去するか")] bool removeSpaces = false,
            [ExcelArgument(Name = "完全一致", Description = "完全一致のみ取得するか")] bool fullMatch = false)
        {
            try
            {
                // 検索パターンを文字列に変換
                string searchPatternStr = searchPattern?.ToString() ?? "";

                // 検索する値がない場合はエラー
                if ( string.IsNullOrEmpty(searchPatternStr) )
                {
                    return new object[1, 1] { { "#VALUE!" } }; // ExcelErrorを避ける
                }

                // 範囲の行数と列数を取得
                int rows = searchRange.GetLength(0);
                int cols = searchRange.GetLength(1);

                // 検索列インデックスが範囲外の場合はエラー
                if ( searchColIndex < 0 || searchColIndex >= cols )
                {
                    return new object[1, 1] { { $"#エラー: 検索列インデックスが範囲外です (0-{cols - 1})" } };
                }

                // 終了列が指定されていない場合は、検索範囲の最右列を使用
                if ( endColIndex == 0 && startColIndex == 0 )
                {
                    endColIndex = cols - 1;
                }
                else if ( endColIndex < startColIndex )
                {
                    // 終了列が開始列より小さい場合は開始列に合わせる
                    endColIndex = startColIndex;
                }

                // インデックスが範囲外にならないよう調整
                startColIndex = Math.Max(0, Math.Min(cols - 1, startColIndex));
                endColIndex = Math.Max(0, Math.Min(cols - 1, endColIndex));

                // 指定された列で正規表現に一致する行を検索
                int matchingRow = -1;
                for ( int row = 0; row < rows; row++ )
                {
                    object cellValue = searchRange[row, searchColIndex];
                    string cellText = cellValue?.ToString() ?? "";

                    // スペース除去オプションがTRUEの場合
                    if ( removeSpaces )
                    {
                        // 半角スペースと全角スペースを除去
                        cellText = cellText.Replace(" ", "").Replace("　", "");
                    }

                    // 完全マッチするかで判断方法を変える。
                    if ( fullMatch )
                    {
                        if ( searchPatternStr == cellText )
                        {
                            matchingRow = row;
                            break;
                        }
                    }
                    else
                    {
                        if ( cellText.Contains(searchPatternStr) )
                        {
                            matchingRow = row;
                            break;
                        }
                    }
                }

                // 一致する行がない場合はN/Aエラーを返す
                if ( matchingRow == -1 )
                {
                    return new object[1, 1] { { "#N/A" } }; // ExcelErrorを避ける
                }

                // 抽出する列数を計算
                int resultColumnsCount = endColIndex - startColIndex + 1;

                // 結果配列を作成
                object[,] result = new object[1, resultColumnsCount];

                // 指定された列範囲のデータを取得
                for ( int j = 0; j < resultColumnsCount; j++ )
                {
                    int sourceColIndex = startColIndex + j;
                    result[0, j] = searchRange[matchingRow, sourceColIndex];
                }

                // 常に配列として返す - 単一セルの場合の特別処理を削除
                return result;
            }
            catch ( Exception ex )
            {
                return new object[1, 1] { { $"#エラー: {ex.Message}" } };
            }
        }

        #endregion

        [ExcelFunction(
            Name = "TEXTBOX_SEARCH",
            Description = "シート上のテキストボックスの内容を検索し、一致する行を返します。",
            Category = "テキスト",
            IsVolatile = true)]
        public static object[,] TextBoxSearch(
            [ExcelArgument(Name = "テキストボックス名", Description = "検索対象のテキストボックスの名前")] string textBoxName,
            [ExcelArgument(Name = "検索文字列", Description = "検索する文字列または正規表現パターン")] string searchPattern,
            [ExcelArgument(Name = "正規表現", Description = "TRUE=正規表現として検索、FALSE=通常の文字列として検索")] bool isRegex)
        {
            try
            {
                // パラメータの検証
                if ( string.IsNullOrEmpty(textBoxName) )
                {
                    return new object[1, 1] { { "#エラー: テキストボックス名が指定されていません" } };
                }

                if ( string.IsNullOrEmpty(searchPattern) )
                {
                    return new object[1, 1] { { "#エラー: 検索パターンが指定されていません" } };
                }

                // テキストボックスのテキストを取得
                string textBoxContent = GetTextBoxContent(textBoxName);

                if ( string.IsNullOrEmpty(textBoxContent) )
                {
                    return new object[1, 1] { { $"#エラー: テキストボックス '{textBoxName}' が見つからないか、テキストがありません" } };
                }

                // テキストを行ごとに分割
                string[] lines = textBoxContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                // 検索パターンに一致する行を検索
                List<string> matchedLines = new List<string>();

                if ( isRegex )
                {
                    // 正規表現として検索
                    Regex regex;
                    try
                    {
                        regex = new Regex(searchPattern);
                    }
                    catch ( Exception ex )
                    {
                        return new object[1, 1] { { $"#正規表現エラー: {ex.Message}" } };
                    }

                    foreach ( string line in lines )
                    {
                        if ( regex.IsMatch(line) )
                        {
                            matchedLines.Add(line);
                        }
                    }
                }
                else
                {
                    // 通常の文字列として検索
                    foreach ( string line in lines )
                    {
                        if ( line.Contains(searchPattern) )
                        {
                            matchedLines.Add(line);
                        }
                    }
                }

                // 一致する行がない場合
                if ( matchedLines.Count == 0 )
                {
                    return new object[1, 1] { { "#N/A" } };
                }

                // 結果を配列として返す
                object[,] result = new object[matchedLines.Count, 1];
                for ( int i = 0; i < matchedLines.Count; i++ )
                {
                    result[i, 0] = matchedLines[i];
                }

                return result;
            }
            catch ( Exception ex )
            {
                return new object[1, 1] { { $"#エラー: {ex.Message}" } };
            }
            finally
            {
                // COMオブジェクトのクリーンアップ
                Marshal.CleanupUnusedObjectsInCurrentContext();
            }
        }

        [ExcelFunction(
            Name = "TEXTBOX_COUNT",
            Description = "シート上のテキストボックスの内容を検索し、一致する行数を返します。",
            Category = "テキスト",
            IsVolatile = true)]
        public static object TextBoxCount(
            [ExcelArgument(Name = "テキストボックス名", Description = "検索対象のテキストボックスの名前")] string textBoxName,
            [ExcelArgument(Name = "検索文字列", Description = "検索する文字列または正規表現パターン")] string searchPattern,
            [ExcelArgument(Name = "正規表現", Description = "TRUE=正規表現として検索、FALSE=通常の文字列として検索")] bool isRegex)
        {
            try
            {
                // パラメータの検証
                if ( string.IsNullOrEmpty(textBoxName) )
                {
                    return "#エラー: テキストボックス名が指定されていません";
                }

                if ( string.IsNullOrEmpty(searchPattern) )
                {
                    return "#エラー: 検索パターンが指定されていません";
                }

                // テキストボックスのテキストを取得
                string textBoxContent = GetTextBoxContent(textBoxName);

                if ( string.IsNullOrEmpty(textBoxContent) )
                {
                    return $"#エラー: テキストボックス '{textBoxName}' が見つからないか、テキストがありません";
                }

                // テキストを行ごとに分割
                string[] lines = textBoxContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                // 検索パターンに一致する行をカウント
                int count = 0;

                if ( isRegex )
                {
                    // 正規表現として検索
                    Regex regex;
                    try
                    {
                        regex = new Regex(searchPattern);
                    }
                    catch ( Exception ex )
                    {
                        return $"#正規表現エラー: {ex.Message}";
                    }

                    foreach ( string line in lines )
                    {
                        if ( regex.IsMatch(line) )
                        {
                            count++;
                        }
                    }
                }
                else
                {
                    // 通常の文字列として検索
                    foreach ( string line in lines )
                    {
                        if ( line.Contains(searchPattern) )
                        {
                            count++;
                        }
                    }
                }

                return count;
            }
            catch ( Exception ex )
            {
                return $"#エラー: {ex.Message}";
            }
            finally
            {
                // COMオブジェクトのクリーンアップ
                Marshal.CleanupUnusedObjectsInCurrentContext();
            }
        }

        /// <summary>
        /// テキストボックスのコンテンツを取得する
        /// </summary>
        /// <param name="textBoxName">テキストボックスの名前</param>
        /// <returns>テキストボックスの内容、見つからない場合は空文字列</returns>
        private static string GetTextBoxContent(string textBoxName)
        {
            if ( string.IsNullOrEmpty(textBoxName) )
            {
                return "";
            }

            try
            {
                // Excelアプリケーションの取得
                Worksheet activeSheet;

                try
                {
                    Application xlApp = (Application)ExcelDnaUtil.Application;
                    activeSheet = xlApp.ActiveSheet;
                    //System.Diagnostics.Debug.WriteLine("アプリケーション取得完了");
                }
                catch ( Exception ex )
                {
                    System.Diagnostics.Debug.WriteLine("アプリケーション取得エラー: " + ex.Message);
                    return "";
                }

                // Shapesコレクションを取得
                Microsoft.Office.Interop.Excel.Shapes shapes = activeSheet.Shapes;

                // コレクションが空の場合
                if ( shapes == null )
                {
                    return "";
                }

                int shapeCount;
                try
                {
                    shapeCount = shapes.Count;
                    if ( shapeCount == 0 )
                    {
                        return "";
                    }
                }
                catch
                {
                    return "";
                }

                // 各シェイプをチェック
                for ( int i = 1; i <= shapeCount; i++ )
                {
                    try
                    {
                        Shape shape = shapes.Item(i);

                        // 名前の一致を確認
                        if ( shape.Name.Equals(textBoxName, StringComparison.OrdinalIgnoreCase) )
                        {
                            // テキストボックスかどうか確認
                            if ( shape.Type == MsoShapeType.msoTextBox )
                            {
                                try
                                {
                                    // TextFrame2が利用可能かチェック
                                    return shape.TextFrame2.TextRange.Text;
                                }
                                catch
                                {
                                    // TextFrame2が失敗した場合はTextFrameを試す
                                    try
                                    {
                                        return shape.TextFrame.Characters().Text;
                                    }
                                    catch
                                    {
                                        return "";
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 個別のシェイプの取得に失敗した場合は次へ
                        continue;
                    }
                }

                return ""; // テキストボックスが見つからなかった場合
            }
            catch
            {
                return ""; // エラーが発生した場合
            }
        }



    }
}

