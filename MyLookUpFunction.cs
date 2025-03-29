using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public class MyLookupFunction
    {

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

    }

}