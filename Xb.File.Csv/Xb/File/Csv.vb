Option Strict On

''' <summary>
''' CSVファイル用関数群
''' </summary>
''' <remarks></remarks>
Public Class Csv

    ''' <summary>
    ''' 文字列をダブルクォートで囲む。文字列中にダブルクォートがある場合、CSV式エスケープを行う。
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Dquote(ByVal text As String) As String
        Return Xb.Str.CsvDquote(text)
    End Function


    ''' <summary>
    ''' DataTableの内容をCSVフォーマットテキストに変換する。
    ''' </summary>
    ''' <param name="dataTable"></param>
    ''' <param name="linefeed"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCsvText(ByRef dataTable As DataTable, _
                                      Optional ByVal linefeed As Xb.Str.LinefeedType = Xb.Str.LinefeedType.Lf, _
                                      Optional ByVal isOutputHeader As Boolean = True, _
                                      Optional ByVal isQuote As Boolean = True) As String

        Dim lfString As String, _
            builder As System.Text.StringBuilder, _
            maxColumnCount, maxRowCount, i, j As Integer

        '渡し値DataTableの値存在チェック
        If (dataTable Is Nothing) Then
            Xb.Util.Out("File.Csv.GetCsvText: DataTableが検出できません。")
            Throw New ArgumentException("DataTableが検出できません。")
        End If

        lfString = Xb.Str.GetLinefeed(linefeed)
        maxColumnCount = dataTable.Columns.Count
        maxRowCount = dataTable.Rows.Count
        builder = New System.Text.StringBuilder()

        For i = 0 To maxRowCount - 1
            If ((i = 0) And (isOutputHeader)) Then
                '一行目にタイトル行を出力する
                For j = 0 To maxColumnCount - 1
                    builder.Append(If(j = 0, "", ",").ToString())

                    builder.Append(If(isQuote, _
                                        Xb.Str.CsvDquote(dataTable.Columns(j).ColumnName), _
                                        dataTable.Columns(j).ColumnName))
                Next
                builder.Append(lfString)
            End If

            '行のカラムデータをダブルクォートで囲み、カンマで区切って連結する。
            For j = 0 To maxColumnCount - 1
                builder.Append(If(j = 0, "", ",").ToString())
                builder.Append(If(isQuote, _
                                    Xb.Str.CsvDquote(dataTable.Rows(i).Item(j).ToString()), _
                                    dataTable.Rows(i).Item(j).ToString()))
            Next
            builder.Append(lfString)
        Next

        Return builder.ToString()

    End Function


    ''' <summary>
    ''' DataTableの内容をCSVファイルに書き出す。
    ''' </summary>
    ''' <param name="dataTable"></param>
    ''' <param name="encode"></param>
    ''' <param name="fileName"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' エンコードを特定したい場合に使用。SJIS限定のときは引数違いの同名関数を使用する。
    ''' </remarks>
    Public Shared Function CreateCsvFile(ByRef dataTable As DataTable, _
                                         ByVal encode As System.Text.Encoding, _
                                         Optional ByVal linefeed As Xb.Str.LinefeedType = Xb.Str.LinefeedType.Lf, _
                                         Optional ByVal fileName As String = "list.csv", _
                                         Optional ByVal isOutputHeader As Boolean = True, _
                                         Optional ByVal isQuote As Boolean = True) As Boolean

        Dim writer As IO.StreamWriter, _
            csvtext As String

        '渡し値DataTableの値存在チェック
        If (dataTable Is Nothing) Then Return False

        'ファイル名を絶対パスに整形する。
        Try
            fileName = Xb.App.Path.GetAbsPath(fileName)
        Catch ex As Exception
            Return False
        End Try

        Try
            csvtext = GetCsvText(dataTable, linefeed, isOutputHeader, isQuote)
            writer = New IO.StreamWriter(fileName, False, encode)
            writer.Write(csvtext)
            writer.Close()
        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function


    ''' <summary>
    ''' DataTableの内容をCSVファイルに書き出す。
    ''' </summary>
    ''' <param name="dataTable"></param>
    ''' <param name="fileName"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateCsvFile(ByRef dataTable As DataTable, _
                                         Optional ByVal fileName As String = "list.csv", _
                                         Optional ByVal isOutputHeader As Boolean = True) As Boolean

        Return CreateCsvFile(dataTable, _
                        System.Text.Encoding.GetEncoding("Shift_JIS"), _
                        Xb.Str.LinefeedType.CrLf, _
                        fileName, _
                        isOutputHeader)
    End Function


    ''' <summary>
    ''' 渡し値CSVフォーマット文字列を、DataTable型に変換して取得する。
    ''' </summary>
    ''' <param name="csvText">CSVフォーマット文字列</param>
    ''' <param name="delimiter">デリミタ</param>
    ''' <returns></returns>
    Public Shared Function GetDataTableByText(ByVal csvText As String, _
                                              Optional ByVal delimiter As String = ",") As DataTable

        '入力チェック
        If (csvText Is Nothing) Then
            Xb.Util.Out("File.Csv.GetDataTableByText: CSVテキストが検出できません。")
            Throw New ArgumentException("CSVテキストが検出できません。")
        End If
        If (delimiter Is Nothing) Then delimiter = ","

        'CSV型の文字列データをパースする処理は、.Netには存在しない。
        'テキストファイルを Microsoft.VisualBasic.FileIO.TextFieldParser を使用して
        'CSVパースするのが手早いため、テンポラリファイルに書き込んだ上でパースする。
        Dim tmpName As String = Xb.App.Path.GetTempFilename("csv"), _
            writer As IO.StreamWriter, _
            encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS"), _
            result As DataTable

        Try
            writer = New IO.StreamWriter(tmpName, False, encode)
            writer.Write(csvText)
            writer.Close()
        Catch ex As Exception
            Xb.Util.Out("File.Csv.GetDataTableByText: CSV変換処理に失敗しました。")
            Throw New ApplicationException("CSV変換処理に失敗しました。")
        End Try

        result = GetDataTableByFile(tmpName, encode, delimiter)

        Try
            IO.File.Delete(tmpName)
        Catch ex As Exception
            Xb.Util.Out("File.Csv.GetDataTableByText: " & ex.Message)
            '削除に失敗したとき、ゴミファイルが残るものの、DataTable取得には成功しているため、
            '何もしないでおく。
        End Try

        Return result

    End Function


    ''' <summary>
    ''' 渡し値CSVファイルのデータを、DataTable型で取得する。
    ''' </summary>
    ''' <param name="fileName">CSVファイルのパス</param>
    ''' <param name="encode"></param>
    ''' <param name="delimiter"></param>
    ''' <returns>
    ''' TextFieldParserに挟む前に、CSV整形処理が出来ればいいが。それやるなら
    ''' TextFieldParser使わずに自力でパースするのと同じか。うーん。
    ''' </returns>
    Public Shared Function GetDataTableByFile(ByVal fileName As String, _
                                              Optional ByVal encode As System.Text.Encoding = Nothing, _
                                              Optional ByVal delimiter As String = ",") As DataTable

        '入力チェック
        '渡し値パスにファイルが存在するか否かを検証する。
        If (Not IO.File.Exists(fileName)) Then
            Xb.Util.Out("File.Csv.GetDataTableByFile: 渡し値パスにファイルが存在しません：" & fileName)
            Throw New Exception("渡し値パスにファイルが存在しません：" & fileName)
        End If

        '渡し値フォーマット
        If (encode Is Nothing) Then encode = System.Text.Encoding.GetEncoding("Shift_JIS")
        If (delimiter Is Nothing) Then delimiter = ","

        Dim maxColCnt As Integer = 0, _
            rows As List(Of String()) = New List(Of String())(), _
            isValidRow As Boolean

        '渡し値パスのファイルを取得する。
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = New Microsoft.VisualBasic.FileIO.TextFieldParser(fileName, encode)
        parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
        parser.Delimiters = New String() {","}

        While (Not parser.EndOfData)
            'ファイルを一行ずつ文字列を取得する。
            Dim columns As String() = parser.ReadFields()

            'ファイル末尾取得時、全ての要素が空文字列の行が取得されてしまう現象への対策。
            isValidRow = True
            If (parser.EndOfData) Then
                '末尾行データのときは、一旦行を不正なものとして検証する。
                isValidRow = False
                For Each column As String In columns
                    '末尾行データに空文字でない何かが含まれているとき、正しい行とする。
                    '末尾行の全ての値が空文字の時は不正データと見做す。
                    '末尾行の最初の要素がアスキーコード=0の表示不能文字、二つ目要素以降は空文字列になっている。
                    If (Not string.IsNullOrEmpty(column) _
                        AndAlso column.Length = 1 _
                        AndAlso Convert.ToInt32(column.Chars(0)) = 0) Then

                        isValidRow = True
                        Exit For
                    End If
                Next
            End If
            '正しい行データのときのみ、追加する。
            If (isValidRow) Then rows.Add(columns)

            If maxColCnt < columns.Length Then
                maxColCnt = columns.Length
            End If
        End While

        '取得したCSVの最大カラム数に基づき、DataTableを定義する。
        Dim dt As New DataTable()
        Dim colName As String
        For i As Integer = 0 To maxColCnt - 1
            colName = System.String.Format("col{0}", (i + 1).ToString())

            '全てのカラムは文字列として定義する。
            dt.Columns.Add(colName, System.Type.GetType("System.String"))
        Next

        'DataTableへ、取得した値をセットする。
        Dim dr As DataRow
        For i As Integer = 0 To rows.Count - 1
            dr = dt.NewRow()

            For j As Integer = 0 To rows(i).Length - 1
                dr(j) = rows(i)(j)
            Next

            dt.Rows.Add(dr)
        Next

        rows.Clear()
        parser.Dispose()

        Return dt

    End Function



End Class
