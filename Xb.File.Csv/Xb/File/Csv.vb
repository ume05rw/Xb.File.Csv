Option Strict On

''' <summary>
''' CSV�t�@�C���p�֐��Q
''' </summary>
''' <remarks></remarks>
Public Class Csv

    ''' <summary>
    ''' ��������_�u���N�H�[�g�ň͂ށB�����񒆂Ƀ_�u���N�H�[�g������ꍇ�ACSV���G�X�P�[�v���s���B
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Dquote(ByVal text As String) As String
        Return Xb.Str.CsvDquote(text)
    End Function


    ''' <summary>
    ''' DataTable�̓��e��CSV�t�H�[�}�b�g�e�L�X�g�ɕϊ�����B
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

        '�n���lDataTable�̒l���݃`�F�b�N
        If (dataTable Is Nothing) Then
            Xb.Util.Out("File.Csv.GetCsvText: DataTable�����o�ł��܂���B")
            Throw New ArgumentException("DataTable�����o�ł��܂���B")
        End If

        lfString = Xb.Str.GetLinefeed(linefeed)
        maxColumnCount = dataTable.Columns.Count
        maxRowCount = dataTable.Rows.Count
        builder = New System.Text.StringBuilder()

        For i = 0 To maxRowCount - 1
            If ((i = 0) And (isOutputHeader)) Then
                '��s�ڂɃ^�C�g���s���o�͂���
                For j = 0 To maxColumnCount - 1
                    builder.Append(If(j = 0, "", ",").ToString())

                    builder.Append(If(isQuote, _
                                        Xb.Str.CsvDquote(dataTable.Columns(j).ColumnName), _
                                        dataTable.Columns(j).ColumnName))
                Next
                builder.Append(lfString)
            End If

            '�s�̃J�����f�[�^���_�u���N�H�[�g�ň͂݁A�J���}�ŋ�؂��ĘA������B
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
    ''' DataTable�̓��e��CSV�t�@�C���ɏ����o���B
    ''' </summary>
    ''' <param name="dataTable"></param>
    ''' <param name="encode"></param>
    ''' <param name="fileName"></param>
    ''' <param name="isOutputHeader"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' �G���R�[�h����肵�����ꍇ�Ɏg�p�BSJIS����̂Ƃ��͈����Ⴂ�̓����֐����g�p����B
    ''' </remarks>
    Public Shared Function CreateCsvFile(ByRef dataTable As DataTable, _
                                         ByVal encode As System.Text.Encoding, _
                                         Optional ByVal linefeed As Xb.Str.LinefeedType = Xb.Str.LinefeedType.Lf, _
                                         Optional ByVal fileName As String = "list.csv", _
                                         Optional ByVal isOutputHeader As Boolean = True, _
                                         Optional ByVal isQuote As Boolean = True) As Boolean

        Dim writer As IO.StreamWriter, _
            csvtext As String

        '�n���lDataTable�̒l���݃`�F�b�N
        If (dataTable Is Nothing) Then Return False

        '�t�@�C�������΃p�X�ɐ��`����B
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
    ''' DataTable�̓��e��CSV�t�@�C���ɏ����o���B
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
    ''' �n���lCSV�t�H�[�}�b�g��������ADataTable�^�ɕϊ����Ď擾����B
    ''' </summary>
    ''' <param name="csvText">CSV�t�H�[�}�b�g������</param>
    ''' <param name="delimiter">�f���~�^</param>
    ''' <returns></returns>
    Public Shared Function GetDataTableByText(ByVal csvText As String, _
                                              Optional ByVal delimiter As String = ",") As DataTable

        '���̓`�F�b�N
        If (csvText Is Nothing) Then
            Xb.Util.Out("File.Csv.GetDataTableByText: CSV�e�L�X�g�����o�ł��܂���B")
            Throw New ArgumentException("CSV�e�L�X�g�����o�ł��܂���B")
        End If
        If (delimiter Is Nothing) Then delimiter = ","

        'CSV�^�̕�����f�[�^���p�[�X���鏈���́A.Net�ɂ͑��݂��Ȃ��B
        '�e�L�X�g�t�@�C���� Microsoft.VisualBasic.FileIO.TextFieldParser ���g�p����
        'CSV�p�[�X����̂��葁�����߁A�e���|�����t�@�C���ɏ������񂾏�Ńp�[�X����B
        Dim tmpName As String = Xb.App.Path.GetTempFilename("csv"), _
            writer As IO.StreamWriter, _
            encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS"), _
            result As DataTable

        Try
            writer = New IO.StreamWriter(tmpName, False, encode)
            writer.Write(csvText)
            writer.Close()
        Catch ex As Exception
            Xb.Util.Out("File.Csv.GetDataTableByText: CSV�ϊ������Ɏ��s���܂����B")
            Throw New ApplicationException("CSV�ϊ������Ɏ��s���܂����B")
        End Try

        result = GetDataTableByFile(tmpName, encode, delimiter)

        Try
            IO.File.Delete(tmpName)
        Catch ex As Exception
            Xb.Util.Out("File.Csv.GetDataTableByText: " & ex.Message)
            '�폜�Ɏ��s�����Ƃ��A�S�~�t�@�C�����c����̂́ADataTable�擾�ɂ͐������Ă��邽�߁A
            '�������Ȃ��ł����B
        End Try

        Return result

    End Function


    ''' <summary>
    ''' �n���lCSV�t�@�C���̃f�[�^���ADataTable�^�Ŏ擾����B
    ''' </summary>
    ''' <param name="fileName">CSV�t�@�C���̃p�X</param>
    ''' <param name="encode"></param>
    ''' <param name="delimiter"></param>
    ''' <returns>
    ''' TextFieldParser�ɋ��ޑO�ɁACSV���`�������o����΂������B������Ȃ�
    ''' TextFieldParser�g�킸�Ɏ��͂Ńp�[�X����̂Ɠ������B���[��B
    ''' </returns>
    Public Shared Function GetDataTableByFile(ByVal fileName As String, _
                                              Optional ByVal encode As System.Text.Encoding = Nothing, _
                                              Optional ByVal delimiter As String = ",") As DataTable

        '���̓`�F�b�N
        '�n���l�p�X�Ƀt�@�C�������݂��邩�ۂ������؂���B
        If (Not IO.File.Exists(fileName)) Then
            Xb.Util.Out("File.Csv.GetDataTableByFile: �n���l�p�X�Ƀt�@�C�������݂��܂���F" & fileName)
            Throw New Exception("�n���l�p�X�Ƀt�@�C�������݂��܂���F" & fileName)
        End If

        '�n���l�t�H�[�}�b�g
        If (encode Is Nothing) Then encode = System.Text.Encoding.GetEncoding("Shift_JIS")
        If (delimiter Is Nothing) Then delimiter = ","

        Dim maxColCnt As Integer = 0, _
            rows As List(Of String()) = New List(Of String())(), _
            isValidRow As Boolean

        '�n���l�p�X�̃t�@�C�����擾����B
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = New Microsoft.VisualBasic.FileIO.TextFieldParser(fileName, encode)
        parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
        parser.Delimiters = New String() {","}

        While (Not parser.EndOfData)
            '�t�@�C������s����������擾����B
            Dim columns As String() = parser.ReadFields()

            '�t�@�C�������擾���A�S�Ă̗v�f���󕶎���̍s���擾����Ă��܂����ۂւ̑΍�B
            isValidRow = True
            If (parser.EndOfData) Then
                '�����s�f�[�^�̂Ƃ��́A��U�s��s���Ȃ��̂Ƃ��Č��؂���B
                isValidRow = False
                For Each column As String In columns
                    '�����s�f�[�^�ɋ󕶎��łȂ��������܂܂�Ă���Ƃ��A�������s�Ƃ���B
                    '�����s�̑S�Ă̒l���󕶎��̎��͕s���f�[�^�ƌ��􂷁B
                    '�����s�̍ŏ��̗v�f���A�X�L�[�R�[�h=0�̕\���s�\�����A��ڗv�f�ȍ~�͋󕶎���ɂȂ��Ă���B
                    If (Not string.IsNullOrEmpty(column) _
                        AndAlso column.Length = 1 _
                        AndAlso Convert.ToInt32(column.Chars(0)) = 0) Then

                        isValidRow = True
                        Exit For
                    End If
                Next
            End If
            '�������s�f�[�^�̂Ƃ��̂݁A�ǉ�����B
            If (isValidRow) Then rows.Add(columns)

            If maxColCnt < columns.Length Then
                maxColCnt = columns.Length
            End If
        End While

        '�擾����CSV�̍ő�J�������Ɋ�Â��ADataTable���`����B
        Dim dt As New DataTable()
        Dim colName As String
        For i As Integer = 0 To maxColCnt - 1
            colName = System.String.Format("col{0}", (i + 1).ToString())

            '�S�ẴJ�����͕�����Ƃ��Ē�`����B
            dt.Columns.Add(colName, System.Type.GetType("System.String"))
        Next

        'DataTable�ցA�擾�����l���Z�b�g����B
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
