Imports System
Imports Microsoft.Office.Interop.Excel

Public Class ExcelUtility
    Implements IDisposable

#Region "定数定義"
    Private Const _C_EXCEL_LINESTYLE_XLCONTINUOUS As Integer = 1        'Excel罫線－実線
    Private Const _C_EXCEL_XLHALIGN_XLHALIGNCENTER As Integer = -4108   'Excel書式－水平中央
    Private Const _C_EXCEL_XLVALIGN_XLVALIGNTOP As Integer = -4160      'Excel書式－垂直上詰め
    Public Const _C_EXCEL_PASTE_TYPE_ALL = -4104
    Public Const _C_EXCEL_PASTE_TYPE_FORMULAS = -4123
    Public Const _C_EXCEL_PASTE_TYPE_VALUES = -4163
    Public Const _C_EXCEL_PASTE_TYPE_FORMATS = -4122

    ''' <summary>
    ''' ファイル保存時のフォーマット
    ''' </summary>
    Public Enum XlFileFormat
        ''' <summary>2003以前の場合のxls</summary>
        xlExcel9795 = 43
        ''' <summary>2007以降の場合のxls</summary>
        xlExcel8 = 56
        ''' <summary>xls作成時に、～2003と2007～をシステムで自動振り分けする場合はこれを使用</summary>
        xlExcelXls = 99999
        ''' <summary>デフォルトそのまま</summary>
        xlWorkbookNormal = -4143
    End Enum

    ''' <summary>
    ''' 線の種類
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum LineStyle
        ''' <summary>実線</summary>
        xlContinuous = 1
        ''' <summary>線無し</summary>
        xlLineStyleNone = -4142
    End Enum

    ''' <summary>
    ''' 罫線を引く位置
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum BorderIdx
        ''' <summary>セルの左の線</summary>
        xlEdgeLeft = 7
        ''' <summary>セルの上の線</summary>
        xlEdgeTop = 8
        ''' <summary>セルの下の線</summary>
        xlEdgeBottom = 9
        ''' <summary>セルの右の線</summary>
        xlEdgeRight = 10
        ''' <summary>セル範囲の内側の横線</summary>
        xlInsideVertical = 11
        ''' <summary>セル範囲の内側の縦線</summary>
        xlInsideHorizontal = 12
        ''' <summary>外枠</summary>
        Sokowaku = -1
        ''' <summary>格子</summary>
        Koushi = -2
    End Enum

    ''' <summary>
    ''' セルの移動方向
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum Shift
        ''' <summary>下</summary>
        xlShiftDown = -4121
    End Enum
#End Region

    Private _excelApp As Object
    Private _excelWbk As Object
    Private _excelWst As Object
    Private _excelRange As Object
    Private _fileFullName As String
    Private _rownum As Long
    Public _visible As Boolean = False

    '''' <summary>シート数</summary>
    'Public ReadOnly Property WorkSheetCount() As Integer
    '    Get
    '        Try
    '            Return _excelWbk.Worksheets.Count()
    '        Catch
    '            Return -1
    '        End Try
    '    End Get
    'End Property

    ''' <summary>
    ''' Excelのバージョン
    ''' </summary>
    'Public ReadOnly Property Version() As Decimal
    '    Get
    '        Try
    '            Return _excelApp.Version
    '        Catch
    '            Return ""
    '        End Try
    '    End Get
    'End Property
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' ファイルを開きます
    ''' </summary>
    ''' <param name="fileFullName">開くファイルをFullNameで指定します</param>
    ''' <remarks></remarks>
    Public Sub Open(ByVal fileFullName As String)
        Try
            _fileFullName = fileFullName
            _excelApp = CreateObject("Excel.application")
            _excelWbk = _excelApp.Workbooks.Open(_fileFullName)

        Catch ex As Exception
            _excelApp.Quit()
            _excelApp = Nothing
            GC.Collect()
            Throw New NotOpenException(fileFullName + "のオープンに失敗しました")
        End Try
    End Sub

    '''' <summary>
    '''' 操作するシートを設定します
    '''' </summary>
    '''' <param name="sheetName">シート名</param>
    '''' <remarks></remarks>
    'Public Sub SetWorksheet(ByVal sheetName As String)
    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBookオブジェクトの参照が有効ではありません")
    '        Exit Sub
    '    End If

    '    Try

    '        _excelWst = _excelWbk.Worksheets(sheetName)
    '    Catch ex As Exception
    '        Throw New NotValueException("Worksheets(" + sheetName + ")の設定に失敗しました")
    '    End Try

    '    GC.Collect()
    'End Sub

    '''' <summary>
    '''' 指定されたシートをアクティブにします。
    '''' </summary>
    '''' <param name="index">アクティブにするシート番号です</param>
    '''' <remarks>index >= 1　です</remarks>
    'Public Sub SetActiveSheet(ByVal index As Integer)
    '    Me._excelWbk.Sheets(index).Select()
    '    Me.SetWorksheet(index)
    'End Sub


    ''' <summary>
    ''' 操作するシートを設定します
    ''' </summary>
    ''' <param name="sheetIndex">シート番号</param>
    ''' <remarks></remarks>
    Public Sub SetWorksheet(ByVal sheetIndex As Integer)
        If _excelWbk Is Nothing Then
            Throw New NotReferException("WorkBookオブジェクトの参照が有効ではありません")
            Exit Sub
        End If

        If Not _excelWst Is Nothing Then
            MRComObject(_excelWst)
        End If


        Try
            _excelWst = _excelWbk.Worksheets(sheetIndex)
        Catch ex As Exception
            Throw New NotValueException("Worksheets(" + sheetIndex.ToString + ")の設定に失敗しました")
        End Try
        GC.Collect()
    End Sub

    '''' <summary>
    '''' 操作しているシートを一番最後にコピーします。
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub CopyWorksheet()
    '    Try
    '        Me._excelApp.ScreenUpdating = False
    '        Dim sheetCnt As Integer = _excelWbk.Worksheets.Count()
    '        _excelWst.Copy(, Me._excelWbk.Sheets(sheetCnt))
    '    Catch ex As Exception
    '        Throw New Exception("コピー失敗:" + ex.Message)
    '    Finally
    '        Me._excelApp.ScreenUpdating = True
    '    End Try
    'End Sub


    '''' <summary>
    '''' 操作しているシートを指定された位置の後ろにコピーします。
    '''' </summary>
    '''' <param name="decIndex">このシート番号の後ろにコピーされます</param>
    '''' <remarks></remarks>
    'Public Sub CopyWorksheet(ByVal decIndex As Integer)
    '    Try
    '        Me._excelApp.ScreenUpdating = False
    '        Dim sheetCnt As Integer = _excelWbk.Worksheets.Count()
    '        _excelWst.Copy(, Me._excelWbk.Sheets(decIndex))
    '    Catch ex As Exception
    '        Throw New Exception("コピー失敗:" + ex.Message)
    '    Finally
    '        Me._excelApp.ScreenUpdating = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' 操作しているシートの名前を変更します。
    '''' </summary>
    '''' <param name="name">新しいシート名を指定します</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal name As String)
    '    Try
    '        Me._excelWst.Name = name
    '    Catch ex As Exception
    '        Throw New Exception("リネーム失敗:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定されたシート番号のシート名を変更します。
    '''' </summary>
    '''' <param name="index">名前を変えるシートのシート番号を指定します。</param>
    '''' <param name="name">新しいシート名を指定します。</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal index As Integer, ByVal name As String)
    '    Try

    '        Me._excelWbk.Sheets(index).Name = name
    '    Catch ex As Exception
    '        Throw New Exception("リネーム失敗:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定されたシート番号のシートを削除します。
    '''' </summary>
    '''' <param name="index">削除するシートのシート番号</param>
    '''' <remarks></remarks>
    'Public Sub DeleteSheet(ByVal index As Integer)
    '    Try
    '        Me._excelApp.DisplayAlerts = False
    '        Me._excelWbk.Sheets(index).Delete()
    '    Catch ex As Exception
    '        Throw New Exception("削除失敗:" + ex.Message)
    '    Finally
    '        Me._excelApp.DisplayAlerts = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定されたシート番号のシートを削除します。
    '''' </summary>
    '''' <param name="sheetName">削除するシートの名称</param>
    '''' <remarks></remarks>
    'Public Sub DeleteSheet(ByVal sheetName As String)
    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBookオブジェクトの参照が有効ではありません")
    '        Exit Sub
    '    End If

    '    Try
    '        Me._excelApp.DisplayAlerts = False
    '        _excelWst = _excelWbk.Worksheets(sheetName)
    '        Me._excelWst.Delete()
    '    Catch ex As Exception
    '        Throw New Exception("削除失敗:" + ex.Message)
    '    Finally
    '        Me._excelApp.DisplayAlerts = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定された列を削除します。
    '''' </summary>
    '''' <param name="delstr">削除する列のアルファベットを指定します。</param>
    '''' <remarks></remarks>
    'Public Sub DeleteCol(ByVal delstr As String)

    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If
    '        'xlShiftToLeft=-4159
    '        '_excelRange = _excelWst.Range(delstr)
    '        _excelRange = _excelWst.Range(delstr & ":" & delstr)
    '        Me._excelRange.delete(Shift:=-4159)

    '    Catch ex As Exception

    '        Throw New Exception("削除エラー:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定された列を削除します。
    '''' </summary>
    '''' <param name="startColStr">削除する列の開始アルファベットを指定します。</param>
    '''' ''' <param name="endColStr">削除する列の終了アルファベットを指定します。</param>
    '''' <remarks></remarks>
    'Public Sub DeleteCols(ByVal startColStr As String, ByVal endColStr As String)

    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If
    '        'xlShiftToLeft=-4159
    '        _excelRange = _excelWst.Range(startColStr + ":" + endColStr)
    '        Me._excelRange.delete(Shift:=-4159)

    '    Catch ex As Exception

    '        Throw New Exception("削除エラー:" + ex.Message)

    '    End Try
    'End Sub

    ''' <summary>
    ''' 指定された行を削除します。
    ''' </summary>
    ''' <param name="rowIndex">削除したい行</param>
    ''' <remarks></remarks>
    Public Sub DeleteRow(ByVal rowIndex As Integer)
        DeleteRows(rowIndex, rowIndex)
    End Sub

    ''' <summary>
    ''' 指定された行を削除します。
    ''' </summary>
    ''' <param name="startRowIndex">削除したい開始行</param>
    ''' ''' <param name="endRowIndex">削除したい終了行</param>
    ''' <remarks></remarks>
    Public Sub DeleteRows(ByVal startRowIndex As Integer, ByVal endRowIndex As Integer)

        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range(startRowIndex & ":" & endRowIndex)
            Me._excelRange.delete()

        Catch ex As Exception

            Throw New Exception("削除エラー:" + ex.Message)

        End Try
    End Sub

    '''' <summary>
    '''' 指定された名前のシート名を変更します。
    '''' </summary>
    '''' <param name="oldname">変更するシートの名前です。</param>
    '''' <param name="name">新しいシート名です。</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal oldname As String, ByVal name As String)
    '    Try
    '        Me._excelWbk.Sheets(oldname).Name = name
    '    Catch ex As Exception
    '        Throw New Exception("リネーム失敗:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定された行をコピーして挿入します。
    '''' </summary>
    '''' <param name="srcRowIndex">コピーする行番号です。</param>
    '''' <param name="count">何行挿入するかを指定します。</param>
    '''' <remarks></remarks>
    'Public Sub CopyRow(ByVal srcRowIndex As Integer, ByVal count As Integer)
    '    Try
    '        If count = 0 Then
    '            Exit Sub
    '        End If
    '        Me._excelWst.Rows(srcRowIndex).Copy()
    '        Me._excelWst.Rows(srcRowIndex + 1).Resize(count).Insert(-4121)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try
    'End Sub
    ''' <summary>
    ''' 指定された行範囲をコピーして挿入します。
    ''' </summary>
    ''' <param name="srcStartRowIndex">コピー範囲の開始行番号です。</param>
    ''' <param name="srcEndRowIndex">コピー範囲の終了行番号です。</param>
    ''' <param name="insertRowIndex">挿入行番号です。</param>
    ''' <remarks></remarks>
    'Public Sub CopyRow(ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertRowIndex As Integer)
    '    Try
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()
    '        Me._excelWst.Rows(insertRowIndex).Insert(-4121)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try
    'End Sub
    ''' <summary>
    ''' 指定された行範囲をコピーして挿入します。
    ''' </summary>
    ''' <param name="srcStartRowIndex">コピー範囲の開始行番号です。</param>
    ''' <param name="srcEndRowIndex">コピー範囲の終了行番号です。</param>
    ''' <param name="insertStartRowIndex">挿入先の開始行番号です。</param>
    ''' <param name="insertEndRowIndex">挿入先の終了行番号です。</param>
    ''' <remarks></remarks>
    'Public Sub CopyRow(ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertStartRowIndex As Integer, ByVal insertEndRowIndex As Integer)
    '    Try
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()
    '        Me._excelWst.Rows(insertStartRowIndex.ToString() + ":" + insertEndRowIndex.ToString()).Insert(-4121)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try
    'End Sub
    '''' <summary>
    '''' 指定された列をコピーして挿入します
    '''' </summary>
    '''' <param name="srcColIndex">コピーする列番号です。</param>
    '''' <param name="count">何列挿入するかを指定します。</param>
    '''' <remarks></remarks>
    'Public Sub CopyCol(ByVal srcColIndex As Integer, ByVal count As Integer)
    '    Dim cnt As Integer
    '    Try
    '        If count = 0 Then
    '            Exit Sub
    '        End If

    '        For cnt = 1 To count
    '            Me._excelWst.Columns(srcColIndex).Copy()
    '            Me._excelWst.Columns(srcColIndex + 1).Resize(1).Insert(-4161)
    '        Next cnt
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try
    'End Sub
    '''' <summary>
    '''' 指定された列をコピーして挿入します
    '''' </summary>
    '''' <param name="srcRangeAddress">コピー元の列を範囲で指定</param>
    '''' <param name="decRangeAddress">貼り付け先の列を範囲で指定</param>
    '''' <remarks></remarks>
    'Public Sub CopyCol(ByVal srcRangeAddress As String, ByVal decRangeAddress As String)

    '    Try
    '        Me._excelWst.Columns(srcRangeAddress).Copy()
    '        Me._excelWst.Columns(decRangeAddress).Resize(1).Insert(-4161)

    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' セルをコピーして貼り付けます。
    '''' </summary>
    '''' <param name="srcRangeAddress">コピー元のセルを範囲で指定</param>
    '''' <param name="decRangeAddress">コピー先のセルを範囲で指定</param>
    '''' <remarks></remarks>
    'Public Sub CopyCells(ByVal srcRangeAddress As String, ByVal decRangeAddress As String, ByVal pasteType As Integer)

    '    Try

    '        Me._excelWst.Range(srcRangeAddress).Copy()
    '        Me._excelWst.Range(decRangeAddress).PasteSpecial(pasteType)

    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("コピーエラー:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' セルを選択します
    '''' </summary>
    '''' <param name="rangeAddress">"A1"形式のセル番地を指定します</param>
    '''' <remarks></remarks>
    'Public Sub SelectCell(ByVal rangeAddress As String)
    '    If Not _excelRange Is Nothing Then
    '        MRComObject(_excelRange)
    '    End If

    '    _excelRange = _excelWst.Range(rangeAddress)
    '    _excelRange.Select()
    'End Sub

    '''' <summary>
    '''' セルの値を設定します。
    '''' </summary>
    '''' <param name="rangeAddress">"A1"形式のセル番地を指定します</param>
    '''' <param name="value">セルに設定する値です</param>
    '''' <remarks></remarks>
    'Public Sub SetValue(ByVal rangeAddress As String, ByVal value As String)
    '    If Not _excelRange Is Nothing Then
    '        MRComObject(_excelRange)
    '    End If

    '    Try

    '        _excelRange = _excelWst.Range(rangeAddress)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException("Range(" + rangeAddress + ")に値が書き込めません")
    '    End Try
    'End Sub

    ''' <summary>
    ''' セルの値を設定します
    ''' </summary>
    ''' <param name="row">行番号を指定します</param>
    ''' <param name="col">列番号を指定します</param>
    ''' <param name="value">セルに設定する値です</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try
            _excelRange = _excelWst.Cells(row, col)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "に値が書き込めません")
        End Try
    End Sub

    ''' <summary>
    ''' セルの値を設定します.
    ''' 文字のサイズを指定します。
    ''' </summary>
    ''' <param name="row">行番号を指定します</param>
    ''' <param name="col">列番号を指定します</param>
    ''' 
    ''' <param name="value">セルに設定する値です</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal row As Integer, ByVal col As Integer, ByVal value As String, ByVal size As Integer)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try

            _excelRange = _excelWst.Cells(row, col)

            _excelRange.Font.Size = size

            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "に値が書き込めません")
        End Try
    End Sub

    ''' <summary>
    ''' セルの値を設定します。
    ''' </summary>
    ''' <param name="rangeAddress">"A1"形式のセル番地を指定します</param>
    ''' <param name="value">セルに設定する値です</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal rangeAddress As String, ByVal value As Object)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try

            _excelRange = _excelWst.Range(rangeAddress)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Range(" + rangeAddress + ")に値が書き込めません")
        End Try
    End Sub

    ''' <summary>
    ''' セルの値を設定します
    ''' </summary>
    ''' <param name="row">行番号を指定します</param>
    ''' <param name="col">列番号を指定します</param>
    ''' <param name="value">セルに設定する値です</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try
            _excelRange = _excelWst.Cells(row, col)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "に値が書き込めません")
        End Try
    End Sub

    ''' <summary>
    ''' セルの値を取得します
    ''' </summary>
    ''' <param name="rangeAddress">"A1"形式でセル番地を指定します</param>
    ''' <returns>セルの値（文字列形式）</returns>
    ''' <remarks></remarks>
    'Public Function GetValue(ByVal rangeAddress As String) As String
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rangeAddress)
    '        Return _excelRange.Value
    '    Catch ex As Exception
    '        Throw New NotValueException("Range(" + rangeAddress + ")の値が読み込めません")
    '    End Try

    'End Function

    '''' <summary>
    '''' セルの値を取得します
    '''' </summary>
    '''' <param name="row">行番号を指定します</param>
    '''' <param name="col">列番号を指定します</param>
    '''' <returns>セルの値（文字列形式）</returns>
    '''' <remarks></remarks>
    Public Function GetValue(ByVal row As Integer, ByVal col As Integer) As String
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Cells(row, col)

            Return _excelRange.Value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")の値が読み込めません")
        End Try
    End Function

    '''' <summary>
    '''' セルの値を加工して取得します
    '''' </summary>
    '''' <param name="row">行番号を指定します</param>
    '''' <param name="col">列番号を指定します</param>
    '''' <returns>セルの値（文字列形式）</returns>
    '''' <remarks>https://minor.hatenablog.com/entry/2016/01/18/230540</remarks>
    'ikubo
    Public Function GetValueWithTag(ByVal row As Integer, ByVal col As Integer) As String
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Cells(row, col)

            Dim strTarget As String = _excelRange.Value


            Dim strKanjiValue As String
            Dim strBoldValue As String
            Dim strUnderlineValue As String
            Dim strRedColorValue As String
            Dim colorRed As Object = RGB(255, 0, 0)
            Dim txtKanjiList As New ArrayList
            Dim txtBoldList As New ArrayList
            Dim txtUnderlineList As New ArrayList
            Dim txtRedColorList As New ArrayList

            'エフェクト文字検索
            For i = 0 To strTarget.Length - 1

                Dim targetChar As Object = _excelRange.Characters(1 + i, 1)
                Dim targetNextChar As Object = _excelRange.Characters(2 + i, 1)

                '漢字取得
                If Asc(targetChar.Text) >= -30561 Then
                    strKanjiValue += strTarget(i)

                    If Asc(targetNextChar.Text) < -30561 Then
                        txtKanjiList.Add(strKanjiValue)
                        strKanjiValue = ""
                    End If
                End If


                '太文字取得
                If targetChar.Font.Bold = True Then

                    strBoldValue += strTarget(i)

                    If targetNextChar.Font.Bold = False Then
                        txtBoldList.Add(strBoldValue)
                        strBoldValue = ""
                    End If

                End If

                '下線取得
                If targetChar.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle Then
                    strUnderlineValue += strTarget(i)
                    If targetNextChar.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone Then
                        txtUnderlineList.Add(strUnderlineValue)
                        strUnderlineValue = ""
                    End If
                End If

                '赤文字取得
                If targetChar.Font.Color = colorRed Then
                    strRedColorValue += strTarget(i)
                    If targetNextChar.Font.Color = False Then
                        txtRedColorList.Add(strRedColorValue)
                        strRedColorValue = ""
                    End If
                End If


            Next

            'ルビを追加してことによって対象文字列の間隔があくとうまく適用されない
            '下線変換
            For i = 0 To txtUnderlineList.Count - 1
                strTarget = strTarget.Replace(txtUnderlineList(i), "<u style = ""text-decoration-color: black"">" + txtUnderlineList(i) + "</u>")
            Next

            '太文字変換
            For i = 0 To txtBoldList.Count - 1
                strTarget = strTarget.Replace(txtBoldList(i), "<strong>" + txtBoldList(i) + "</strong>")
            Next

            '赤文字変換
            For i = 0 To txtRedColorList.Count - 1
                strTarget = strTarget.Replace(txtRedColorList(i), "<span style = ""color:red"">" + txtRedColorList(i) + "</span>")
            Next

            'ルビ変換
            For i = 0 To txtKanjiList.Count - 1
                strTarget = strTarget.Replace(txtKanjiList(i), "<ruby>" + txtKanjiList(i))
            Next


            strTarget = strTarget.Replace("（", "<rt>").Replace("）", "</rt></ruby>")

            Return strTarget


        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")の値が読み込めません")
        End Try
    End Function



    '''' <summary>
    '''' A列の行数を取得します
    '''' </summary>
    '''' <returns>A列の行数</returns>
    '''' <remarks></remarks>
    Public Function GetRowNum() As Long
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range("A65536")
            'xlUP = -4162
            _rownum = _excelRange.End(-4162).Row

            Return _rownum
        Catch ex As Exception
            Throw New NotValueException("行数が取得できません")
        End Try
    End Function

    '''' <summary>
    '''' 指定列の行数を取得します
    '''' </summary>
    '''' <param name="colStr">A-Z表記の列</param>
    '''' <returns>指定列の行数</returns>
    '''' <remarks></remarks>
    Public Function GetRowNum(ByVal colStr As String) As Long
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range(colStr + "65536")
            'xlUP = -4162
            _rownum = _excelRange.End(-4162).Row

            '       Return _excelRange.Value
            Return _rownum
        Catch ex As Exception
            Throw New NotValueException("行数が取得できません")
        End Try
    End Function

    Public Function GetRowNum(ByVal colStr As Integer) As Long
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range(colStr + "65536")
            'xlUP = -4162
            _rownum = _excelRange.End(-4162).Row

            '       Return _excelRange.Value
            Return _rownum
        Catch ex As Exception
            Throw New NotValueException("行数が取得できません")
        End Try
    End Function

    '''' <summary>
    '''' 複数範囲のセルの値を2次元配列で一括設定します
    '''' </summary>
    '''' <param name="baseRangeAddress">値を設定する左上のセルを指定します</param>
    '''' <param name="value">設定する2次元配列です。</param>
    '''' <remarks></remarks>
    'Public Sub SetValues(ByVal baseRangeAddress As String, ByVal value As Object(,))
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If


    '        _excelRange = _excelWst.Range(baseRangeAddress).Resize(value.GetUpperBound(0) + 1, value.GetUpperBound(1) + 1)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException(baseRangeAddress + "に書き込めません" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 複数範囲のセルの値を2次元配列で一括設定します
    '''' </summary>
    '''' <param name="row">値を設定する左上のセルの行番号を指定します</param>
    '''' <param name="col">値を設定する左上のセルの列番号を指定します</param>
    '''' <param name="value">設定する2次元配列です。</param>
    '''' <remarks></remarks>
    'Public Sub SetValues(ByVal row As Integer, ByVal col As Integer, ByVal value As Object(,))
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Cells(row, col).Resize(value.GetUpperBound(0) + 1, value.GetUpperBound(1) + 1)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "に値が書き込めません")
    '    End Try
    'End Sub


    '''' <summary>
    '''' 指定されたセル範囲の値を2次元配列で取得します（文字列型）
    '''' </summary>
    '''' <param name="rangeAddress">"A1"形式で取得するセル範囲を指定します</param>
    '''' <returns>セルの値です（2次元配列）</returns>
    '''' <remarks></remarks>
    ''''Public Function GetValues(ByVal rangeAddress As String) As String(,)
    'Public Function GetValues(ByVal rangeAddress As String) As Object(,)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rangeAddress)
    '        Return _excelRange.Value
    '    Catch ex As Exception
    '        Throw New NotValueException(rangeAddress + "から読み込めません")
    '    End Try
    'End Function

    '''' <summary>
    '''' 指定されたセル範囲の値を2次元配列で取得します（文字列型）
    '''' </summary>
    '''' <param name="row">値を取得する左上のセルの行番号を指定します</param>
    '''' <param name="col">値を取得する左上のセルの列番号を指定します</param>
    '''' <param name="offsetX">取得する行数を指定します</param>
    '''' <param name="offsetY">取得する列数を指定します</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    Public Function GetValues(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer) As Object(,)
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))
            Return _excelRange.Value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")から読み込めません")
        End Try
    End Function

    ''' <summary>
    ''' 印刷を行う
    ''' </summary>
    ''' <param name="printName">プリンター名</param>
    ''' <param name="count">印刷枚数</param>
    ''' <remarks></remarks>
    Public Sub Print(ByVal printName As String, ByVal count As Integer)
        Try
            Me._excelWst.PrintOut(1, 100, count, False, printName, Type.Missing, Type.Missing, Type.Missing)
        Catch ex As Exception
            Throw New Exception("印刷失敗:" + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 保存します
    ''' </summary>
    ''' <param name="fileFullName">保存ファイル名を指定します</param>
    ''' <param name="Alret">Excelからの確認メッセージの表示/非表示を指定します</param>
    ''' <remarks></remarks>
    Public Sub SaveAs(ByVal fileFullName As String, Optional ByVal Alret As Boolean = True)
        Try
            If Alret = False Then
                _excelApp.DisplayAlerts = False
            End If
            _excelWbk.SaveAs(fileFullName)
        Catch ex As Exception
            Throw ex
        Finally
            _excelApp.DisplayAlerts = True
        End Try
    End Sub

    ''' <summary>
    ''' 保存ファイル形式を指定して保存します。
    ''' テンプレートファイルと保存ファイルの拡張子が違う場合に利用します。
    ''' </summary>
    ''' <param name="fileFullName">保存ファイル名を指定します</param>
    ''' <param name="format">保存ファイル形式を指定します</param>
    ''' <remarks></remarks>
    Public Sub SaveAs(ByVal fileFullName As String, ByVal format As XlFileFormat)
        Try
            '保存ダイアログ非表示
            _excelApp.DisplayAlerts = False

            Dim wkFormat As XlFileFormat = format

            'xlsを振り分けて登録する場合
            If format = XlFileFormat.xlExcelXls Then
                'Ver.12 = Excel2007
                'If Version >= 12 Then
                '    wkFormat = XlFileFormat.xlExcel8
                'Else
                wkFormat = XlFileFormat.xlExcel9795
                'End If
            End If

            _excelWbk.SaveAs(fileFullName, FileFormat:=wkFormat)
        Catch ex As Exception
            Throw ex
        Finally
            _excelApp.DisplayAlerts = True
        End Try
    End Sub

    ''' <summary>
    ''' ファイルを閉じます。
    ''' </summary>
    ''' <param name="save">保存して閉じるかどうかを指定します</param>
    ''' <remarks></remarks>
    Public Sub Close(Optional ByVal save As Boolean = False)
        On Error Resume Next
        If save = True Then
            _excelWbk.Save()
            _excelWbk.Saved = True
        Else
            _excelWbk.Saved = True
        End If
        _excelWbk.Close()
        _excelApp.Quit()
        Me.Dispose()
    End Sub

    '''' <summary>
    '''' 印刷範囲を設定します。
    '''' </summary>
    '''' <param name="startcell">開始セル</param>
    '''' <param name="endcell">終了セル</param>
    '''' <remarks></remarks>
    'Public Sub SetPrintArea(ByVal startcell As String, ByVal endcell As String)
    '    _excelWst.PageSetup.PrintArea = startcell & ":" & endcell
    'End Sub

    '''' <summary>
    '''' 印刷範囲をクリアします。
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub SetPrintAreaClear()
    '    _excelWst.PageSetup.PrintArea = ""
    'End Sub



    ''' <summary>
    ''' 終了処理です
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements System.IDisposable.Dispose
        MRComObject(_excelRange)
        MRComObject(_excelWst)
        MRComObject(_excelWbk)
        MRComObject(_excelApp)
        _excelRange = Nothing
        _excelWst = Nothing
        _excelWbk = Nothing
        _excelApp = Nothing
        GC.Collect()
    End Sub

    '''' <summary>
    '''' Excelの画面への表示/非表示を設定します
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Property Visible() As Boolean
    '    Get
    '        Return Me._visible
    '    End Get
    '    Set(ByVal value As Boolean)
    '        Me._visible = value
    '        Me._excelApp.Visible = Me._visible
    '        Me._excelWbk.Activate()

    '    End Set
    'End Property

    ''' <summary>
    ''' ファイルオープンに失敗した場合の例外
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NotOpenException
        Inherits System.ApplicationException
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    Private Class NotReferException
        Inherits System.ApplicationException
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    Private Class NotValueException
        Inherits System.ApplicationException
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class


    Private Sub MRComObject(ByRef objCom As Object)
        'COM オブジェクトの使用後、明示的に COM オブジェクトへの参照を解放する
        Try
            '提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
            If Not objCom Is Nothing AndAlso System.Runtime.InteropServices. _
                                                      Marshal.IsComObject(objCom) Then
                Dim I As Integer
                Do
                    I = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                Loop Until I <= 0
            End If
        Catch
        Finally
            '参照を解除する
            objCom = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 指定されたセル範囲内に規定線を引きます
    ''' </summary>
    ''' <param name="row">線を引く左上のセルの行番号を指定します</param>
    ''' <param name="col">線を引く左上のセルの列番号を指定します</param>
    ''' <param name="offsetX">線を引く行数を指定します</param>
    ''' <param name="offsetY">線を引く列数を指定します</param>
    ''' <param name="lineStyle">線の種類を指定します</param>
    ''' <remarks></remarks>
    'Public Sub SetBorders(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal lineStyle As Integer)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        With _excelRange.Borders
    '            .LineStyle = _C_EXCEL_LINESTYLE_XLCONTINUOUS
    '        End With

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定されたセル範囲内の指定位置に指定線を引きます
    '''' </summary>
    '''' <param name="row">線を引く左上のセルの行番号を指定します</param>
    '''' <param name="col">線を引く左上のセルの列番号を指定します</param>
    '''' <param name="offsetX">線を引く行数を指定します</param>
    '''' <param name="offsetY">線を引く列数を指定します</param>
    '''' <param name="bordersIndex">線を引く位置を指定します</param>
    '''' <param name="lineStyle">線の種類を指定します</param>
    '''' <remarks></remarks>
    'Public Sub SetBorder(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal bordersIndex As Integer, ByVal lineStyle As Integer)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        With _excelRange.Borders(bordersIndex)
    '            .LineStyle = lineStyle
    '        End With

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定されたセル範囲内の指定位置に指定線を引きます
    '''' </summary>
    '''' <param name="row">線を引く左上のセルの行番号を指定します</param>
    '''' <param name="col">線を引く左上のセルの列番号を指定します</param>
    '''' <param name="offsetX">線を引く行数を指定します</param>
    '''' <param name="offsetY">線を引く列数を指定します</param>
    '''' <param name="bordersIndex">線を引く位置を指定します</param>
    '''' <param name="lineStyle">線の種類を指定します</param>
    '''' <remarks></remarks>
    'Public Sub SetBorder(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal bordersIndex As BorderIdx, ByVal lineStyle As LineStyle)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        '外枠の場合は別関数
    '        If bordersIndex = BorderIdx.Sokowaku Then
    '            _excelRange.BorderAround(lineStyle)
    '        ElseIf bordersIndex = BorderIdx.Koushi Then
    '            SetBorders(row, col, offsetX, offsetY, lineStyle)
    '        Else
    '            With _excelRange.Borders(bordersIndex)
    '                .LineStyle = lineStyle
    '            End With
    '        End If

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub



    ''' <summary>
    ''' 指定されたセル範囲内を結合します
    ''' </summary>
    ''' <param name="row">結合する左上のセルの行番号を指定します</param>
    ''' <param name="col">結合する左上のセルの列番号を指定します</param>
    ''' <param name="offsetX">結合する行数を指定します</param>
    ''' <param name="offsetY">結合する列数を指定します</param>
    ''' <remarks></remarks>
    'Public Sub Marge(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        With _excelRange
    '            .HorizontalAlignment = _C_EXCEL_XLHALIGN_XLHALIGNCENTER     '横位置－中央
    '            '                .VerticalAlignment = _C_EXCEL_XLVALIGN_XLVALIGNTOP          '縦位置－上詰め
    '            .VerticalAlignment = _C_EXCEL_XLHALIGN_XLHALIGNCENTER          '縦位置－上詰め
    '        End With

    '        '「確認メッセージ非表示」に設定
    '        _excelApp.DisplayAlerts = False

    '        'セル結合実行
    '        _excelRange.Merge()

    '        '「確認メッセージ表示」に戻す
    '        _excelApp.DisplayAlerts = True

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub


    '''' <summary>
    '''' 指定した座標から指定した座標に線を引きます。
    '''' </summary>
    '''' <param name="start_x">開始位置のx座標</param>
    '''' <param name="start_y">開始位置のy座標</param>
    '''' <param name="end_x">終了位置のx座標</param>
    '''' <param name="end_y">終了位置のy座標</param>
    '''' <remarks></remarks>
    'Public Sub AddLine(ByVal start_x As Double, ByVal start_y As Double, ByVal end_x As Double, ByVal end_y As Double)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        Dim selection As Object

    '        selection = _excelWst.Shapes.AddLine(start_x, start_y, end_x, end_y)
    '        selection.Line.Weight = 3.5
    '        selection.Line.Visible = True
    '        selection.Line.ForeColor.RGB = RGB(0, 0, 0)



    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定した行（ROW）の縦真ん中に指定した長さの線を引きます。
    '''' </summary>
    '''' <param name="row">開始位置の行（真ん中に引きます）</param>
    '''' <param name="start_x">開始位置のx座標</param>
    '''' <param name="end_x">終了位置のx座標</param>
    'Public Sub AddLineRowCenter(ByVal row As Integer, ByVal start_x As Double, ByVal end_x As Double)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        Dim y As Double
    '        Dim height As Double
    '        Dim top As Double
    '        Dim selection As Object
    '        _excelRange = _excelWst.Cells(row, 1)
    '        height = _excelRange.Height
    '        top = _excelRange.Top

    '        y = top + (height / 2)


    '        selection = _excelWst.Shapes.AddLine(start_x, y, end_x, y)
    '        selection.Line.Weight = 3.5
    '        selection.Line.Visible = True
    '        selection.Line.ForeColor.RGB = RGB(0, 0, 0)


    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定したセルの背景色を変更します。
    '''' </summary>
    '''' <param name="rangeAddress">"A1"形式でセル番地を指定します</param>
    '''' <param name="color">色</param>
    '''' <remarks></remarks>
    'Public Sub SetBackColor(ByVal rangeAddress As String, ByVal color As System.Drawing.Color)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rangeAddress)

    '        _excelRange.Interior.Color = RGB(color.R, color.G, color.B)

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定したセルの背景色を変更します。
    '''' </summary>
    '''' <param name="start_row">開始セルの行</param>
    '''' <param name="start_col">開始セルの列</param>
    '''' <param name="end_row">終了セルの行</param>
    '''' <param name="end_col">終了セルの列</param>
    '''' <param name="color">色</param>
    '''' <remarks></remarks>
    'Public Sub SetBackColor(ByVal start_row As Integer, ByVal start_col As Integer, ByVal end_row As Integer, ByVal end_col As Integer, ByVal color As System.Drawing.Color)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(start_row, start_col), _excelWst.Cells(end_row, end_col))

    '        _excelRange.Interior.Color = RGB(color.R, color.G, color.B)



    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定したセルの背景色を変更します。
    '''' </summary>
    '''' <param name="row">セルの行</param>
    '''' <param name="col">セルの列</param>
    '''' <param name="color">色</param>
    '''' <remarks></remarks>
    'Public Sub SetBackColor(ByVal row As Integer, ByVal col As Integer, ByVal color As System.Drawing.Color)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Cells(row, col)

    '        _excelRange.Interior.Color = RGB(color.R, color.G, color.B)



    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 指定したセルを太字にします。
    '''' </summary>
    '''' <param name="row">セルの行</param>
    '''' <param name="col">セルの列</param>
    '''' <param name="val">True:太字にする / False:太字を解除する</param>
    '''' <remarks></remarks>
    'Public Sub SetBold(ByVal row As Integer, ByVal col As Integer, Optional ByVal val As Boolean = True)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Cells(row, col)

    '        _excelRange.Font.Bold = val

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub



    '''' <summary>
    '''' 指定したセルを太字にします。
    '''' </summary>
    '''' <param name="rangeAddress">"A1"形式でセル番地を指定します</param>
    '''' <param name="val">True:太字にする / False:太字を解除する</param>
    '''' <remarks></remarks>
    'Public Sub SetBold(ByVal rangeAddress As String, Optional ByVal val As Boolean = True)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rangeAddress)

    '        _excelRange.Font.Bold = val

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' セルの左座標を取得します
    '''' </summary>
    '''' <param name="row">セルの行</param>
    '''' <param name="col">セルの列</param>
    '''' <returns>左座標</returns>
    '''' <remarks></remarks>
    'Public Function GetCellLeft(ByVal row As Integer, ByVal col As Integer) As Object
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        Dim left As Double

    '        _excelRange = _excelWst.Cells(row, col)
    '        left = _excelRange.Left

    '        GetCellLeft = left

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Function
    '''' <summary>
    '''' セルの右座標を取得します
    '''' </summary>
    '''' <param name="row">セルの行</param>
    '''' <param name="col">セルの列</param>
    '''' <returns>右座標</returns>
    '''' <remarks></remarks>
    'Public Function GetCellRight(ByVal row As Integer, ByVal col As Integer) As Object
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If
    '        Dim left As Double

    '        Dim right As Double
    '        Dim width As Double

    '        _excelRange = _excelWst.Cells(row, col)
    '        left = _excelRange.Left
    '        width = _excelRange.Width

    '        right = left + width

    '        GetCellRight = right

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Function

    'Public Sub Protect()
    '    Try
    '        _excelWst.Protect()
    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    'Public Sub UnProtect()
    '    Try
    '        _excelWst.Unprotect()
    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 数値をExcelのカラム名に変更
    '''' </summary>
    '''' <param name="numCol">変換する数値</param>
    '''' <returns>アルファベット文字列</returns>
    '''' <remarks></remarks>
    'Public Shared Function ColToChr(ByVal numCol As Integer) As String
    '    Const BASE_COL As Integer = Asc("A") - 1

    '    If numCol <= 0 Then
    '        Return ""
    '    End If

    '    Dim divide As Integer = numCol Mod 26
    '    If divide = 0 Then
    '        divide = 26
    '    End If

    '    Dim str As String = Chr(BASE_COL + divide)
    '    If numCol = divide Then
    '        Return str
    '    End If

    '    '再帰的に呼び出し
    '    str = ColToChr((numCol - divide) / 26) + str

    '    Return str

    'End Function

    '''' <summary>
    '''' テンプレートシートの指定された行範囲をコピーして指定したシートの指定行に挿入します。
    '''' </summary>
    '''' <param name="TempName">テンプレートシート名</param>
    '''' <param name="sheetName">挿入シート名</param>
    '''' <param name="srcStartRowIndex">コピー範囲の開始行番号</param>
    '''' <param name="srcEndRowIndex">コピー範囲の終了行番号</param>
    '''' <param name="insertStartRowIndex">挿入開始行番号</param>
    '''' <param name="insertEndRowIndex">挿入終了行番号</param>
    '''' <remarks></remarks>
    'Public Sub CopyRowsByTemp(ByVal TempName As String, ByVal sheetName As String, ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertStartRowIndex As Integer, ByVal insertEndRowIndex As Integer)

    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBookオブジェクトの参照が有効ではありません")
    '        Exit Sub
    '    End If

    '    Try
    '        Me._excelWst = _excelWbk.Worksheets(TempName)
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()

    '        Me._excelWst = _excelWbk.Worksheets(sheetName)
    '        Me._excelWst.Rows(insertStartRowIndex.ToString() + ":" + insertEndRowIndex.ToString()).Insert(Shift.xlShiftDown)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("挿入エラー:" + ex.Message)

    '    End Try

    'End Sub

    '''' <summary>
    '''' 指定行の上に改ページ（横方向）を挿入する
    '''' </summary>
    '''' <param name="rowindex">指定行</param>
    '''' <remarks></remarks>
    'Public Sub AddHPageBreak(ByVal rowindex)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rowindex & ":" & rowindex)
    '        Me._excelWst.HPageBreaks.Add(_excelRange)


    '    Catch ex As Exception
    '        Throw New Exception("改ページ追加エラー:" + ex.Message)
    '    End Try
    'End Sub


    '''' <summary>
    '''' セルの値を設定します
    '''' </summary>
    '''' <param name="row">行番号を指定します</param>
    '''' <param name="col">列番号を指定します</param>
    '''' <param name="value">セルに設定する値です</param>
    '''' <remarks></remarks>
    'Public Sub SetValueSize(ByVal row As Integer, ByVal col As Integer, ByVal value As String)
    '    If Not _excelRange Is Nothing Then
    '        MRComObject(_excelRange)
    '    End If

    '    Try


    '        _excelRange = _excelWst.Cells(row, col)

    '        _excelRange.Font.Size = 14





    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "に値が書き込めません")
    '    End Try
    'End Sub


    '//下村20180517
    Public Property EnableCalculation() As Boolean
        Get
            Try
                Return DirectCast(Me._excelWst.EnableCalculation, Boolean)
            Catch
                Return False
            End Try
        End Get
        Set(value As Boolean)
            Me._excelWst.EnableCalculation = value
        End Set
    End Property

End Class

