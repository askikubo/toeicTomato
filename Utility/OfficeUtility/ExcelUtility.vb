Imports System
Imports Microsoft.Office.Interop.Excel

Public Class ExcelUtility
    Implements IDisposable

#Region "�萔��`"
    Private Const _C_EXCEL_LINESTYLE_XLCONTINUOUS As Integer = 1        'Excel�r���|����
    Private Const _C_EXCEL_XLHALIGN_XLHALIGNCENTER As Integer = -4108   'Excel�����|��������
    Private Const _C_EXCEL_XLVALIGN_XLVALIGNTOP As Integer = -4160      'Excel�����|������l��
    Public Const _C_EXCEL_PASTE_TYPE_ALL = -4104
    Public Const _C_EXCEL_PASTE_TYPE_FORMULAS = -4123
    Public Const _C_EXCEL_PASTE_TYPE_VALUES = -4163
    Public Const _C_EXCEL_PASTE_TYPE_FORMATS = -4122

    ''' <summary>
    ''' �t�@�C���ۑ����̃t�H�[�}�b�g
    ''' </summary>
    Public Enum XlFileFormat
        ''' <summary>2003�ȑO�̏ꍇ��xls</summary>
        xlExcel9795 = 43
        ''' <summary>2007�ȍ~�̏ꍇ��xls</summary>
        xlExcel8 = 56
        ''' <summary>xls�쐬���ɁA�`2003��2007�`���V�X�e���Ŏ����U�蕪������ꍇ�͂�����g�p</summary>
        xlExcelXls = 99999
        ''' <summary>�f�t�H���g���̂܂�</summary>
        xlWorkbookNormal = -4143
    End Enum

    ''' <summary>
    ''' ���̎��
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum LineStyle
        ''' <summary>����</summary>
        xlContinuous = 1
        ''' <summary>������</summary>
        xlLineStyleNone = -4142
    End Enum

    ''' <summary>
    ''' �r���������ʒu
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum BorderIdx
        ''' <summary>�Z���̍��̐�</summary>
        xlEdgeLeft = 7
        ''' <summary>�Z���̏�̐�</summary>
        xlEdgeTop = 8
        ''' <summary>�Z���̉��̐�</summary>
        xlEdgeBottom = 9
        ''' <summary>�Z���̉E�̐�</summary>
        xlEdgeRight = 10
        ''' <summary>�Z���͈͂̓����̉���</summary>
        xlInsideVertical = 11
        ''' <summary>�Z���͈͂̓����̏c��</summary>
        xlInsideHorizontal = 12
        ''' <summary>�O�g</summary>
        Sokowaku = -1
        ''' <summary>�i�q</summary>
        Koushi = -2
    End Enum

    ''' <summary>
    ''' �Z���̈ړ�����
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum Shift
        ''' <summary>��</summary>
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

    '''' <summary>�V�[�g��</summary>
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
    ''' Excel�̃o�[�W����
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
    ''' �R���X�g���N�^
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' �t�@�C�����J���܂�
    ''' </summary>
    ''' <param name="fileFullName">�J���t�@�C����FullName�Ŏw�肵�܂�</param>
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
            Throw New NotOpenException(fileFullName + "�̃I�[�v���Ɏ��s���܂���")
        End Try
    End Sub

    '''' <summary>
    '''' ���삷��V�[�g��ݒ肵�܂�
    '''' </summary>
    '''' <param name="sheetName">�V�[�g��</param>
    '''' <remarks></remarks>
    'Public Sub SetWorksheet(ByVal sheetName As String)
    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBook�I�u�W�F�N�g�̎Q�Ƃ��L���ł͂���܂���")
    '        Exit Sub
    '    End If

    '    Try

    '        _excelWst = _excelWbk.Worksheets(sheetName)
    '    Catch ex As Exception
    '        Throw New NotValueException("Worksheets(" + sheetName + ")�̐ݒ�Ɏ��s���܂���")
    '    End Try

    '    GC.Collect()
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ�V�[�g���A�N�e�B�u�ɂ��܂��B
    '''' </summary>
    '''' <param name="index">�A�N�e�B�u�ɂ���V�[�g�ԍ��ł�</param>
    '''' <remarks>index >= 1�@�ł�</remarks>
    'Public Sub SetActiveSheet(ByVal index As Integer)
    '    Me._excelWbk.Sheets(index).Select()
    '    Me.SetWorksheet(index)
    'End Sub


    ''' <summary>
    ''' ���삷��V�[�g��ݒ肵�܂�
    ''' </summary>
    ''' <param name="sheetIndex">�V�[�g�ԍ�</param>
    ''' <remarks></remarks>
    Public Sub SetWorksheet(ByVal sheetIndex As Integer)
        If _excelWbk Is Nothing Then
            Throw New NotReferException("WorkBook�I�u�W�F�N�g�̎Q�Ƃ��L���ł͂���܂���")
            Exit Sub
        End If

        If Not _excelWst Is Nothing Then
            MRComObject(_excelWst)
        End If


        Try
            _excelWst = _excelWbk.Worksheets(sheetIndex)
        Catch ex As Exception
            Throw New NotValueException("Worksheets(" + sheetIndex.ToString + ")�̐ݒ�Ɏ��s���܂���")
        End Try
        GC.Collect()
    End Sub

    '''' <summary>
    '''' ���삵�Ă���V�[�g����ԍŌ�ɃR�s�[���܂��B
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub CopyWorksheet()
    '    Try
    '        Me._excelApp.ScreenUpdating = False
    '        Dim sheetCnt As Integer = _excelWbk.Worksheets.Count()
    '        _excelWst.Copy(, Me._excelWbk.Sheets(sheetCnt))
    '    Catch ex As Exception
    '        Throw New Exception("�R�s�[���s:" + ex.Message)
    '    Finally
    '        Me._excelApp.ScreenUpdating = True
    '    End Try
    'End Sub


    '''' <summary>
    '''' ���삵�Ă���V�[�g���w�肳�ꂽ�ʒu�̌��ɃR�s�[���܂��B
    '''' </summary>
    '''' <param name="decIndex">���̃V�[�g�ԍ��̌��ɃR�s�[����܂�</param>
    '''' <remarks></remarks>
    'Public Sub CopyWorksheet(ByVal decIndex As Integer)
    '    Try
    '        Me._excelApp.ScreenUpdating = False
    '        Dim sheetCnt As Integer = _excelWbk.Worksheets.Count()
    '        _excelWst.Copy(, Me._excelWbk.Sheets(decIndex))
    '    Catch ex As Exception
    '        Throw New Exception("�R�s�[���s:" + ex.Message)
    '    Finally
    '        Me._excelApp.ScreenUpdating = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' ���삵�Ă���V�[�g�̖��O��ύX���܂��B
    '''' </summary>
    '''' <param name="name">�V�����V�[�g�����w�肵�܂�</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal name As String)
    '    Try
    '        Me._excelWst.Name = name
    '    Catch ex As Exception
    '        Throw New Exception("���l�[�����s:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ�V�[�g�ԍ��̃V�[�g����ύX���܂��B
    '''' </summary>
    '''' <param name="index">���O��ς���V�[�g�̃V�[�g�ԍ����w�肵�܂��B</param>
    '''' <param name="name">�V�����V�[�g�����w�肵�܂��B</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal index As Integer, ByVal name As String)
    '    Try

    '        Me._excelWbk.Sheets(index).Name = name
    '    Catch ex As Exception
    '        Throw New Exception("���l�[�����s:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ�V�[�g�ԍ��̃V�[�g���폜���܂��B
    '''' </summary>
    '''' <param name="index">�폜����V�[�g�̃V�[�g�ԍ�</param>
    '''' <remarks></remarks>
    'Public Sub DeleteSheet(ByVal index As Integer)
    '    Try
    '        Me._excelApp.DisplayAlerts = False
    '        Me._excelWbk.Sheets(index).Delete()
    '    Catch ex As Exception
    '        Throw New Exception("�폜���s:" + ex.Message)
    '    Finally
    '        Me._excelApp.DisplayAlerts = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ�V�[�g�ԍ��̃V�[�g���폜���܂��B
    '''' </summary>
    '''' <param name="sheetName">�폜����V�[�g�̖���</param>
    '''' <remarks></remarks>
    'Public Sub DeleteSheet(ByVal sheetName As String)
    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBook�I�u�W�F�N�g�̎Q�Ƃ��L���ł͂���܂���")
    '        Exit Sub
    '    End If

    '    Try
    '        Me._excelApp.DisplayAlerts = False
    '        _excelWst = _excelWbk.Worksheets(sheetName)
    '        Me._excelWst.Delete()
    '    Catch ex As Exception
    '        Throw New Exception("�폜���s:" + ex.Message)
    '    Finally
    '        Me._excelApp.DisplayAlerts = True
    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ����폜���܂��B
    '''' </summary>
    '''' <param name="delstr">�폜�����̃A���t�@�x�b�g���w�肵�܂��B</param>
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

    '        Throw New Exception("�폜�G���[:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ����폜���܂��B
    '''' </summary>
    '''' <param name="startColStr">�폜�����̊J�n�A���t�@�x�b�g���w�肵�܂��B</param>
    '''' ''' <param name="endColStr">�폜�����̏I���A���t�@�x�b�g���w�肵�܂��B</param>
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

    '        Throw New Exception("�폜�G���[:" + ex.Message)

    '    End Try
    'End Sub

    ''' <summary>
    ''' �w�肳�ꂽ�s���폜���܂��B
    ''' </summary>
    ''' <param name="rowIndex">�폜�������s</param>
    ''' <remarks></remarks>
    Public Sub DeleteRow(ByVal rowIndex As Integer)
        DeleteRows(rowIndex, rowIndex)
    End Sub

    ''' <summary>
    ''' �w�肳�ꂽ�s���폜���܂��B
    ''' </summary>
    ''' <param name="startRowIndex">�폜�������J�n�s</param>
    ''' ''' <param name="endRowIndex">�폜�������I���s</param>
    ''' <remarks></remarks>
    Public Sub DeleteRows(ByVal startRowIndex As Integer, ByVal endRowIndex As Integer)

        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Range(startRowIndex & ":" & endRowIndex)
            Me._excelRange.delete()

        Catch ex As Exception

            Throw New Exception("�폜�G���[:" + ex.Message)

        End Try
    End Sub

    '''' <summary>
    '''' �w�肳�ꂽ���O�̃V�[�g����ύX���܂��B
    '''' </summary>
    '''' <param name="oldname">�ύX����V�[�g�̖��O�ł��B</param>
    '''' <param name="name">�V�����V�[�g���ł��B</param>
    '''' <remarks></remarks>
    'Public Sub RenameSheetName(ByVal oldname As String, ByVal name As String)
    '    Try
    '        Me._excelWbk.Sheets(oldname).Name = name
    '    Catch ex As Exception
    '        Throw New Exception("���l�[�����s:" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' �w�肳�ꂽ�s���R�s�[���đ}�����܂��B
    '''' </summary>
    '''' <param name="srcRowIndex">�R�s�[����s�ԍ��ł��B</param>
    '''' <param name="count">���s�}�����邩���w�肵�܂��B</param>
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

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try
    'End Sub
    ''' <summary>
    ''' �w�肳�ꂽ�s�͈͂��R�s�[���đ}�����܂��B
    ''' </summary>
    ''' <param name="srcStartRowIndex">�R�s�[�͈͂̊J�n�s�ԍ��ł��B</param>
    ''' <param name="srcEndRowIndex">�R�s�[�͈͂̏I���s�ԍ��ł��B</param>
    ''' <param name="insertRowIndex">�}���s�ԍ��ł��B</param>
    ''' <remarks></remarks>
    'Public Sub CopyRow(ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertRowIndex As Integer)
    '    Try
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()
    '        Me._excelWst.Rows(insertRowIndex).Insert(-4121)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try
    'End Sub
    ''' <summary>
    ''' �w�肳�ꂽ�s�͈͂��R�s�[���đ}�����܂��B
    ''' </summary>
    ''' <param name="srcStartRowIndex">�R�s�[�͈͂̊J�n�s�ԍ��ł��B</param>
    ''' <param name="srcEndRowIndex">�R�s�[�͈͂̏I���s�ԍ��ł��B</param>
    ''' <param name="insertStartRowIndex">�}����̊J�n�s�ԍ��ł��B</param>
    ''' <param name="insertEndRowIndex">�}����̏I���s�ԍ��ł��B</param>
    ''' <remarks></remarks>
    'Public Sub CopyRow(ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertStartRowIndex As Integer, ByVal insertEndRowIndex As Integer)
    '    Try
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()
    '        Me._excelWst.Rows(insertStartRowIndex.ToString() + ":" + insertEndRowIndex.ToString()).Insert(-4121)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try
    'End Sub
    '''' <summary>
    '''' �w�肳�ꂽ����R�s�[���đ}�����܂�
    '''' </summary>
    '''' <param name="srcColIndex">�R�s�[�����ԍ��ł��B</param>
    '''' <param name="count">����}�����邩���w�肵�܂��B</param>
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

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try
    'End Sub
    '''' <summary>
    '''' �w�肳�ꂽ����R�s�[���đ}�����܂�
    '''' </summary>
    '''' <param name="srcRangeAddress">�R�s�[���̗��͈͂Ŏw��</param>
    '''' <param name="decRangeAddress">�\��t����̗��͈͂Ŏw��</param>
    '''' <remarks></remarks>
    'Public Sub CopyCol(ByVal srcRangeAddress As String, ByVal decRangeAddress As String)

    '    Try
    '        Me._excelWst.Columns(srcRangeAddress).Copy()
    '        Me._excelWst.Columns(decRangeAddress).Resize(1).Insert(-4161)

    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' �Z�����R�s�[���ē\��t���܂��B
    '''' </summary>
    '''' <param name="srcRangeAddress">�R�s�[���̃Z����͈͂Ŏw��</param>
    '''' <param name="decRangeAddress">�R�s�[��̃Z����͈͂Ŏw��</param>
    '''' <remarks></remarks>
    'Public Sub CopyCells(ByVal srcRangeAddress As String, ByVal decRangeAddress As String, ByVal pasteType As Integer)

    '    Try

    '        Me._excelWst.Range(srcRangeAddress).Copy()
    '        Me._excelWst.Range(decRangeAddress).PasteSpecial(pasteType)

    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("�R�s�[�G���[:" + ex.Message)

    '    End Try
    'End Sub

    '''' <summary>
    '''' �Z����I�����܂�
    '''' </summary>
    '''' <param name="rangeAddress">"A1"�`���̃Z���Ԓn���w�肵�܂�</param>
    '''' <remarks></remarks>
    'Public Sub SelectCell(ByVal rangeAddress As String)
    '    If Not _excelRange Is Nothing Then
    '        MRComObject(_excelRange)
    '    End If

    '    _excelRange = _excelWst.Range(rangeAddress)
    '    _excelRange.Select()
    'End Sub

    '''' <summary>
    '''' �Z���̒l��ݒ肵�܂��B
    '''' </summary>
    '''' <param name="rangeAddress">"A1"�`���̃Z���Ԓn���w�肵�܂�</param>
    '''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
    '''' <remarks></remarks>
    'Public Sub SetValue(ByVal rangeAddress As String, ByVal value As String)
    '    If Not _excelRange Is Nothing Then
    '        MRComObject(_excelRange)
    '    End If

    '    Try

    '        _excelRange = _excelWst.Range(rangeAddress)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException("Range(" + rangeAddress + ")�ɒl���������߂܂���")
    '    End Try
    'End Sub

    ''' <summary>
    ''' �Z���̒l��ݒ肵�܂�
    ''' </summary>
    ''' <param name="row">�s�ԍ����w�肵�܂�</param>
    ''' <param name="col">��ԍ����w�肵�܂�</param>
    ''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal row As Integer, ByVal col As Integer, ByVal value As String)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try
            _excelRange = _excelWst.Cells(row, col)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "�ɒl���������߂܂���")
        End Try
    End Sub

    ''' <summary>
    ''' �Z���̒l��ݒ肵�܂�.
    ''' �����̃T�C�Y���w�肵�܂��B
    ''' </summary>
    ''' <param name="row">�s�ԍ����w�肵�܂�</param>
    ''' <param name="col">��ԍ����w�肵�܂�</param>
    ''' 
    ''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
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
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "�ɒl���������߂܂���")
        End Try
    End Sub

    ''' <summary>
    ''' �Z���̒l��ݒ肵�܂��B
    ''' </summary>
    ''' <param name="rangeAddress">"A1"�`���̃Z���Ԓn���w�肵�܂�</param>
    ''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal rangeAddress As String, ByVal value As Object)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try

            _excelRange = _excelWst.Range(rangeAddress)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Range(" + rangeAddress + ")�ɒl���������߂܂���")
        End Try
    End Sub

    ''' <summary>
    ''' �Z���̒l��ݒ肵�܂�
    ''' </summary>
    ''' <param name="row">�s�ԍ����w�肵�܂�</param>
    ''' <param name="col">��ԍ����w�肵�܂�</param>
    ''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
    ''' <remarks></remarks>
    Public Sub SetValue(ByVal row As Integer, ByVal col As Integer, ByVal value As Object)
        If Not _excelRange Is Nothing Then
            MRComObject(_excelRange)
        End If

        Try
            _excelRange = _excelWst.Cells(row, col)
            _excelRange.Value = value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "�ɒl���������߂܂���")
        End Try
    End Sub

    ''' <summary>
    ''' �Z���̒l���擾���܂�
    ''' </summary>
    ''' <param name="rangeAddress">"A1"�`���ŃZ���Ԓn���w�肵�܂�</param>
    ''' <returns>�Z���̒l�i������`���j</returns>
    ''' <remarks></remarks>
    'Public Function GetValue(ByVal rangeAddress As String) As String
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rangeAddress)
    '        Return _excelRange.Value
    '    Catch ex As Exception
    '        Throw New NotValueException("Range(" + rangeAddress + ")�̒l���ǂݍ��߂܂���")
    '    End Try

    'End Function

    '''' <summary>
    '''' �Z���̒l���擾���܂�
    '''' </summary>
    '''' <param name="row">�s�ԍ����w�肵�܂�</param>
    '''' <param name="col">��ԍ����w�肵�܂�</param>
    '''' <returns>�Z���̒l�i������`���j</returns>
    '''' <remarks></remarks>
    Public Function GetValue(ByVal row As Integer, ByVal col As Integer) As String
        Try
            If Not _excelRange Is Nothing Then
                MRComObject(_excelRange)
            End If

            _excelRange = _excelWst.Cells(row, col)

            Return _excelRange.Value
        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")�̒l���ǂݍ��߂܂���")
        End Try
    End Function

    '''' <summary>
    '''' �Z���̒l�����H���Ď擾���܂�
    '''' </summary>
    '''' <param name="row">�s�ԍ����w�肵�܂�</param>
    '''' <param name="col">��ԍ����w�肵�܂�</param>
    '''' <returns>�Z���̒l�i������`���j</returns>
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

            '�G�t�F�N�g��������
            For i = 0 To strTarget.Length - 1

                Dim targetChar As Object = _excelRange.Characters(1 + i, 1)
                Dim targetNextChar As Object = _excelRange.Characters(2 + i, 1)

                '�����擾
                If Asc(targetChar.Text) >= -30561 Then
                    strKanjiValue += strTarget(i)

                    If Asc(targetNextChar.Text) < -30561 Then
                        txtKanjiList.Add(strKanjiValue)
                        strKanjiValue = ""
                    End If
                End If


                '�������擾
                If targetChar.Font.Bold = True Then

                    strBoldValue += strTarget(i)

                    If targetNextChar.Font.Bold = False Then
                        txtBoldList.Add(strBoldValue)
                        strBoldValue = ""
                    End If

                End If

                '�����擾
                If targetChar.Font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle Then
                    strUnderlineValue += strTarget(i)
                    If targetNextChar.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone Then
                        txtUnderlineList.Add(strUnderlineValue)
                        strUnderlineValue = ""
                    End If
                End If

                '�ԕ����擾
                If targetChar.Font.Color = colorRed Then
                    strRedColorValue += strTarget(i)
                    If targetNextChar.Font.Color = False Then
                        txtRedColorList.Add(strRedColorValue)
                        strRedColorValue = ""
                    End If
                End If


            Next

            '���r��ǉ����Ă��Ƃɂ���đΏە�����̊Ԋu�������Ƃ��܂��K�p����Ȃ�
            '�����ϊ�
            For i = 0 To txtUnderlineList.Count - 1
                strTarget = strTarget.Replace(txtUnderlineList(i), "<u style = ""text-decoration-color: black"">" + txtUnderlineList(i) + "</u>")
            Next

            '�������ϊ�
            For i = 0 To txtBoldList.Count - 1
                strTarget = strTarget.Replace(txtBoldList(i), "<strong>" + txtBoldList(i) + "</strong>")
            Next

            '�ԕ����ϊ�
            For i = 0 To txtRedColorList.Count - 1
                strTarget = strTarget.Replace(txtRedColorList(i), "<span style = ""color:red"">" + txtRedColorList(i) + "</span>")
            Next

            '���r�ϊ�
            For i = 0 To txtKanjiList.Count - 1
                strTarget = strTarget.Replace(txtKanjiList(i), "<ruby>" + txtKanjiList(i))
            Next


            strTarget = strTarget.Replace("�i", "<rt>").Replace("�j", "</rt></ruby>")

            Return strTarget


        Catch ex As Exception
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")�̒l���ǂݍ��߂܂���")
        End Try
    End Function



    '''' <summary>
    '''' A��̍s�����擾���܂�
    '''' </summary>
    '''' <returns>A��̍s��</returns>
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
            Throw New NotValueException("�s�����擾�ł��܂���")
        End Try
    End Function

    '''' <summary>
    '''' �w���̍s�����擾���܂�
    '''' </summary>
    '''' <param name="colStr">A-Z�\�L�̗�</param>
    '''' <returns>�w���̍s��</returns>
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
            Throw New NotValueException("�s�����擾�ł��܂���")
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
            Throw New NotValueException("�s�����擾�ł��܂���")
        End Try
    End Function

    '''' <summary>
    '''' �����͈͂̃Z���̒l��2�����z��ňꊇ�ݒ肵�܂�
    '''' </summary>
    '''' <param name="baseRangeAddress">�l��ݒ肷�鍶��̃Z�����w�肵�܂�</param>
    '''' <param name="value">�ݒ肷��2�����z��ł��B</param>
    '''' <remarks></remarks>
    'Public Sub SetValues(ByVal baseRangeAddress As String, ByVal value As Object(,))
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If


    '        _excelRange = _excelWst.Range(baseRangeAddress).Resize(value.GetUpperBound(0) + 1, value.GetUpperBound(1) + 1)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException(baseRangeAddress + "�ɏ������߂܂���" + ex.Message)
    '    End Try
    'End Sub

    '''' <summary>
    '''' �����͈͂̃Z���̒l��2�����z��ňꊇ�ݒ肵�܂�
    '''' </summary>
    '''' <param name="row">�l��ݒ肷�鍶��̃Z���̍s�ԍ����w�肵�܂�</param>
    '''' <param name="col">�l��ݒ肷�鍶��̃Z���̗�ԍ����w�肵�܂�</param>
    '''' <param name="value">�ݒ肷��2�����z��ł��B</param>
    '''' <remarks></remarks>
    'Public Sub SetValues(ByVal row As Integer, ByVal col As Integer, ByVal value As Object(,))
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Cells(row, col).Resize(value.GetUpperBound(0) + 1, value.GetUpperBound(1) + 1)
    '        _excelRange.Value = value
    '    Catch ex As Exception
    '        Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "�ɒl���������߂܂���")
    '    End Try
    'End Sub


    '''' <summary>
    '''' �w�肳�ꂽ�Z���͈͂̒l��2�����z��Ŏ擾���܂��i������^�j
    '''' </summary>
    '''' <param name="rangeAddress">"A1"�`���Ŏ擾����Z���͈͂��w�肵�܂�</param>
    '''' <returns>�Z���̒l�ł��i2�����z��j</returns>
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
    '        Throw New NotValueException(rangeAddress + "����ǂݍ��߂܂���")
    '    End Try
    'End Function

    '''' <summary>
    '''' �w�肳�ꂽ�Z���͈͂̒l��2�����z��Ŏ擾���܂��i������^�j
    '''' </summary>
    '''' <param name="row">�l���擾���鍶��̃Z���̍s�ԍ����w�肵�܂�</param>
    '''' <param name="col">�l���擾���鍶��̃Z���̗�ԍ����w�肵�܂�</param>
    '''' <param name="offsetX">�擾����s�����w�肵�܂�</param>
    '''' <param name="offsetY">�擾����񐔂��w�肵�܂�</param>
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
            Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + ")����ǂݍ��߂܂���")
        End Try
    End Function

    ''' <summary>
    ''' ������s��
    ''' </summary>
    ''' <param name="printName">�v�����^�[��</param>
    ''' <param name="count">�������</param>
    ''' <remarks></remarks>
    Public Sub Print(ByVal printName As String, ByVal count As Integer)
        Try
            Me._excelWst.PrintOut(1, 100, count, False, printName, Type.Missing, Type.Missing, Type.Missing)
        Catch ex As Exception
            Throw New Exception("������s:" + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' �ۑ����܂�
    ''' </summary>
    ''' <param name="fileFullName">�ۑ��t�@�C�������w�肵�܂�</param>
    ''' <param name="Alret">Excel����̊m�F���b�Z�[�W�̕\��/��\�����w�肵�܂�</param>
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
    ''' �ۑ��t�@�C���`�����w�肵�ĕۑ����܂��B
    ''' �e���v���[�g�t�@�C���ƕۑ��t�@�C���̊g���q���Ⴄ�ꍇ�ɗ��p���܂��B
    ''' </summary>
    ''' <param name="fileFullName">�ۑ��t�@�C�������w�肵�܂�</param>
    ''' <param name="format">�ۑ��t�@�C���`�����w�肵�܂�</param>
    ''' <remarks></remarks>
    Public Sub SaveAs(ByVal fileFullName As String, ByVal format As XlFileFormat)
        Try
            '�ۑ��_�C�A���O��\��
            _excelApp.DisplayAlerts = False

            Dim wkFormat As XlFileFormat = format

            'xls��U�蕪���ēo�^����ꍇ
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
    ''' �t�@�C������܂��B
    ''' </summary>
    ''' <param name="save">�ۑ����ĕ��邩�ǂ������w�肵�܂�</param>
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
    '''' ����͈͂�ݒ肵�܂��B
    '''' </summary>
    '''' <param name="startcell">�J�n�Z��</param>
    '''' <param name="endcell">�I���Z��</param>
    '''' <remarks></remarks>
    'Public Sub SetPrintArea(ByVal startcell As String, ByVal endcell As String)
    '    _excelWst.PageSetup.PrintArea = startcell & ":" & endcell
    'End Sub

    '''' <summary>
    '''' ����͈͂��N���A���܂��B
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub SetPrintAreaClear()
    '    _excelWst.PageSetup.PrintArea = ""
    'End Sub



    ''' <summary>
    ''' �I�������ł�
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
    '''' Excel�̉�ʂւ̕\��/��\����ݒ肵�܂�
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
    ''' �t�@�C���I�[�v���Ɏ��s�����ꍇ�̗�O
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
        'COM �I�u�W�F�N�g�̎g�p��A�����I�� COM �I�u�W�F�N�g�ւ̎Q�Ƃ��������
        Try
            '�񋟂��ꂽ�����^�C���Ăяo���\���b�p�[�̎Q�ƃJ�E���g���f�N�������g���܂�
            If Not objCom Is Nothing AndAlso System.Runtime.InteropServices. _
                                                      Marshal.IsComObject(objCom) Then
                Dim I As Integer
                Do
                    I = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                Loop Until I <= 0
            End If
        Catch
        Finally
            '�Q�Ƃ���������
            objCom = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' �w�肳�ꂽ�Z���͈͓��ɋK����������܂�
    ''' </summary>
    ''' <param name="row">������������̃Z���̍s�ԍ����w�肵�܂�</param>
    ''' <param name="col">������������̃Z���̗�ԍ����w�肵�܂�</param>
    ''' <param name="offsetX">���������s�����w�肵�܂�</param>
    ''' <param name="offsetY">���������񐔂��w�肵�܂�</param>
    ''' <param name="lineStyle">���̎�ނ��w�肵�܂�</param>
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
    '''' �w�肳�ꂽ�Z���͈͓��̎w��ʒu�Ɏw����������܂�
    '''' </summary>
    '''' <param name="row">������������̃Z���̍s�ԍ����w�肵�܂�</param>
    '''' <param name="col">������������̃Z���̗�ԍ����w�肵�܂�</param>
    '''' <param name="offsetX">���������s�����w�肵�܂�</param>
    '''' <param name="offsetY">���������񐔂��w�肵�܂�</param>
    '''' <param name="bordersIndex">���������ʒu���w�肵�܂�</param>
    '''' <param name="lineStyle">���̎�ނ��w�肵�܂�</param>
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
    '''' �w�肳�ꂽ�Z���͈͓��̎w��ʒu�Ɏw����������܂�
    '''' </summary>
    '''' <param name="row">������������̃Z���̍s�ԍ����w�肵�܂�</param>
    '''' <param name="col">������������̃Z���̗�ԍ����w�肵�܂�</param>
    '''' <param name="offsetX">���������s�����w�肵�܂�</param>
    '''' <param name="offsetY">���������񐔂��w�肵�܂�</param>
    '''' <param name="bordersIndex">���������ʒu���w�肵�܂�</param>
    '''' <param name="lineStyle">���̎�ނ��w�肵�܂�</param>
    '''' <remarks></remarks>
    'Public Sub SetBorder(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer, ByVal bordersIndex As BorderIdx, ByVal lineStyle As LineStyle)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        '�O�g�̏ꍇ�͕ʊ֐�
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
    ''' �w�肳�ꂽ�Z���͈͓����������܂�
    ''' </summary>
    ''' <param name="row">�������鍶��̃Z���̍s�ԍ����w�肵�܂�</param>
    ''' <param name="col">�������鍶��̃Z���̗�ԍ����w�肵�܂�</param>
    ''' <param name="offsetX">��������s�����w�肵�܂�</param>
    ''' <param name="offsetY">��������񐔂��w�肵�܂�</param>
    ''' <remarks></remarks>
    'Public Sub Marge(ByVal row As Integer, ByVal col As Integer, ByVal offsetX As Integer, ByVal offsetY As Integer)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(_excelWst.Cells(row, col), _excelWst.Cells(row + offsetX, col + offsetY))

    '        With _excelRange
    '            .HorizontalAlignment = _C_EXCEL_XLHALIGN_XLHALIGNCENTER     '���ʒu�|����
    '            '                .VerticalAlignment = _C_EXCEL_XLVALIGN_XLVALIGNTOP          '�c�ʒu�|��l��
    '            .VerticalAlignment = _C_EXCEL_XLHALIGN_XLHALIGNCENTER          '�c�ʒu�|��l��
    '        End With

    '        '�u�m�F���b�Z�[�W��\���v�ɐݒ�
    '        _excelApp.DisplayAlerts = False

    '        '�Z���������s
    '        _excelRange.Merge()

    '        '�u�m�F���b�Z�[�W�\���v�ɖ߂�
    '        _excelApp.DisplayAlerts = True

    '    Catch ex As Exception
    '        Throw New NotValueException(ex.Message)
    '    End Try
    'End Sub


    '''' <summary>
    '''' �w�肵�����W����w�肵�����W�ɐ��������܂��B
    '''' </summary>
    '''' <param name="start_x">�J�n�ʒu��x���W</param>
    '''' <param name="start_y">�J�n�ʒu��y���W</param>
    '''' <param name="end_x">�I���ʒu��x���W</param>
    '''' <param name="end_y">�I���ʒu��y���W</param>
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
    '''' �w�肵���s�iROW�j�̏c�^�񒆂Ɏw�肵�������̐��������܂��B
    '''' </summary>
    '''' <param name="row">�J�n�ʒu�̍s�i�^�񒆂Ɉ����܂��j</param>
    '''' <param name="start_x">�J�n�ʒu��x���W</param>
    '''' <param name="end_x">�I���ʒu��x���W</param>
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
    '''' �w�肵���Z���̔w�i�F��ύX���܂��B
    '''' </summary>
    '''' <param name="rangeAddress">"A1"�`���ŃZ���Ԓn���w�肵�܂�</param>
    '''' <param name="color">�F</param>
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
    '''' �w�肵���Z���̔w�i�F��ύX���܂��B
    '''' </summary>
    '''' <param name="start_row">�J�n�Z���̍s</param>
    '''' <param name="start_col">�J�n�Z���̗�</param>
    '''' <param name="end_row">�I���Z���̍s</param>
    '''' <param name="end_col">�I���Z���̗�</param>
    '''' <param name="color">�F</param>
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
    '''' �w�肵���Z���̔w�i�F��ύX���܂��B
    '''' </summary>
    '''' <param name="row">�Z���̍s</param>
    '''' <param name="col">�Z���̗�</param>
    '''' <param name="color">�F</param>
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
    '''' �w�肵���Z���𑾎��ɂ��܂��B
    '''' </summary>
    '''' <param name="row">�Z���̍s</param>
    '''' <param name="col">�Z���̗�</param>
    '''' <param name="val">True:�����ɂ��� / False:��������������</param>
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
    '''' �w�肵���Z���𑾎��ɂ��܂��B
    '''' </summary>
    '''' <param name="rangeAddress">"A1"�`���ŃZ���Ԓn���w�肵�܂�</param>
    '''' <param name="val">True:�����ɂ��� / False:��������������</param>
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
    '''' �Z���̍����W���擾���܂�
    '''' </summary>
    '''' <param name="row">�Z���̍s</param>
    '''' <param name="col">�Z���̗�</param>
    '''' <returns>�����W</returns>
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
    '''' �Z���̉E���W���擾���܂�
    '''' </summary>
    '''' <param name="row">�Z���̍s</param>
    '''' <param name="col">�Z���̗�</param>
    '''' <returns>�E���W</returns>
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
    '''' ���l��Excel�̃J�������ɕύX
    '''' </summary>
    '''' <param name="numCol">�ϊ����鐔�l</param>
    '''' <returns>�A���t�@�x�b�g������</returns>
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

    '    '�ċA�I�ɌĂяo��
    '    str = ColToChr((numCol - divide) / 26) + str

    '    Return str

    'End Function

    '''' <summary>
    '''' �e���v���[�g�V�[�g�̎w�肳�ꂽ�s�͈͂��R�s�[���Ďw�肵���V�[�g�̎w��s�ɑ}�����܂��B
    '''' </summary>
    '''' <param name="TempName">�e���v���[�g�V�[�g��</param>
    '''' <param name="sheetName">�}���V�[�g��</param>
    '''' <param name="srcStartRowIndex">�R�s�[�͈͂̊J�n�s�ԍ�</param>
    '''' <param name="srcEndRowIndex">�R�s�[�͈͂̏I���s�ԍ�</param>
    '''' <param name="insertStartRowIndex">�}���J�n�s�ԍ�</param>
    '''' <param name="insertEndRowIndex">�}���I���s�ԍ�</param>
    '''' <remarks></remarks>
    'Public Sub CopyRowsByTemp(ByVal TempName As String, ByVal sheetName As String, ByVal srcStartRowIndex As Integer, ByVal srcEndRowIndex As Integer, ByVal insertStartRowIndex As Integer, ByVal insertEndRowIndex As Integer)

    '    If _excelWbk Is Nothing Then
    '        Throw New NotReferException("WorkBook�I�u�W�F�N�g�̎Q�Ƃ��L���ł͂���܂���")
    '        Exit Sub
    '    End If

    '    Try
    '        Me._excelWst = _excelWbk.Worksheets(TempName)
    '        Me._excelWst.Rows(srcStartRowIndex.ToString() + ":" + srcEndRowIndex.ToString()).Copy()

    '        Me._excelWst = _excelWbk.Worksheets(sheetName)
    '        Me._excelWst.Rows(insertStartRowIndex.ToString() + ":" + insertEndRowIndex.ToString()).Insert(Shift.xlShiftDown)
    '        Me._excelApp.CutCopyMode = False

    '    Catch ex As Exception

    '        Throw New Exception("�}���G���[:" + ex.Message)

    '    End Try

    'End Sub

    '''' <summary>
    '''' �w��s�̏�ɉ��y�[�W�i�������j��}������
    '''' </summary>
    '''' <param name="rowindex">�w��s</param>
    '''' <remarks></remarks>
    'Public Sub AddHPageBreak(ByVal rowindex)
    '    Try
    '        If Not _excelRange Is Nothing Then
    '            MRComObject(_excelRange)
    '        End If

    '        _excelRange = _excelWst.Range(rowindex & ":" & rowindex)
    '        Me._excelWst.HPageBreaks.Add(_excelRange)


    '    Catch ex As Exception
    '        Throw New Exception("���y�[�W�ǉ��G���[:" + ex.Message)
    '    End Try
    'End Sub


    '''' <summary>
    '''' �Z���̒l��ݒ肵�܂�
    '''' </summary>
    '''' <param name="row">�s�ԍ����w�肵�܂�</param>
    '''' <param name="col">��ԍ����w�肵�܂�</param>
    '''' <param name="value">�Z���ɐݒ肷��l�ł�</param>
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
    '        Throw New NotValueException("Cell(" + row.ToString + "," + col.ToString + "�ɒl���������߂܂���")
    '    End Try
    'End Sub


    '//����20180517
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

