Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyImportPriceShort As ImportPriceShort         '���������� ���� ������������ �������
    Public MyImportPriceFull As ImportPriceFull           '���������� ���� ������� �������
    Public MyErrorMessage As ErrorMessage                 '���������� ���� � ���������� �� �������
End Module
