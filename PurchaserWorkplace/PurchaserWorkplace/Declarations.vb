Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '�������� ��� ������������ Scala
    Public UserID As Integer                              '��� ������������ Scala
    Public UserName As String                             '����� ������������ Scala
    Public FullName As String                             '��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MySupplierCode As String                       '��� ����������
    Public MyWH As String                                 '��� ������
    Public MyOrderID As String                            '����� ����������� ������
    Public MyPurchOrderSum As Double                      '����� ����������� ������

    Public MySupplierSelectList As SupplierSelectList               '���������� ���� ������ �����������
    Public MyConsolidatedOrders As ConsolidatedOrders               '���������� ���� ������������ ���������� �������
    Public MyEditConsolidatedOrder As EditConsolidatedOrder         '���������� ���� �������� / �������������� ����������� ������
    Public MyImportCommonConfirmation As ImportCommonConfirmation   '���������� ���� ������� ������������ ������������� � ��������
    Public MyErrorForm As ErrorForm                                 '���������� ���� ������ ��������� �� ������
    Public MyCreateSPTasks As CreateSPTasks                         '���������� ���� �������� ����� � ����������
    Public MySupplierInfo As SupplierInfo                           '���������� ���� ���������� �� ���������� �� ���� �������

End Module
