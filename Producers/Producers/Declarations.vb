Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '�������� ��� ������������ Scala
    Public UserID As Integer                              '��� ������������ Scala
    Public FullName As String                             '��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������
    Public MyManufacturerCode As String                   '��� �������������

    Public MyManufacturersSelectList As ManufacturersSelectList '���������� ���� ������ ��������������
    Public MyEditManufacturer As EditManufacturer         '���������� ���� �������� / �������������� �������������
End Module
