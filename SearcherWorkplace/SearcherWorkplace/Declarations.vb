Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala
    Public PurchCode As String                            '��� ��������� Scala
    Public PurchName As String                            '��� ��������� Scala
    Public IsManager As Integer                           '������� - �������� �� ����������

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyRequestNum As Integer                        '����� ������� �� ����� ����������
    Public MyItemSrchID As Integer                        'ID ������ � ������������ �������
    Public ImportFileName As String                       '��� ����� ��� ������� ������
    Public MySupplierID As Integer                        '��� ���������� ��� ������� ������
    Public MySupplierCode As String                       '��� ���������� � Scala ��� ������� ������
    Public MySupplierName As String                       '�������� ���������� ��� ������� ������
    Public ExcelVersion As String                         '������ ����� ��� ������� 
    Public MyRowIndex As Integer                          'ID ������ � ������� (��� ������������ ����)

    Public MySupplierSelect As SupplierSelect             '���������� ���� ������ �����������
    Public MySupplierSelectList As SupplierSelectList     '���������� ���� ������� ��� ������ �����������
    Public MyAddItem As AddItem                           '���������� ���� ���������� ������ � ������
    Public MyItemSelect As ItemSelect                     '���������� ���� ������ ������ Scala
    Public MyItemSelectList As ItemSelectList             '���������� ���� ������ ������ �� ���������
    Public MyAttachmentsList As AttachmentsList           '���������� ���� �����������
    Public MySearcherList As SearcherList                 '���������� ���� ������ �����������
    Public MyOrdersList As OrdersList                     '���������� ���� �������� ������� 0 ����

    Public SortColumnNum As Integer                       '����� ������� � ����������� ��� ����������
    Public SortColOrder As System.ComponentModel.ListSortDirection '����������� ����������
End Module
