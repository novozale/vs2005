Module Declarations

    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MySuccess As Boolean                           '���������� ���������� ��������
    Public LoadFlag As Integer                            '�������� ����� �� ��������� (0) ��� ��������� (1)

    Public MyAddPriceValue As AddPriceValue               '���������� ���� ���������� �������� ����� - �����
    Public MyEditPriceValue As EditPriceValue             '���������� ���� �������������� �������� ����� - �����
    Public MyShipmentCost As ShipmentCost                 '���������� ���� ����� ����� ��������� �������� (������)
    Public MyAddShipmentCost As AddShipmentCost           '���������� ���� ���������� ��������� ��������
    Public MySupplierSelect As SupplierSelect             '���������� ���� ������ �����������
    Public MySupplierSelectList As SupplierSelectList     '���������� ���� ������� ��� ������ �����������
    Public MyAddSInvoice As AddSInvoice                   '���������� ���� ���������� ������� �� �������
    Public MyAddPInvoice As AddPInvoice                   '���������� ���� ���������� ������� �� �������
    Public MyAddRelOrder As AddRelOrder                   '���������� ���� ���������� ������� �� �������

    Public Destination As String                          '����� ���������� (��� "������� �� �������" �� 100 ����������  �� ������� ��� ���� PriceType �.�. 1)
    Public PriceType As Integer                           '��� ������ 0 �������������, 1 - �� 100 ���������� .
    Public PriceFrom As Double                            '�� ������ �������� ����� (��������, �� 0 ��)�� �������
    Public PriceTo As Double                              '�� ������ �������� ����� (��������, �� 100 ��)
    Public PriceVal As Double                             '���������� �������� ������
    Public MinCost As Double                              '����������� ��������

    Public MyRecordID As String                           'GUID ������ � ��������� ��������
End Module
