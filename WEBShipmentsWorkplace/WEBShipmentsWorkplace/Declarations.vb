Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '�������� ��� ������������ Scala
    Public UserID As Integer                              '��� ������������ Scala
    Public UserName As String                             '����� ������������ Scala
    Public FullName As String                             '��� ������������ Scala
    Public SalesmanCode As String                         '��� ��������
    Public SalesmanName As String                         '��� ��������

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyOrderNum As String                           '����� ������
    Public CustomerID As String                           '��� ����������
    Public CreditAmount As Double                         '������ ������� � ������
    Public CreditInDays As Integer                        '������ ������� � ����
    Public MinMargin As Double                            '����������� ����� � ������
    Public MyPermission As Boolean                        '����� ��� ��� ��������� �����
    Public MyMarginReason As String                       '������� �������� � ������ ���� �������������
    Public OrderID As String                              '����� ������
    Public OrderSum As Double                             '����� ������
    Public CurrCode As String                             '������
    Public Avance1Type As Double                          '����� ������� 1 ���� �� ������
    Public Avance2Type As Double                          '����� ������� 2 ���� �� ������
    Public MyPayment As Double                            '����� �������� �� ������
    Public InvoiceDebt As Double                          '�������� ���� �� ������ �������� RUR
    Public OrderDebt As Double                            '�������� ���� �� ������� ����������� � �������� (��� ������ ������) RUR
    Public OverduePaymentQTY As Integer                   '���-�� �������� � ������������ �������
    Public Overdue As Double                              '����� �������� � ������������ �������
    Public CmdToShip As Boolean                           '���� ������� �� �������� ����� ������ ��� ���
    Public MyReason As String                             '������� ��������
    Public MyShipmentsID As String                        '����� ��������
    Public MyOperationFlag As Integer                     '���� - ��������� �������� ��� ���
    Public MyCustomerCode As String                       '��� ����������
    Public MyWH As String                                 '��� ������

    Public MyOrderDetails As OrderDetails                 '���������� ���� ������� �� ������
    Public MyLowMarginReason As LowMarginReason           '���������� ���� ����� ������� �������� � ������ ������
    Public MyCreditDialog As CreditDialog                 '���������� ���� ������ ���������� �� ���������� �������
    Public MyNonCreditDialog As NonCreditDialog           '���������� ���� ������ ���������� �� ������������ �������
    Public MyCreditInfo As CreditInfo                     '���������� ���� ������ ��������� ���������� �� ���������� �������
    Public MyNonCreditInfo As NonCreditInfo               '���������� ���� ������ ��������� ���������� �� ������������ �������
    Public MySendInfo As SendInfo                         '���������� ���� �������� ���������� ������� �� EMail
    Public MyShipment As Shipment                         '���������� ���� �������
    Public MyContactInfo As ContactInfo                   '���������� ���� ������ ��������� ��  CRM
    Public MyDelAddresses As DelAddresses                 '���������� ���� ������ ������� �� Scala

    Public MyProjectIsApproved As Integer                 '---��������� �� ����� - ��������� ��� �����������
    Public MyProjectID As String                          '---ID �������

    Public IsWEBOrder As Integer                          '---�������� �� ����� ������� � WEB ����� (1) ��� ��� (0)
End Module
