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

    Public MyCustomerCode As String                       '��� ����������
    Public MyWH As String                                 '��� ������
    Public MyGroupOrIndividualFlag As Integer             '������ � ������ (1) ��� ������������� (0)
    Public MyShipmentsID As String                        '����� ��������
    Public MyOrderNum As String                           '����� ������
    Public MyOperationFlag As Integer                     '���� - ��������� �������� ��� ���

    Public MyEmail As Integer                             '������ ������� �������� ����� ��� ����������� (0) - �������; (1) - �� ���������� �� CRM ��������
    Public MyContact As Integer                           '������ ������� ������� (0) - �������; (1) - �� Scala (�������� �������)

    Public MyCustomerSelectList As CustomerSelectList     '���������� ���� ������ �����������
    Public MyShipmentsList As ShipmentsList               '���������� ���� ������ �������� �� ����������� �������
    Public MyShipment As Shipment                         '���������� ���� �������
    Public MyOrderDetails As OrderDetails                 '���������� ���� ������� �� ������
    Public MyCreditDialog As CreditDialog                 '���������� ���� ������ ���������� �� ���������� �������
    Public MyLowMarginReason As LowMarginReason           '���������� ���� ����� ������� �������� � ������ ������
    Public MyNonCreditDialog As NonCreditDialog           '���������� ���� ������ ���������� �� ������������ �������
    Public MyCreditInfo As CreditInfo                     '���������� ���� ������ ��������� ���������� �� ���������� �������
    Public MyNonCreditInfo As NonCreditInfo               '���������� ���� ������ ��������� ���������� �� ������������ �������
    Public MySendInfo As SendInfo                         '���������� ���� �������� ���������� ������� �� EMail
    Public MyContactInfo As ContactInfo                   '���������� ���� ������ ��������� ��  CRM
    Public MyDelAddresses As DelAddresses                 '���������� ���� ������ ������� �� Scala
    Public MyConfiguration As Configuration               '���������� ���� ��������


    Public CustomerID As String                           '��� ����������
    Public CreditAmount As Double                         '������ ������� � ������
    Public CreditInDays As Integer                        '������ ������� � ����
    Public MinMargin As Double                            '����������� ����� � ������
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
    Public MyPermission As Boolean                        '����� ��� ��� ��������� �����
    Public MyReason As String                             '������� ��������
    Public MyProjectIsApproved As Integer                 '---��������� �� ����� - ��������� ��� �����������
    Public MyProjectID As String                          '---ID �������

    Public IsWEBOrder As Integer                          '---�������� �� ����� ������� � WEB ����� (1) ��� ��� (0)
End Module
