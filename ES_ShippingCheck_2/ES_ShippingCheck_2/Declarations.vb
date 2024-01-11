Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala
    Public SalesmanCode As String                         '��� �������� Scala
    Public SalesmanName As String                         '��� �������� Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public CustomerID As String                           '��� ����������
    Public OrderID As String                              '����� ������

    Public CreditAmount As Double                         '������ ������� � ������
    Public CreditInDays As Integer                        '������ ������� � ����

    Public MinMargin As Double                            '����������� ����� � ������
    Public MyMarginReason As String                       '������� �������� � ������ ���� �������������

    Public MyLowMarginReason As LowMarginReason           '���������� ���� ����� ������� �������� � ������ ������
    Public MyCreditDialog As CreditDialog                 '���������� ���� ������ ���������� �� ���������� �������
    Public MyNonCreditDialog As NonCreditDialog           '���������� ���� ������ ���������� �� ������������ �������
    Public MyNonCreditInfo As NonCreditInfo               '���������� ���� ������ ��������� ���������� �� ������������ �������
    Public MyCreditInfo As CreditInfo                     '���������� ���� ������ ��������� ���������� �� ���������� �������


    Public OrderSum As Double                             '����� ������
    Public CurrCode As String                             '������
    Public Avance1Type As Double                          '����� ������� 1 ���� �� ������
    Public Avance2Type As Double                          '����� ������� 2 ���� �� ������
    Public MyPayment As Double                            '����� �������� �� ������
    Public InvoiceDebt As Double                          '�������� ���� �� ������ �������� RUR
    Public OrderDebt As Double                            '�������� ���� �� ������� ����������� � �������� (��� ������ ������) RUR
    Public OverduePaymentQTY As Integer                   '���-�� �������� � ������������ �������
    Public Overdue As Double                              '����� �������� � ������������ �������

    Public MyPermission As Boolean                        '����� ��� ��� ��������� �����
    Public CmdToShip As Boolean                           '���� ������� �� �������� ����� ������ ��� ���
    Public MyReason As String                             '������� ��������

    Public MyProjectIsApproved As Integer                 '---��������� �� ����� - ��������� ��� �����������
    Public MyProjectID As String                          '---ID �������

    Public IsWEBOrder As Integer                          '---�������� �� ����� ������� � WEB ����� (1) ��� ��� (0)
End Module
