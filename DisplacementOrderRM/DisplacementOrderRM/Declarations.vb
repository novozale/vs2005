Module Declarations
    Public CompanyID As String                                '������ - ����� �������� � Scala
    Public Year As String                                     'c����� - ��� � Scala
    Public UserCode As String                                 '�������� ��� ������������ Scala
    Public UserID As Integer                                  '��� ������������ Scala
    Public FullName As String                                 '��� ������������ Scala

    Public MyConnStr As String                                '������ ���������� � ��
    Public MyNETConnStr As String                             '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                         '���������� � ��
    Public MyRec As ADODB.Recordset                           '��������� �� � ������ ���������

    Public WHFrom As String                                   '����� ��������
    Public WHFromCode As String                               '����� �������� - ���
    Public WHTo As String                                     '����� �������
    Public WHToCode As String                                 '����� ������� - ���
    Public MyOrderID As Integer                               '����� ����������� ������

    Public MyPermission As Boolean                            '����������� ��� ��� � ����������������� ������

    Public MyConsolidatedOrders As ConsolidatedOrders         '���������� ���� ������������ ���������� ������� (��������)
    Public MyEditConsolidatedOrder As EditConsolidatedOrder   '���������� ���� �������� / �������������� ����������� ������
End Module
