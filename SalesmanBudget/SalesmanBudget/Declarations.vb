Module Declarations
    Public CompanyID As String                              '������ - ����� �������� � Scala
    Public Year As String                                   'c����� - ��� � Scala
    Public UserCode As String                               '�������� ��� ������������ Scala

    Public MyConnStr As String                              '������ ���������� � ��
    Public MyNETConnStr As String                           '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                       '���������� � ��
    Public MyRec As ADODB.Recordset                         '��������� �� � ������ ���������

    Public StrTotalStart As Integer                         '� ����� ������ ���������� ���� ������� �������� ������
    Public IndustryQTY As Integer                           '���-�� ����� ��������
    Public TypeQTY As Integer                               '���-�� ����� �����
    Public MarketQTY As Integer                             '���-�� ����� ������
    Public IKAQTY As Integer                                '���-�� ����� ����� IKA
    Public ActiveAreaStart As Integer                       '������ - ������ ������ �������� ��������
    Public ActiveAreaFinish As Integer                      '������ - ����� ������ �������� ��������
    Public PassiveAreaStart As Integer                      '������ - ������ ������ ���������� ��������
    Public PassiveAreaFinish As Integer                     '������ - ����� ������ ���������� ��������
    Public NewAreaStart As Integer                          '������ - ������ ������ ����� ��������
    Public NewAreaFinish As Integer                         '������ - ����� ������ ����� ��������
End Module
