Module Declarations

    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyItemSelectList As ItemSelectList             '���������� ���� ������� ��� ������ �������
    Public MyItemSelectList2 As ItemSelectList2           '���������� ���� ������� ��� ������ �������
    Public MyAddCustom As AddCustom                       '���������� ���� ���������� ������ �������� ���, ROP � �����. ������
    Public MyEditCustom As EditCustom                     '���������� ���� �������������� ������ �������� ���, ROP � �����. ������

    Public MySuccess As Boolean                           '���������� ���������� ��������

    Public MyMGZ As Double                                '�������� ���
    Public MyROP As Double                                '�������� ROP
    Public MyInsuranceLVL As Double                       '�������� ���������� ������
    Public MyDueDate As Date                              '�� ����� ���� ��������� ������ ��������

    Public MyWorkLevel As Integer                         '������� ������ (0) - DC, (1) - ��������� �����

End Module
