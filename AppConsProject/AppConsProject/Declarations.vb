Public Module Declarations

    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala

    Public Conn As ADODB.Connection
    Public NETConnStr As String
    Public ConnStr As String

    Public FrmAddProduct As AddProduct           ' ����� ��� ���������� ������ �� ����. �����
    Public FrmEditProduct As EditProduct     ' ����� ��� ���������� ������ �� ����. �����    
    Public Rec As ADODB.Recordset

    Public SelectList As ItemSelectList
    Public SelectList2 As ItemSelectList2

    Public IsSuccess As Boolean

    Public MinQty As Double   '����������� ������� �������
    Public MaxQty As Double   '������������ ������� �������    


End Module
