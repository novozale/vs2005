Module declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '��� ������������ Scala
    Public SalesmanCode As String                         '��� �������� Scala
    Public SalesmanName As String                         '��� �������� Scala
    Public ScalaDate As Date                              '��������� ���� Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyDoc As Xml.XmlDocument                       'XML - ��������
    Public MyHeaderNode As Xml.XmlNode                    '���� - ��������� ��
    Public MyFirstItemNode As Xml.XmlNode                 '���� - ��������� ������� �� ��
    Public MyItemNodeList As Xml.XmlNodeList              '������ ����� - ������� ��

    Public MyErrorForm As ErrorForm                       '���������� ���� ������ ��������� �� ������

    Public appXLSRC As Object                             'Excel - ��������
End Module
