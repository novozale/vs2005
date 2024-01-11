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

    Public MyEditHeader As EditHeader                     '���������� ���� �������������� ���������
    Public MyCustomerSelect As CustomerSelect             '���������� ���� ������ ��������
    Public MyCustomerSelectList As CustomerSelectList     '���������� ���� ������� ��� ������ ��������
    Public MyOrderLines As OrderLines                     '���������� ���� ����� ����� �����������
    Public MySupplierSelect As SupplierSelect             '���������� ���� ������ �����������
    Public MySupplierSelectList As SupplierSelectList     '���������� ���� ������� ��� ������ �����������
    Public MyItemSelectList As ItemSelectList             '���������� ���� ������� ��� ������ �������
    Public MyShowWHDetails As ShowWHDetails               '���������� ���� ���������� �� �������
    Public MyShowBatchInfo As ShowBatchInfo               '���������� ���� ���������� �� �������
    Public MyALTItems As ALTItems                         '���������� ���� �� ������� �������������� �������
    Public MyAddToOrder As AddToOrder                     '���������� ���� ���������� ������ � �����
    Public MyEditInOrder As EditInOrder                   '���������� ���� �������������� ������ � ������
    Public MyItemSelect As ItemSelect                     '���������� ���� ������ ������
    Public MySelectItemBySuppCode As SelectItemBySuppCode '���������� ���� ������ ������ �� ���� ������ ����������
    Public MySendReturnProposal As SendReturnProposal     '���������� ���� �������� / �������� ������������� �����������
    Public MyShipmentsCost As ShipmentsCost               '���������� ���� ����� ��������� ��������
    Public MyEDOInfo As EDOInfo                           '���������� ���� ����� �������������� ���������� �� ���
    Public MyEstimatedIncome As EstimatedIncome           '���������� ���� � ����������� �� ��������� �������
    Public MySearchSupplier As SearchSupplier             '���������� ���� ������ ������ �� ����� ����������
    Public MyEditRequest As EditRequest                   '���������� ���� �������� / �������������� ������ �� ����� ����������
    Public MyAttachmentsList As AttachmentsList           '���������� ���� �������������� ������
    Public MyAddItem As AddItem                           '���������� ���� ���������� ������ � ������
    Public MyCPList As CPList                             '���������� ���� ������ ��
    Public MyContactInfo As ContactInfo                   '���������� ���� ������ ��������� ��  CRM
    Public MySalesCommentsToProposal As SalesCommentsToProposal  '���������� ���� ����� ������������ �� ������ �����������
    Public MyCommentAndCancelReason As CommentAndCancelReason '���������� ���� ����� ������������ � ������ ������ � ������ ������
    Public MyRestoreSearch As RestoreSearch               '���������� ���� �������������� ������ ����� �����
    Public MyCorrectRequestDate As CorrectRequestDate     '���������� ���� ������������� ����������� ���� �������������� ��
    Public MyPrintProposal As PrintProposal               '���������� ���� ������ ��

    Public MyOrderNum As String                           '����� ������
    Public CurrencyCode As Integer                        '��� ������ (��� ���� ������)
    Public CurrencyName As String                         '�������� ������
    Public CurrencyValue As Double                        '�������� ������ � ������ �� ������� ����
    Public CurrencyValueOrder As Double                   '�������� ������ � ������ �� ���� �������� ������
    Public MyRequestNum As Integer                        '����� ������� �� ����� ����������
    Public MyItemSrchID As Integer                        'ID ������ � ������� �� �����
    Public MyItemPropID As Integer                        'ID ������ � ����������� ������
    Public MyCPID As String                               'ID ������������� ����������� ��� �������� � ���� ����������� �� �����������
    Public MyRowIndex As Integer                          'ID ������ � ������� (��� ������������ ����)
    Public MyRez1 As Integer                              '��������� ��������

    Public CustomerNumber                                 '��� �������
    Public WHNum                                          '����� ������
    Public MySuccess As Boolean                           '���������� ���������� ��������
    Public MyItemID As String                             '��� ������
    Public MyItemSuppID As String                         '��� ������ ����������
    Public MyItemName As String                           '��� ������
    Public MyQty As Double                                '���������� �����������
    Public MyUOM As Integer                               '��� ������� ���������
    Public MySum As Double                                '����� �����������
    Public MySS As Double                                 '�������������
    Public MyDiscount As String                           '������
    Public MyItemCode As String                           '��� ������ � Scala ��� ������� ������������ �� Excel
    Public MySuppID As String                             '��� ����������
    Public MySuppName                                     '�������� ����������

    Public ImportFileName As String                       '��� ����� ��� ������� ������
    Public ExcelVersion As String                         '������ ����� ��� ������� 

    Public DeliveryDate As Date                           '�������������� ���� ��������
    Public DeliveryDateFlag As Integer                    '���� - ��������� ���� �� ���� ������� (1) ��� ���(0)
    Public WeekQTY As Double                              '���� �������� � ������� (������ ���� >= 0; 0 - � ������� �� ������)
    Public DelWeekQTY As Double                           '���� �������� �� ������� � ������� (������ ���� >= 0; 0 - ���� ��� �������� (���������))

    Public MinMarginLevelManager As Double                '����������� �������� �����, ������������ ����������
    Public MinMarginLevelDirector As Double               '����������� �������� �����, ������������ ����������

    Public MyCCPermission As Boolean                      '������  ��� ��� � ������ CRMManagers
    Public MyCPPermission As Boolean                      '������  ��� ��� � ������ ProposalManager
    Public MyPermission As Boolean                        '������  ��� ��� � ������ CRMDirector

    Public IsSelfDelivery As Integer                      '��������� (1) ��� ��� (0)
    Public MyOperationResult As Integer                   '��������� ���������� �������� (1) - ��������� (0) - ���

    Public SortColumnNum As Integer                       '����� ������� � ����������� ��� ����������
    Public SortColOrder As System.ComponentModel.ListSortDirection '����������� ����������
End Module
