Module Declarations
    Public CompanyID As String                            '������ - ����� �������� � Scala
    Public Year As String                                 'c����� - ��� � Scala
    Public UserCode As String                             '�������� ��� ������������ Scala

    Public MyConnStr As String                            '������ ���������� � ��
    Public MyNETConnStr As String                         '.NET ������ ���������� � ��
    Public MyConn As ADODB.Connection                     '���������� � ��
    Public MyRec As ADODB.Recordset                       '��������� �� � ������ ���������

    Public MyCity As City                                 '���������� ���� ����� ���������� �� ������
    Public MyCityID As Integer                            '��� ������
    Public MyManufacturer As Manufacturer                 '���������� ���� ����� ���������� �� �������������
    Public MyManufacturerID As Integer                    '��� �������������
    Public MySalesman As Salesman                         '���������� ���� ����� ���������� �� ��������
    Public MySalesmanID As String                         '��� ��������
    Public MyProductGroup As ProductGroup                 '���������� ���� ����� ���������� �� ������ ���������
    Public MyProductGroupID As String                     '��� ������ ���������
    Public MyProductSubGroup As ProductSubGroup           '���������� ���� ����� ���������� �� ��������� ���������
    Public MyProductSubGroupID As String                  '��� ��������� ���������
    Public MyProduct As Product                           '���������� ���� ����� ���������� �� ��������
    Public MyProductID As String                          '��� �������� � Scala
    Public MyCustomerSelectList As CustomerSelectList     '���������� ���� ������� ��� ������ ��������
    Public MyCustomer As Customer                         '���������� ���� ����� ���������� �� �������
    Public MyCustomerID As String                         '��� ������� � Scala
    Public MyDiscountGroup As DiscountGroup               '���������� ���� ����� / �������������� ������ �� ������ �������
    Public MyDiscountSubgroup As DiscountSubgroup         '���������� ���� ����� / �������������� ������ �� ��������� �������
    Public MyDiscountItem As DiscountItem                 '���������� ���� ����� / �������������� ������ �� �����
    Public MyItemList As ItemList                         '���������� ���� ������ ������
    Public MyItemSelectList As ItemSelectList             '���������� ���� ������ ������ � ���������
    Public MyAgreedRange As AgreedRange                   '���������� ���� ����� �������������� ������������
    Public MyBasePrice As BasePrice                       '���������� ���� ������ � ������ �������� ����� �����
    Public MyIndPrice As IndPrice                         '���������� ���� ������ � ������ ��������������� ����� �����
    Public MyUploadFilesToDB As UploadFilesToDB           '���������� ���� �������� �������� � ��
    Public MyUploadFilesToCatalog As UploadFilesToCatalog '���������� ���� �������� �������� � �������
    Public MyUploadInfoToSaintGobain As UploadInfoToSaintGobain '���������� ���� �������� ���������� ��� ��� ������
    Public MyDeletePictures As DeletePictures             '���������� ���� �������� ������� ������������������� ��������
    Public MyTransferNamesDescrToDB As TransferNamesDescrToDB '���������� ���� �������� �������� � ��������, ���������� �� WEB, � ��
    Public MyMatchPictAndScalaCode As MatchPictAndScalaCode '���������� ���� ���������� �������� � ����� Scala
    Public MyDeletePictureFromDB As DeletePictureFromDB     '���������� ���� �������� ��������� �� ��
    Public MyLoadOnePictToDB As LoadOnePictToDB             '���������� ���� �������� ����� �������� � ��
    Public MyCheckUpdateNamesDescr As CheckUpdateNamesDescr '���������� ���� �������� ������������ �������� / �������� ��������� � �� �������� � ��
    Public MyDownloadInfoFromSE As DownloadInfoFromSE       '���������� ���� ��������� ������ � ������� ������� ��������
    Public MyDownloadInfoFromABB As DownloadInfoFromABB     '���������� ���� ��������� ������ � ������� ABB
    Public MyUploadDataToMagento As UploadDataToMagento     '���������� ���� �������� ������ �� ���� magento
    Public MyUploadPicturesToMagento As UploadPicturesToMagento '���������� ���� �������� �������� �� ���� magento
    Public MyUploadAvailabilityToMagento As UploadAvailabilityToMagento '���������� ���� �������� ���������� � ��������� ���������� �� ���� Magento
    Public MyErrWindow As ErrWindow                         '���������� ���� ������ ��������� �� ������
    Public MyCASH_FullUpload As CASH_FullUpload             '���������� ���� �������� ���� ����������������� ������������ �� �����
    Public MyCASH_CustomUpload As CASH_CustomUpload         '���������� ���� �������� ���������� ������ ������������ �� �����
    Public MyErrorMessage As ErrorMessage                   '���������� ���� � ���������� �� �������

    Public MyFilterColumn As Integer                      '�������, �� ������� ������������ ������
    Public MyAccessToken As String = "g7rvo6kkef82uwv5isvwokk3n2ohicyh" '����� ��� ������ ������� � Magento
    Public MyGlobalStr As String                            '���������� ������
End Module
