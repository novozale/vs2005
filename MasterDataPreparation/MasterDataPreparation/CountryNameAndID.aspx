<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CountryNameAndID.aspx.vb" Inherits="MasterDataPreparation.CountryNameAndID" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Название и код страны - производителя</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+-+%d1%81%d1%82%d1%80%d0%b0%d0%bd%d1%8b+%d0%bf%d1%80%d0%be%d0%b8%d0%b7%d0%b2%d0%be%d0%b4%d0%b8%d1%82%d0%b5%d0%bb%d1%8f&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Список стран - производителей&nbsp;</h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Название и код страны - производителя</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="ID">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("ID")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Имя страны производителя">
                                            <ItemTemplate>
                                                <%#Eval("Name")%>  
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Name" Text='<%# Bind("Name") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код страны производителя">
                                            <ItemTemplate>
                                                <%#Eval("Code")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Code" Text='<%# Bind("Code") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                       <asp:TemplateField HeaderText="Commands">
                                            <ItemTemplate>
                                                <asp:LinkButton runat="server" ID="LBEdit" Text="Редактировать" CommandName="Edit" /> 
                                                <asp:LinkButton runat="server" ID="LBDelete" Text="Удалить" CommandName="Delete" />               
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:LinkButton runat="server" ID="LBUpdate" Text="Обновить" CommandName="Update" /> 
                                                <asp:LinkButton runat="server" ID="LBCancel" Text="Отмена" CommandName="Cancel" />
                                            </EditItemTemplate>
                                       </asp:TemplateField>
                                   </Columns>
                                    <FooterStyle BackColor="#CCCCCC" />
                                    <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="#CCCCCC" />
                                    </asp:GridView>
                                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=sqlcls;Initial Catalog=ScaDataDB;Persist Security Info=True;User ID=sa;Password=sqladmin"
                                        SelectCommand="SELECT     ID, Name, Code&#13;&#10;FROM         tbl_CountryNameAndCode&#13;&#10;ORDER BY Code, Name" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_CountryNameAndCode&#13;&#10;WHERE     (ID = @ID)" InsertCommand="INSERT INTO tbl_CountryNameAndCode&#13;&#10;                      (ID, Name, Code)&#13;&#10;VALUES     (NEWID(), @Name, @Code)" UpdateCommand="UPDATE    tbl_CountryNameAndCode&#13;&#10;SET              Name = @Name, Code = @Code&#13;&#10;WHERE     (ID = @ID)">
                                        <DeleteParameters>
                                            <asp:Parameter Name="ID" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="Name" />
                                            <asp:Parameter Name="Code" />
                                            <asp:Parameter Name="ID" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="Name" />
                                            <asp:Parameter Name="Code" />
                                        </InsertParameters>
                                    </asp:SqlDataSource>
                                </div>
                            </td>
                        </tr>
                        <tr style="width: 100%; font-family: Arial; font-size: x-small;" >
                            <td style="vertical-align:top;">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px; background-color:#cccccc; text-align: right;">
                                   <table style="width: 100%;">
                                       <tr style="width: 100%;">
                                            <td style="width: 20%;">
                                                Название страны производителя
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:TextBox runat="server" ID="InsertName" Width="100%"/>
                                            </td>
                                            <td style="width: 20%;">
                                                Код страны производителя
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:TextBox runat="server" ID="InsertCode" Width="100%"/>
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:Button runat="server" ID="Button1" Text="Insert" CommandName="InsertCountry" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>