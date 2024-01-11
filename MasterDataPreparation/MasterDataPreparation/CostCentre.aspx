<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CostCentre.aspx.vb" Inherits="MasterDataPreparation.CostCentre" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Параметры кост-центров</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Установка дополнительных параметров кост-центров 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Кост-центры</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="CCNo">
                                    <Columns>
                                        <asp:TemplateField HeaderText="N кост-центра">
                                            <ItemTemplate>
                                                <%#Eval("CCNo")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Город">
                                            <ItemTemplate>
                                                <%#Eval("City")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="City" Text='<%# Bind("City") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Название кост-центра">
                                            <ItemTemplate>
                                                <%#Eval("CCName")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="CCName" Text='<%# Bind("CCName") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="B2B">
                                            <ItemTemplate>
                                                <%#Eval("B2B")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="B2B" Text='<%# Bind("B2B") %>' Width="100%" >
                                                    <asp:ListItem>Да</asp:ListItem>
                                                    <asp:ListItem>Нет</asp:ListItem>
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Retail">
                                            <ItemTemplate>
                                                <%#Eval("Retail")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="Retail" Text='<%# Bind("Retail") %>' Width="100%" >
                                                    <asp:ListItem>Да</asp:ListItem>
                                                    <asp:ListItem>Нет</asp:ListItem>
                                                </asp:DropDownList>
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
                                        SelectCommand="SELECT CCNo, City, CCName, B2B, Retail FROM tbl_CostCentre ORDER BY CCNo" ProviderName="System.Data.SqlClient" DeleteCommand="delete from tbl_CostCentre Where CCNo = @CCNo" UpdateCommand="UPDATE    tbl_CostCentre&#13;&#10;SET City = @City, CCName = @CCName, B2B = @B2B, Retail = @Retail&#13;&#10;WHERE     (CCNo = @CCNo)" InsertCommand="INSERT INTO tbl_CostCentre&#13;&#10;                      (CCNo, City, CCName, B2B, Retail)&#13;&#10;VALUES     (@CCNo, @City, @CCName, @B2B, @Retail)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="CCNo" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="City" />
                                            <asp:Parameter Name="CCName" />
                                            <asp:Parameter Name="B2B" />
                                            <asp:Parameter Name="Retail" />
                                            <asp:Parameter Name="CCNo" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="CCNo" />
                                            <asp:Parameter Name="City" />
                                            <asp:Parameter Name="CCName" />
                                            <asp:Parameter Name="B2B" />
                                            <asp:Parameter Name="Retail" />
                                        </InsertParameters>
                                    </asp:SqlDataSource>
                                </div>
                            </td>
                        </tr>
                        <tr style="width: 100%; font-family: Arial; font-size: x-small;" >
                            <td style="vertical-align:top; " >
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px; background-color:#cccccc; text-align: right;">
                                   <table style="width: 100%;">
                                       <tr style="width: 100%;">
                                            <td style="width: 12%;">
                                                Номер кост-центра
                                            </td>
                                            <td style="width: 12%;">
                                                <asp:TextBox runat="server" ID="InsertCCNo" Width="100%"/>
                                            </td>
                                            <td style="width: 12%;">
                                                Город
                                            </td>
                                            <td style="width: 12%;">
                                                <asp:TextBox runat="server" ID="InsertCity" Width="100%"/>
                                            </td>
                                            <td style="width: 12%;">
                                                Название кост-центра
                                            </td>
                                            <td style="width: 12%;">
                                                <asp:TextBox runat="server" ID="InsertCCName" Width="100%"/>
                                            </td>
                                            <td style="width: 5%;">
                                                B2B
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:DropDownList runat="server" ID="InsertB2B" Width="40px"  >
                                                    <asp:ListItem>Да</asp:ListItem>
                                                    <asp:ListItem>Нет</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 5%;">
                                                Retail
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:DropDownList runat="server" ID="InsertRetail" Width="40px"  >
                                                    <asp:ListItem>Да</asp:ListItem>
                                                    <asp:ListItem>Нет</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 8%;">
                                                <asp:Button runat="server" ID="Button1" Text="Insert" CommandName="InsertNewPar" />
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

