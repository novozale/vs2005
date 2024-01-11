<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Correctures.aspx.vb" Inherits="MasterDataPreparation.Correctures" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Исключения из расчетов</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Учет возвратов, согласование отчетов с Magnitude 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top; height: 659px;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Исключения из расчетов</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="ID">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("ID")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Тип коррект.">
                                            <ItemTemplate>
                                                <%#Eval("OperTip")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="OperTip" Text='<%# Bind("OperTip") %>' Width="100%" >
                                                    <asp:ListItem>0</asp:ListItem>
                                                    <asp:ListItem>1</asp:ListItem>
                                                    <asp:ListItem>3</asp:ListItem>
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                       </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код поставщика">
                                            <ItemTemplate>
                                                <%#Eval("SupCode")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="SupCode" Text='<%# Bind("SupCode") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Поставщик">
                                            <ItemTemplate>
                                                <%#Eval("Supplier")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Supplier" Text='<%# Bind("Supplier") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код покупателя">
                                            <ItemTemplate>
                                                <%#Eval("CustCode")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="CustCode" Text='<%# Bind("CustCode") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Покупатель">
                                            <ItemTemplate>
                                                <%#Eval("Customer")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Customer" Text='<%# Bind("Customer") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Стоимость">
                                            <ItemTemplate>
                                                <%#Eval("Sales")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Sales" Text='<%# Bind("Sales") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Себестоимость">
                                            <ItemTemplate>
                                                <%#Eval("Cost")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="Cost" Text='<%# Bind("Cost") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Дата">
                                            <ItemTemplate>
                                                <%#Eval("InsDate")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="InsDate" Text='<%# Bind("InsDate") %>' Width="100%"/>
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
                                        SelectCommand="SELECT  ID, OperTip, SupCode, Supplier, CustCode, Customer, Sales, Cost, InsDate FROM  tbl_SalesFeatures ORDER BY InsDate" ProviderName="System.Data.SqlClient" DeleteCommand="delete from tbl_SalesFeatures WHERE (ID = @ID)" InsertCommand="INSERT INTO tbl_SalesFeatures  (OperTip, SupCode, Supplier, CustCode, Customer, Sales, Cost, InsDate) VALUES (@OperTip, @SupCode, @Supplier, @CustCode, @Customer, @Sales, @Cost, CONVERT(DATETIME, @InsDate, 103))" UpdateCommand="UPDATE  tbl_SalesFeatures SET ID = NEWID(), OperTip = @OperTip, SupCode = @SupCode, Supplier = @Supplier, CustCode = @CustCode, Customer = @Customer, Sales = @Sales, Cost = @Cost, InsDate = CONVERT(DATETIME, @InsDate, 103) WHERE (ID = @ID)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="ID" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="OperTip" />
                                            <asp:Parameter Name="SupCode" />
                                            <asp:Parameter Name="Supplier" />
                                            <asp:Parameter Name="CustCode" />
                                            <asp:Parameter Name="Customer" />
                                            <asp:Parameter Name="Sales" />
                                            <asp:Parameter Name="Cost" />
                                            <asp:Parameter Name="InsDate" />
                                            <asp:Parameter Name="ID" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="OperTip" />
                                            <asp:Parameter Name="SupCode" />
                                            <asp:Parameter Name="Supplier" />
                                            <asp:Parameter Name="CustCode" />
                                            <asp:Parameter Name="Customer" />
                                            <asp:Parameter Name="Sales" />
                                            <asp:Parameter Name="Cost" />
                                            <asp:Parameter Name="InsDate" />
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
                                            <td style="width: 4%;">
                                                Тип коррект.
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:DropDownList runat="server" ID="InsertOperTip" Width="40px"  >
                                                    <asp:ListItem>0</asp:ListItem>
                                                    <asp:ListItem>1</asp:ListItem>
                                                    <asp:ListItem>3</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 4%;">
                                                Код поставщика
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertSupCode" Width="70px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                Поставщик
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertSupplier" Width="70px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                Код покупателя
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertCustCode" Width="70px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                Покупатель
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertCustomer" Width="70px"/>
                                            </td>
                                           <td style="width: 4%;">
                                                Стоимость
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertSales" Width="70px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                Себестоимость
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertCost" Width="70px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                Дата
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:TextBox runat="server" ID="InsertInsDate" Width="70px"/>
                                            </td>
                                            <td style="width: 60px;">
                                                <asp:Button runat="server" ID="Button1" Text="Insert" CommandName="InsertNewPar" Width="50px"/>
                                            </td>
                                            <td style="width: 4%;">
                                                &nbsp;&nbsp;&nbsp;&nbsp;
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
