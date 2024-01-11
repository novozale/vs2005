<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CustomerIerarchy.aspx.vb" Inherits="MasterDataPreparation.CustomerIerarchy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Объединение клиентов</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Формирование объединения клиентов 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Объединение клиентов</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="CustomerCode,Flag">
                                    <Columns>
                                        <asp:TemplateField HeaderText="Код клиента">
                                            <ItemTemplate>
                                                <%#Eval("CustomerCode")%>  
                                            </ItemTemplate>
                                       </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Клиент">
                                            <ItemTemplate>
                                                <%#Eval("CustomerName")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="CustomerName" Text='<%# Bind("CustomerName") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Общий код">
                                            <ItemTemplate>
                                                <%#Eval("JoinCode")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="JoinCode" Text='<%# Bind("JoinCode") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Уровень группировки">
                                            <ItemTemplate>
                                                <%#Eval("Flag")%> 
                                            </ItemTemplate>
                                       </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Вид клиента">
                                            <ItemTemplate>
                                                <%#Eval("Vid")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="Vid" Text='<%# Bind("Vid") %>' Width="100%" >
                                                    <asp:ListItem>Customer</asp:ListItem>
                                                    <asp:ListItem>Supplier</asp:ListItem>
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
                                    <AlternatingRowStyle BackColor="White" />
                                    </asp:GridView>
                                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=sqlcls;Initial Catalog=ScaDataDB;Persist Security Info=True;User ID=sa;Password=sqladmin"
                                        SelectCommand="SELECT CustomerCode, CustomerName, JoinCode, Flag, Vid FROM tbl_RexelCustomerJoin ORDER BY Vid, CASE WHEN ISNULL(JoinCode,'') = N'' THEN CustomerCode ELSE JoinCode END, Flag" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_RexelCustomerJoin WHERE (CustomerCode = @CustomerCode) AND (Flag = @Flag)" UpdateCommand="UPDATE tbl_RexelCustomerJoin SET CustomerName = @CustomerName, JoinCode = @JoinCode, Vid = @Vid WHERE (CustomerCode = @CustomerCode) AND (Flag = @Flag)" InsertCommand="INSERT INTO tbl_RexelCustomerJoin (CustomerCode, CustomerName, JoinCode, Flag, Vid) VALUES  (@CustomerCode, @CustomerName, @JoinCode, @Flag, @Vid)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="CustomerCode" PropertyName="SelectedValue" />
                                            <asp:ControlParameter ControlID="GridView1" Name="Flag" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="CustomerCode" />
                                            <asp:Parameter Name="CustomerName" />
                                            <asp:Parameter Name="JoinCode" />
                                            <asp:Parameter Name="Flag" />
                                            <asp:Parameter Name="Vid" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="CustomerCode" />
                                            <asp:Parameter Name="CustomerName" />
                                            <asp:Parameter Name="JoinCode" />
                                            <asp:Parameter Name="Flag" />
                                            <asp:Parameter Name="Vid" />
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
                                            <td style="width: 11%;">
                                                Код клиента
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:TextBox runat="server" ID="InsertCustomerCode" Width="100%"/>
                                            </td>
                                            <td style="width: 11%;">
                                                Клиент
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:TextBox runat="server" ID="InsertCustomerName" Width="100%"/>
                                            </td>
                                            <td style="width: 11%;">
                                                Общий код
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:TextBox runat="server" ID="InsertJoinCode" Width="100%"/>
                                            </td>
                                            <td style="width: 5%;">
                                                Уровень группировки
                                            </td>
                                            <td style="width: 5%;">
                                                <asp:DropDownList runat="server" ID="InsertFlag" Width="40px"  >
                                                    <asp:ListItem>1</asp:ListItem>
                                                    <asp:ListItem>2</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 5%;">
                                                Вид клиента
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:DropDownList runat="server" ID="InsertVid" Width="100px"  >
                                                    <asp:ListItem>Customer</asp:ListItem>
                                                    <asp:ListItem>Supplier</asp:ListItem>
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

