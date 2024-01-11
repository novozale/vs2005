<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RexelProductCategories.aspx.vb" Inherits="MasterDataPreparation.RexelProductCategories" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Категории товаров Rexel</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Классификация категорий товаров Rexel 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Категории товаров Rexel</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="CategoryNum">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("CategoryNum")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Название">
                                            <ItemTemplate>
                                                <%#Eval("CategoryName")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="CategoryName" Text='<%# Bind("CategoryName") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код группы товара">
                                            <ItemTemplate>
                                                <%#Eval("RPGCode")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="RPGCode" Text='<%# Bind("RPGCode") %>' Width="100%" DataSourceID="SqlDataSource2" DataTextField="RPGCode" DataValueField="RPGCode" >
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код типа товара">
                                            <ItemTemplate>
                                                <%#Eval("SRPGCode")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList runat="server" ID="SRPGCode" Text='<%# Bind("SRPGCode") %>' Width="100%" DataSourceID="SqlDataSource3" DataTextField="SRPGCode" DataValueField="SRPGCode" >
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
                                        SelectCommand="SELECT CategoryNum, CategoryName, RPGCode, SRPGCode FROM tbl_RexelProductCategory ORDER BY CategoryNum" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_RexelProductCategory WHERE  (CategoryNum = @CategoryNum) AND (CategoryNum NOT IN (SELECT ISNULL(RexelProductCategory, '') AS CategoryNum                FROM tbl_ItemCard0300 GROUP BY RexelProductCategory))" InsertCommand="INSERT INTO tbl_RexelProductCategory  (CategoryNum, CategoryName, RPGCode, SRPGCode) VALUES  (@CategoryNum, @CategoryName, @RPGCode, @SRPGCode)" UpdateCommand="UPDATE tbl_RexelProductCategory SET CategoryName = @CategoryName, RPGCode = @RPGCode, SRPGCode = @SRPGCode WHERE (CategoryNum = @CategoryNum)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="CategoryNum" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="CategoryNum" />
                                            <asp:Parameter Name="CategoryName" />
                                            <asp:Parameter Name="RPGCode" />
                                            <asp:Parameter Name="SRPGCode" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="CategoryNum" />
                                            <asp:Parameter Name="CategoryName" />
                                            <asp:Parameter Name="RPGCode" />
                                            <asp:Parameter Name="SRPGCode" />
                                        </InsertParameters>
                                    </asp:SqlDataSource>
                                    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:ScaDataDBConnectionString %>"
                                        SelectCommand="SELECT RPGCode, RPGCode + ' ' + RussianName AS NAME FROM tbl_RexelProductGroup ORDER BY RPGCode">
                                    </asp:SqlDataSource>
                                    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ScaDataDBConnectionString %>"
                                        SelectCommand="SELECT [SRPGCode], [SRPGCode] + ' ' + [RussianName] AS [NAME] FROM [tbl_RexelProductType] ORDER BY [SRPGCode]">
                                    </asp:SqlDataSource>
                                    &nbsp;&nbsp;
                                </div>
                            </td>
                        </tr>
                        <tr style="width: 100%; font-family: Arial; font-size: x-small;" >
                            <td style="vertical-align:top;">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px; background-color:#cccccc; text-align: right;">
                                   <table style="width: 100%;">
                                       <tr style="width: 100%;">
                                            <td style="width: 11%;">
                                                ID
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:TextBox runat="server" ID="InsertCategoryNum" Width="100%"/>
                                            </td>
                                            <td style="width: 11%;">
                                                Название
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:TextBox runat="server" ID="InsertCategoryName" Width="100%"/>
                                            </td>
                                            <td style="width: 11%;">
                                                Код группы товара
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:DropDownList runat="server" ID="InsertRPGCode" Width="100%" DataSourceID="SqlDataSource2" DataTextField="Name" DataValueField="RPGCode" >
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 11%;">
                                                Код типа товара
                                            </td>
                                            <td style="width: 11%;">
                                                <asp:DropDownList runat="server" ID="InsertSRPGCode" Width="100%" DataSourceID="SqlDataSource3" DataTextField="NAME" DataValueField="SRPGCode" >
                                                </asp:DropDownList>
                                            </td>
                                            <td style="width: 12%;">
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
