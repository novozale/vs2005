<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PurchaseOrderR4.aspx.vb" Inherits="ForecastOrderWEB.PurchaseOrder" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Заказ на закупку в Scala</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="color: navy; font-family: Arial; font-weight: bold; font-size: 14pt;">
        &nbsp;Заказ на закупку в Scala&nbsp;
        <asp:Label ID="Label6" runat="server" Font-Bold="False" Width="168px" BorderStyle="Solid" BorderWidth="1px"></asp:Label>&nbsp;
        на сумму
        <asp:Label ID="Label7" runat="server" Font-Bold="False" Width="230px"></asp:Label>
        <asp:Button ID="Button2" runat="server" Text="Пересчитать заказ" /><br />
        <asp:Button ID="Button3" runat="server" Text="Перенести в Scala" Width="196px" /><br />
        <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="Small" Text="Поставщик" Width="108px"></asp:Label>
        <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Size="Small" Width="129px"></asp:Label>
        <asp:Label ID="Label1" runat="server" Width="55%" Font-Bold="True" ForeColor="Red" Height="47px"></asp:Label><br />
        <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Size="Small" Text="Склад" Width="108px"></asp:Label>
        <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Size="Small" Width="129px"></asp:Label><br />
        <table style="width: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset;">
            <tr>
                <td colspan="3" rowspan="3" style="height: 200px; width: 100%;">
                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="1" DataKeyNames="Code">
                        <Columns>
                            <asp:TemplateField HeaderText="Код запаса">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="Code" Text='<%# Bind("Code") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Название запаса">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="Name" Text='<%# Bind("Name") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Ед. измерения">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="UOM" Text='<%# Bind("UOM") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Закуп. цена">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="Price" Text='<%# Bind("Price") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Валюта">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="Curr" Text='<%# Bind("Curr") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Рекомендуемый заказ">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="RecQTY" Text='<%# Bind("RecQTY") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Кратность в заказе">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="Mult" Text='<%# Bind("Mult") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Кол-во в заказ">
                               <ItemTemplate>
                                    <asp:TextBox runat="server" ID="QTY" Text='<%# Bind("QTY") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                      </Columns>
                        <FooterStyle BackColor="#CCCCCC" />
                        <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                        <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" />
                    </asp:GridView>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Data Source=sqlcls;Initial Catalog=ScaDataDB;Persist Security Info=True;User ID=sa;Password=sqladmin"
                        ProviderName="<%$ ConnectionStrings:ScaDataDBConnectionString.ProviderName %>"
                        SelectCommand="spp_ForecastOrderR4_PurchaseOrderPreparation" SelectCommandType="StoredProcedure">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="Label3" DefaultValue="" Name="MySupCode" PropertyName="Text"
                                Type="String" />
                            <asp:ControlParameter ControlID="Label5" DefaultValue="" Name="MyWarNo" PropertyName="Text"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </td>
            </tr>
        </table>
    </div>
        <br />
        <asp:Button ID="Button1" runat="server" Text="Перенести в Scala" Width="196px" /><br />
        <br />
        &nbsp;
    </form>
</body>
</html>