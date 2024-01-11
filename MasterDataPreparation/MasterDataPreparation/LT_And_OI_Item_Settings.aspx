<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LT_And_OI_Item_Settings.aspx.vb" Inherits="MasterDataPreparation.LT_And_OI_Item_Settings" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Редактирование времени доставки от поставщиков и промежутков между заказами для запасов</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+ROP+%d0%b8+%d0%9c%d0%96%d0%97&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">Редактирование времени доставки от поставщиков и промежутков между заказами для закупок складского ассортимента и его аналогов<div style="color: Red; font-family: Arial; font-weight: bold; font-size: 14pt;">для запасов<br /></h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="50" DataKeyNames="Code">
                                    <Columns>
                                        <asp:TemplateField HeaderText="Код запаса">
                                            <ItemTemplate>
                                                <%#Eval("Code")%>  
                                            </ItemTemplate>
                                       </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Среднее время доставки из Scala">
                                            <ItemTemplate>
                                                <%#Eval("LT")%> 
                                            </ItemTemplate>
                                         </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Ручное время доставки">
                                            <ItemTemplate>
                                                <%#Eval("ManualLT")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="OI" Text='<%# Bind("ManualLT") %>' Width="100%"/>
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
                                        SelectCommand="SELECT Code, LT, ManualLT  FROM tbl_ForecastOrderR4_Product ORDER BY Code" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_ForecastOrderR4_Product WHERE  (Code = @Code) " UpdateCommand="UPDATE tbl_ForecastOrderR4_Product SET  ManualLT = @ManualLT WHERE  (Code = @Code)" InsertCommand="INSERT INTO tbl_ForecastOrderR4_Product (Code,  LT, ManualLT) VALUES (@Code, @LT, @ManualLT)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="Code" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="Code" />
                                            <asp:Parameter Name="ManualLT" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="Code" />
                                            <asp:Parameter Name="LT" />
                                            <asp:Parameter Name="ManualLT" />
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
                                            <td style="width: 22%;">
                                                Код запаса
                                            </td>
                                            <td style="width: 22%;">
                                                <asp:TextBox runat="server" ID="InsertCode" Width="100%"/>
                                            </td>
                                            <td style="width: 22%;">
                                                Ручное время доставки
                                            </td>
                                            <td style="width: 22%;">
                                                <asp:TextBox runat="server" ID="InsertLT" Width="100%"/>
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
        <br />
        <asp:Button ID="Button2" runat="server" Text="Пересчитать" Width="196px" /><br />
        <br />
        <asp:Label ID="Label1" runat="server" Width="758px"></asp:Label>
    </form>
</body>
</html>

