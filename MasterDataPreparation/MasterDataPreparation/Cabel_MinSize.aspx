<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Cabel_MinSize.aspx.vb" Inherits="MasterDataPreparation.Cabel_MinSize" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Минимальный размер куска кабеля</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp;Список кодов кабельной продукции с минимальными размерами куска кабеля</h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Список кабельной продукции</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="ID">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("ID")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код Запаса">
                                            <ItemTemplate>
                                                <%#Eval("ItemCode")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Запас">
                                            <ItemTemplate>
                                                <%#Eval("ItemName")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код поставщика">
                                            <ItemTemplate>
                                                <%#Eval("SupplierCode")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Поставщик">
                                            <ItemTemplate>
                                                <%#Eval("SupplierName")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Минимальный размер куска">
                                            <ItemTemplate>
                                                <%#Eval("MinSize")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="QTY" Text='<%# Bind("MinSize") %>' Width="100%"/>
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
                                        SelectCommand="SELECT     tbl_Cabel_MinSize.ID, tbl_Cabel_MinSize.ItemCode, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, N''))) + ' ' + LTRIM(RTRIM(ISNULL(SC010300.SC01003, &#13;&#10;                      N''))))) AS ItemName, ISNULL(SC010300.SC01058, N'') AS SupplierCode, ISNULL(PL010300.PL01002, N'') AS SupplierName, tbl_Cabel_MinSize.MinSize&#13;&#10;FROM         PL010300 INNER JOIN&#13;&#10;                      SC010300 ON PL010300.PL01001 = SC010300.SC01058 RIGHT OUTER JOIN&#13;&#10;                      tbl_Cabel_MinSize ON SC010300.SC01001 = tbl_Cabel_MinSize.ItemCode&#13;&#10;ORDER BY tbl_Cabel_MinSize.ItemCode" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_Cabel_MinSize&#13;&#10;WHERE     (ID = @ID)" UpdateCommand="UPDATE    tbl_Cabel_MinSize&#13;&#10;SET  MinSize = @MinSize&#13;&#10;WHERE     (ID = @ID)" InsertCommand="INSERT INTO tbl_Cabel_MinSize&#13;&#10;                      (ID, ItemCode, MinSize)&#13;&#10;VALUES     (NEWID(), @ItemCode, @MinSize)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="ID" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="ID" />
                                            <asp:Parameter Name="MinSize" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="ItemCode" />
                                            <asp:Parameter Name="MinSize" />
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
                                                код запаса
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:TextBox runat="server" ID="InsertItemCode" Width="100%"/>
                                            </td>
                                           <td style="width: 20%;">
                                                Минимальный размер куска
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:TextBox runat="server" ID="InsertMinSize" Width="100%"/>
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:Button runat="server" ID="Button1" Text="Insert" CommandName="InsertNew" />
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

