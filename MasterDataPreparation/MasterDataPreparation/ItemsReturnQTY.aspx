<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ItemsReturnQTY.aspx.vb" Inherits="MasterDataPreparation.ItemsReturnQTY" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Запасы, которые могут быть возвращены</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+%d0%b2%d0%be%d0%b7%d0%b2%d1%80%d0%b0%d1%82%d0%be%d0%b2&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp;Список запасов, которые могут быть возвращены поставщикам</h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Список запасов с группировкой по поставщикам</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="ID">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("ID")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код поставщика">
                                            <ItemTemplate>
                                                <%#Eval("SuppID")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Поставщик">
                                            <ItemTemplate>
                                                <%#Eval("SuppName")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Код запаса">
                                            <ItemTemplate>
                                                <%#Eval("SC01001")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Запас">
                                            <ItemTemplate>
                                                <%#Eval("Name")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Кол-во, которое можно вернуть">
                                            <ItemTemplate>
                                                <%#Eval("QTY")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="QTY" Text='<%# Bind("QTY") %>' Width="100%"/>
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
                                        SelectCommand="SELECT     tbl_ItemsReturnQTY.ID, ISNULL(SC010300.SC01058, N'000000') AS SuppID, ISNULL(PL010300.PL01002, N'Неизвестен') AS SuppName, &#13;&#10;                      tbl_ItemsReturnQTY.SC01001, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, N''))) + ' ' + LTRIM(RTRIM(ISNULL(SC010300.SC01003, N''))))) AS Name, &#13;&#10;                      CONVERT(float, tbl_ItemsReturnQTY.QTY) AS QTY&#13;&#10;FROM         PL010300 RIGHT OUTER JOIN&#13;&#10;                      SC010300 ON PL010300.PL01001 = SC010300.SC01058 RIGHT OUTER JOIN&#13;&#10;                      tbl_ItemsReturnQTY ON SC010300.SC01001 = tbl_ItemsReturnQTY.SC01001&#13;&#10;ORDER BY SuppID, tbl_ItemsReturnQTY.SC01001" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_ItemsReturnQTY&#13;&#10;WHERE     (ID = @ID)" UpdateCommand="UPDATE    tbl_ItemsReturnQTY&#13;&#10;SET  QTY = @QTY&#13;&#10;WHERE     (ID = @ID)" InsertCommand="INSERT INTO tbl_ItemsReturnQTY&#13;&#10;                      (ID, SC01001, QTY)&#13;&#10;VALUES     (NEWID(), @SC01001, @QTY)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="ID" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="QTY" />
                                            <asp:Parameter Name="ID" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="SC01001" />
                                            <asp:Parameter Name="QTY" />
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
                                                <asp:TextBox runat="server" ID="InsertSC01001" Width="100%"/>
                                            </td>
                                           <td style="width: 20%;">
                                                Кол - во, которое можно вернуть
                                            </td>
                                            <td style="width: 20%;">
                                                <asp:TextBox runat="server" ID="InsertQTY" Width="100%"/>
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

