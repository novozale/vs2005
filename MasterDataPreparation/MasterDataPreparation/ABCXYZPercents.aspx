<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ABCXYZPercents.aspx.vb" Inherits="MasterDataPreparation.ABCXYZPercents" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server" >
    Dim Conn As New OleDbConnection _
             ("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")

    
</script>


<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Редактирование процентов для расчёта ABC и XYZ</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+ROP+%d0%b8+%d0%9c%d0%96%d0%97&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Редактирование процентов для расчёта ABC и XYZ 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top;">
                                <div style="width: 98%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Проценты для расчёта ABC и XYZ</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="ID">
                                        <Columns>
                                            <asp:TemplateField HeaderText="Категория">
                                                <ItemTemplate>
                                                    <%#Eval("Param") %>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="% от" >
                                                <EditItemTemplate>
                                                    <asp:TextBox ID="PercentLow" runat="server" Text='<%# Bind("PercentLow") %>'></asp:TextBox>
                                                </EditItemTemplate>
                                                <ItemTemplate>
                                                    <%#Eval("PercentLow")%>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="% до">
                                                <EditItemTemplate>
                                                    <asp:TextBox ID="PercentHigh" runat="server" Text='<%# Bind("PercentHigh") %>'></asp:TextBox>
                                                </EditItemTemplate>
                                                <ItemTemplate>
                                                    <%#Eval("PercentHigh")%>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Commands">
                                                 <ItemTemplate>
                                                     <asp:LinkButton runat="server" ID="LBEdit" Text="Редактировать" CommandName="Edit" /> 
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
                                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ScaDataDBConnectionString %>"
                                        ProviderName="<%$ ConnectionStrings:ScaDataDBConnectionString.ProviderName %>"
                                        SelectCommand="select ID, Param, PercentLow, PercentHigh from tbl_ABCXYZ_Percents "
                                        UpdateCommand="UPDATE tbl_ABCXYZ_Percents SET PercentLow = @PercentLow, PercentHigh = @PercentHigh WHERE (ID = @ID)">
                                        <UpdateParameters>
                                            <asp:Parameter Name="PercentLow" DbType="Double" />
                                            <asp:Parameter Name="PercentHight" DbType="Double" />
                                            <asp:Parameter Name="ID" />
                                        </UpdateParameters>
                                    </asp:SqlDataSource>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Button ID="Button1" runat="server" Text="Пересчитать" Width="196px" ></asp:Button><br>
        <asp:Label ID="Label1" runat="server" Width="758px"></asp:Label>
   </form>
</body>
</html>
