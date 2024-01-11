<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MarginCoeff.aspx.vb" Inherits="MasterDataPreparation.MarginCoeff" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server" >
    Dim Conn As New OleDbConnection _
             ("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")

    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Настройка коэффициента маржи для прайс - листа</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="color: navy; font-family: Arial; font-weight: bold; font-size: 14pt;">
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+%d1%86%d0%b5%d0%bd%d0%be%d0%be%d0%b1%d1%80%d0%b0%d0%b7%d0%be%d0%b2%d0%b0%d0%bd%d0%b8%d1%8f&rs:Command=Render" Font-Size="Small">Возврат в главное окно</asp:HyperLink><br />
        <br />Настройка коэффициента маржи для прайс - листа<br />
        &nbsp;<table style="width: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset;">
            <tr>
                <td colspan="3" rowspan="3" style="height: 200px; width: 100%;">
                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="SY24002">
                        <Columns>
                            <asp:TemplateField HeaderText="N группы запасов">
                                <ItemTemplate>
                                    <%#Eval("SY24002")%>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Название группы запасов">
                                <ItemTemplate>
                                    <%#Eval("SY24003") %>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Коэфф. маржи">
                                <EditItemTemplate>
                                    <asp:TextBox ID="MarginCoeff" runat="server" Text='<%# Bind("MarginCoeff") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <%#Eval("MarginCoeff")%>
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
                        SelectCommand="SELECT SY24002, SY24003, MarginCoeff  FROM  tbl_PriceCoeff_ByGroup"
                        UpdateCommand="UPDATE tbl_PriceCoeff_ByGroup SET  MarginCoeff = @MarginCoeff WHERE  (SY24002 = @SY24002)">
                        <UpdateParameters>
                            <asp:Parameter Name="MarginCoeff" />
                            <asp:Parameter Name="SY24002" />
                        </UpdateParameters>
                    </asp:SqlDataSource>
                </td>
            </tr>
        </table>
    </div>
        <br />
        <asp:Button ID="Button1" runat="server" Text="Пересчитать" Width="196px" /><br />
        <br />
        <asp:Label ID="Label1" runat="server" Width="758px"></asp:Label>
    </form>
</body>
</html>