<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MarginCoeffByItem.aspx.vb" Inherits="MasterDataPreparation.MarginCoeffByItem" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server" >
    Dim Conn As New OleDbConnection _
             ("Provider=SQLOLEDB.1;Server=sqlcls;Database=ScaDataDB;User ID = sa;Password=sqladmin; ")

    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Настройка коэффициента маржи для прайс - листа по запасам</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+%d1%86%d0%b5%d0%bd%d0%be%d0%be%d0%b1%d1%80%d0%b0%d0%b7%d0%be%d0%b2%d0%b0%d0%bd%d0%b8%d1%8f&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Список запасов Rexel с индивидуальным коэфф. маржи 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; width: 1202px;">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Список запасов Rexel с индивидуальным коэфф. маржи</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="SC01001">
                                    <Columns>
                                        <asp:TemplateField HeaderText="Код запаса">
                                            <ItemTemplate>
                                                <%#Eval("SC01001")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Название запаса">
                                            <ItemTemplate>
                                                <%#Eval("SC01002")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Коэфф. маржи">
                                            <ItemTemplate>
                                                <%#Eval("MarginCoeff")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="MarginCoeff" Text='<%# Bind("MarginCoeff") %>' Width="100%"/>
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
                                        SelectCommand="SELECT     SC01001, SC01002, MarginCoeff&#13;&#10;FROM         tbl_PriceCoeff_ByItem&#13;&#10;ORDER BY SC01001" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_PriceCoeff_ByItem&#13;&#10;WHERE     (SC01001 = @SC01001)" UpdateCommand="UPDATE    tbl_PriceCoeff_ByItem&#13;&#10;SET              MarginCoeff = @MarginCoeff&#13;&#10;WHERE     (SC01001 = @SC01001)" InsertCommand="INSERT INTO tbl_PriceCoeff_ByItem&#13;&#10;                      (SC01001, SC01002, MarginCoeff)&#13;&#10;SELECT     SC01001, LTRIM(RTRIM(SC01002 + ' ' + SC01003)) AS SC01002, @MarginCoeff  AS MarginCoeff&#13;&#10;FROM         SC010300&#13;&#10;WHERE     (SC01001 = @SC01001)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="SC01001" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="SC01001" PropertyName="SelectedValue" />
                                            <asp:ControlParameter ControlID="GridView1" Name="MarginCoeff" PropertyName="SelectedValue" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:ControlParameter ControlID="InsertSC01001" Name="SC01001" PropertyName="Text" />
                                            <asp:ControlParameter ControlID="InsertMarginCoeff" Name="MarginCoeff" PropertyName="Text" />
                                        </InsertParameters>
                                    </asp:SqlDataSource>
                                </div>
                            </td>
                        </tr>
                        <tr style="width: 100%; font-family: Arial; font-size: x-small;" >
                            <td style="vertical-align:top; width: 1202px;">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px; background-color:#cccccc; text-align: right;">
                                   <table style="width: 100%;">
                                       <tr style="width: 100%;">
                                            <td style="width: 22%;">
                                                Код запаса
                                            </td>
                                            <td style="width: 23%;">
                                                <asp:TextBox runat="server" ID="InsertSC01001" Width="100%"/>
                                            </td>
                                            <td style="width: 22%;">
                                                Коэффициент маржи
                                            </td>
                                            <td style="width: 23%;">
                                                <asp:TextBox runat="server" ID="InsertMarginCoeff" Width="100%"/>
                                            </td>
                                            <td style="width: 10%;">
                                                <asp:Button runat="server" ID="Button1" Text="Insert" CommandName="InsertNewMC" />
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

