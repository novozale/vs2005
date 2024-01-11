<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RexelProductGroup.aspx.vb" Inherits="MasterDataPreparation.RexelProductGroup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Группы товаров Rexel</title>

</head>
<body>
    <form id="form1" runat="server">
        <asp:HyperLink ID="HyperLink1" runat="server" Font-Bold="True" Font-Size="XX-Small" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5&rs:Command=Render">Возврат в главное окно</asp:HyperLink>
        <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Классификация групп товаров Rexel 
        </h2>
        <table style="width: 100%" >
            <tr style="width: 100%">
                <td style="width: 100%; vertical-align:top;">
                    <table style="width: 100%">
                        <tr style="width: 100%">
                            <td style="vertical-align:top; ">
                                <div style="width: 98%; height: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset; position: static; left: 0px; top: 0px;">
                                    <h2 style="color: navy; font-family: Arial">&nbsp;&nbsp; Группы товаров Rexel</h2>
                                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AllowPaging="True" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" PageSize="20" DataKeyNames="RPGCode">
                                    <Columns>
                                        <asp:TemplateField HeaderText="ID">
                                            <ItemTemplate>
                                                <%#Eval("RPGCode")%>  
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Название">
                                            <ItemTemplate>
                                                <%#Eval("RussianName")%>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="RussianName" Text='<%# Bind("RussianName") %>' Width="100%"/>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Английское название">
                                            <ItemTemplate>
                                                <%#Eval("EnglishName")%> 
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox runat="server" ID="EnglishName" Text='<%# Bind("EnglishName") %>' Width="100%"/>
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
                                        SelectCommand="SELECT RPGCode, RussianName, EnglishName FROM tbl_RexelProductGroup ORDER BY RPGCode" ProviderName="System.Data.SqlClient" DeleteCommand="DELETE FROM tbl_RexelProductGroup WHERE (RPGCode = @RPGCode) AND (RPGCode NOT IN  (SELECT  ISNULL(RPGCode, '') AS RPGCode FROM tbl_RexelProductCategory GROUP BY RPGCode))" InsertCommand="INSERT INTO  tbl_RexelProductGroup (RPGCode, RussianName, EnglishName) VALUES  (@RPGCode, @RussianName, @EnglishName)" UpdateCommand="UPDATE  tbl_RexelProductGroup SET  RussianName = @RussianName, EnglishName = @EnglishName WHERE  (RPGCode = @RPGCode)">
                                        <DeleteParameters>
                                            <asp:ControlParameter ControlID="GridView1" Name="RPGCode" PropertyName="SelectedValue" />
                                        </DeleteParameters>
                                        <UpdateParameters>
                                            <asp:Parameter Name="RussianName" />
                                            <asp:Parameter Name="EnglishName" />
                                            <asp:Parameter Name="RPGCode" />
                                        </UpdateParameters>
                                        <InsertParameters>
                                            <asp:Parameter Name="RPGCode" />
                                            <asp:Parameter Name="RussianName" />
                                            <asp:Parameter Name="EnglishName" />
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
                                            <td style="width: 15%;">
                                                ID
                                            </td>
                                            <td style="width: 15%;">
                                                <asp:TextBox runat="server" ID="InsertRPGCode" Width="100%"/>
                                            </td>
                                            <td style="width: 15%;">
                                                Название
                                            </td>
                                            <td style="width: 15%;">
                                                <asp:TextBox runat="server" ID="InsertRussianName" Width="100%"/>
                                            </td>
                                            <td style="width: 15%;">
                                                Английское название
                                            </td>
                                            <td style="width: 15%;">
                                                <asp:TextBox runat="server" ID="InsertEnglishName" Width="100%"/>
                                            </td>
                                            <td style="width: 10%;">
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


