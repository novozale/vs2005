<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FreightIn.aspx.vb" Inherits="MasterDataPreparation.FreightIn" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>��������� % ���������� ��� ���������� ������ ������</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="color: navy; font-family: Arial; font-weight: bold; font-size: 14pt;">
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="http://spbprd5/ReportServer/Pages/ReportViewer.aspx?%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80-%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5%2f%d0%9c%d0%b0%d1%81%d1%82%d0%b5%d1%80+%d0%b4%d0%b0%d0%bd%d0%bd%d1%8b%d0%b5+%d0%b4%d0%bb%d1%8f+%d1%86%d0%b5%d0%bd%d0%be%d0%be%d0%b1%d1%80%d0%b0%d0%b7%d0%be%d0%b2%d0%b0%d0%bd%d0%b8%d1%8f&rs:Command=Render" Font-Size="Small">������� � ������� ����</asp:HyperLink><br />
        <br />��������� % ���������� ��� ���������� ������ ������<br />
        &nbsp;<table style="width: 100%; border-left-color: navy; border-bottom-color: navy; border-top-style: outset; border-top-color: navy; border-right-style: outset; border-left-style: outset; border-right-color: navy; border-bottom-style: outset;">
            <tr>
                <td colspan="3" rowspan="3" style="height: 200px; width: 100%;">
                    <asp:GridView ID="GridView1" runat="server" style="width: 100%; font-family: Arial; font-size: x-small;" AutoGenerateColumns="False" CellPadding="3" DataSourceID="SqlDataSource1" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="Name">
                        <Columns>
                            <asp:TemplateField HeaderText="�������� ������������">
                                <ItemTemplate>
                                    <%#Eval("Name")%>
                                </ItemTemplate>
                            </asp:TemplateField>
                          <asp:TemplateField HeaderText="�������� ������������">
                                <EditItemTemplate>
                                    <asp:TextBox ID="Value" runat="server" Text='<%# Bind("Value") %>'></asp:TextBox>
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <%#Eval("Value")%>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Commands">
                                 <ItemTemplate>
                                     <asp:LinkButton runat="server" ID="LBEdit" Text="�������������" CommandName="Edit" /> 
                                 </ItemTemplate>
                                 <EditItemTemplate>
                                     <asp:LinkButton runat="server" ID="LBUpdate" Text="��������" CommandName="Update" /> 
                                     <asp:LinkButton runat="server" ID="LBCancel" Text="������" CommandName="Cancel" />
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
                        SelectCommand="SELECT Name, Value FROM tbl_MasterData_FreightCoeff "
                        UpdateCommand="UPDATE tbl_MasterData_FreightCoeff SET  Value = @Value WHERE (Name = @Name)">
                        <UpdateParameters>
                            <asp:Parameter Name="Value" />
                            <asp:Parameter Name="Name" />
                        </UpdateParameters>
                    </asp:SqlDataSource>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>