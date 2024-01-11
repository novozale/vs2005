<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RelocationOrderR4.aspx.vb" Inherits="ForecastOrderWEB.RelocationOrderR4" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.OleDB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript">
    function MoveToExcel()
    {
        var grid = document.getElementById("<%=GridView1.ClientID%>");
        var code;
        var QTY;
        var Count;
        
        var MyObj = new ActiveXObject ("Excel.Application");
        var MyWRKBook = MyObj.Workbooks.Add();
        
        //---Вывод заголовка
        var today = new Date();
        MyWRKBook.ActiveSheet.Range("A1") = "Предлагаемое перемещение кабеля от " + today.toLocaleDateString();
        var DCLabel = document.getElementById("<%=Label5.ClientID%>");
        MyWRKBook.ActiveSheet.Range("A2") = "исходный склад (DC): " + DCLabel.innerHTML;
        var RWHLabel = document.getElementById("<%=Label3.ClientID%>");
        MyWRKBook.ActiveSheet.Range("A3") = "Пополняемый склад: " + RWHLabel.innerHTML;
        MyWRKBook.ActiveSheet.Range("A5") = "Код Товара";
        MyWRKBook.ActiveSheet.Range("B5") = "Количество";
        
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 30;
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 30;
        
        MyWRKBook.ActiveSheet.Range("A5:B5").Select();
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(5).LineStyle = -4142;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(6).LineStyle = -4142;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(7).LineStyle = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(7).Weight = 4;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(7).ColorIndex = -4105;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(8).LineStyle = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(8).Weight = 4;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(8).ColorIndex = -4105;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(9).LineStyle = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(9).Weight = 4;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(9).ColorIndex = -4105;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(10).LineStyle = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(10).Weight = 4;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(10).ColorIndex = -4105;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(11).LineStyle = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(11).Weight = 4;
        MyWRKBook.ActiveSheet.Range("A5:B5").Borders(11).ColorIndex = -4105;
        MyWRKBook.ActiveSheet.Range("A5:B5").Interior.ColorIndex = 36;
        MyWRKBook.ActiveSheet.Range("A5:B5").Interior.Pattern = 1;
        MyWRKBook.ActiveSheet.Range("A5:B5").Interior.PatternColorIndex = -4105;
        
        MyWRKBook.ActiveSheet.Range("A1:A3").Select();
        MyWRKBook.ActiveSheet.Range("A1:A3").Font.Bold = true;
        MyWRKBook.ActiveSheet.Range("A5:B5").Select();
        MyWRKBook.ActiveSheet.Range("A5:B5").Font.Bold = true;
        

        //---вывод тела таблицы
        if (grid.rows.length > 0) {
            Count = 6;
            for (i = 1; i < grid.rows.length-1; i++) {
                code = grid.rows[i].cells[0].innerText;
                QTY = grid.rows[i].cells[5].childNodes[0].value;
                if (QTY != '')
                {
                    if((code.substring(0,2) == '02') || (code.substring(0,2) == '03') || (code.substring(0,2) == '04') || (code.substring(0,2) == '05') || (code.substring(0,2) == '06'))
                    {
                        //alert(code + QTY);
                        MyWRKBook.ActiveSheet.Range("A" + Count.toString()).NumberFormat = "@";
                        MyWRKBook.ActiveSheet.Range("A" + Count.toString()) = code;
                        MyWRKBook.ActiveSheet.Range("B" + Count.toString()) = QTY;
                        Count = Count + 1;
                        grid.rows[i].cells[5].childNodes[0].value = "";
                    }
                }
            }
        }
        
        MyObj.visible = true;

    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Заказ на перемещение в Scala</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="color: navy; font-family: Arial; font-weight: bold; font-size: 14pt;">
        &nbsp;Заказ на перемещение в Scala&nbsp;
        <asp:Label ID="Label6" runat="server" Font-Bold="False" Width="168px" BorderStyle="Solid" BorderWidth="1px"></asp:Label>&nbsp;&nbsp;<br />
        <asp:Button ID="Button3" runat="server" Text="Перенести в Scala" Width="196px" /><br />
        <asp:Button ID="Button2" runat="server" Text="Перенести кабель в Excel" Width="196px" OnClientClick="MoveToExcel()"/><br />
        <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="Small" Text="Пополняемый склад" Width="128px"></asp:Label>
        <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Size="Small" Width="129px"></asp:Label>
        <asp:Label ID="Label1" runat="server" Width="55%" Font-Bold="True" ForeColor="Red" Height="94px" Font-Names="Arial" Font-Size="0.8em"></asp:Label><br />
        <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Size="Small" Text="DC" Width="128px"></asp:Label>
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
                            <asp:TemplateField HeaderText="Свободно на DC">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="FreeDC" Text='<%# Bind("FreeDC") %>' Width="100%"/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Рекомендуемый заказ">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="RecQTY" Text='<%# Bind("RecQTY") %>' Width="100%"/>
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
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ScaDataDBConnectionString %>"
                        ProviderName="<%$ ConnectionStrings:ScaDataDBConnectionString.ProviderName %>"
                        SelectCommand="spp_ForecastOrderR4_RelocationOrderPreparation" SelectCommandType="StoredProcedure">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="Label3" DefaultValue="" Name="MyWarNo" PropertyName="Text"
                                Type="String" />
                            <asp:ControlParameter ControlID="Label5" DefaultValue="" Name="MySrcWH" PropertyName="Text"
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