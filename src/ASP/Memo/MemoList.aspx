<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MemoList.aspx.cs" Inherits="Memo_MemoList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>제목 없음</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="#CCCCCC"
            BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px" CellPadding="4" CellSpacing="2"
            DataKeyNames="Num" DataSourceID="SqlDataSource1" ForeColor="Black">
            <FooterStyle BackColor="#CCCCCC" />
            <Columns>
                <asp:BoundField DataField="Num" HeaderText="Num" InsertVisible="False" ReadOnly="True"
                    SortExpression="Num" />
                <asp:BoundField DataField="Name" HeaderText="Name" SortExpression="Name" />
                <asp:BoundField DataField="Email" HeaderText="Email" SortExpression="Email" />
                <asp:BoundField DataField="Title" HeaderText="Title" SortExpression="Title" />
                <asp:BoundField DataField="PostDate" HeaderText="PostDate" SortExpression="PostDate" />
                <asp:BoundField DataField="PostIP" HeaderText="PostIP" SortExpression="PostIP" />
            </Columns>
            <RowStyle BackColor="White" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#CCCCCC" ForeColor="Black" HorizontalAlign="Left" />
            <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ASPConnectionString %>"
            DeleteCommand="DELETE FROM [Memos] WHERE [Num] = @Num" InsertCommand="INSERT INTO [Memos] ([Name], [Email], [Title], [PostDate], [PostIP]) VALUES (@Name, @Email, @Title, @PostDate, @PostIP)"
            SelectCommand="SELECT * FROM [Memos] ORDER BY [Num] DESC" UpdateCommand="UPDATE [Memos] SET [Name] = @Name, [Email] = @Email, [Title] = @Title, [PostDate] = @PostDate, [PostIP] = @PostIP WHERE [Num] = @Num">
            <DeleteParameters>
                <asp:Parameter Name="Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="Name" Type="String" />
                <asp:Parameter Name="Email" Type="String" />
                <asp:Parameter Name="Title" Type="String" />
                <asp:Parameter Name="PostDate" Type="DateTime" />
                <asp:Parameter Name="PostIP" Type="String" />
                <asp:Parameter Name="Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="Name" Type="String" />
                <asp:Parameter Name="Email" Type="String" />
                <asp:Parameter Name="Title" Type="String" />
                <asp:Parameter Name="PostDate" Type="DateTime" />
                <asp:Parameter Name="PostIP" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
        <br />
        &nbsp;<asp:FormView ID="FormView1" runat="server" DataKeyNames="Num" DataSourceID="SqlDataSource1"
            DefaultMode="Insert">
            <EditItemTemplate>
                Num:
                <asp:Label ID="NumLabel1" runat="server" Text='<%# Eval("Num") %>'></asp:Label><br />
                Name:
                <asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>'>
                </asp:TextBox><br />
                Email:
                <asp:TextBox ID="EmailTextBox" runat="server" Text='<%# Bind("Email") %>'>
                </asp:TextBox><br />
                Title:
                <asp:TextBox ID="TitleTextBox" runat="server" Text='<%# Bind("Title") %>'>
                </asp:TextBox><br />
                PostDate:
                <asp:TextBox ID="PostDateTextBox" runat="server" Text='<%# Bind("PostDate") %>'>
                </asp:TextBox><br />
                PostIP:
                <asp:TextBox ID="PostIPTextBox" runat="server" Text='<%# Bind("PostIP") %>'>
                </asp:TextBox><br />
                <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update"
                    Text="업데이트">
                </asp:LinkButton>
                <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                    Text="취소">
                </asp:LinkButton>
            </EditItemTemplate>
            <InsertItemTemplate>
                Name:
                <asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>'>
                </asp:TextBox><br />
                Email:
                <asp:TextBox ID="EmailTextBox" runat="server" Text='<%# Bind("Email") %>'>
                </asp:TextBox><br />
                Title:
                <asp:TextBox ID="TitleTextBox" runat="server" Text='<%# Bind("Title") %>'>
                </asp:TextBox><br />
                PostDate:
                <asp:TextBox ID="PostDateTextBox" runat="server" Text='<%# Bind("PostDate") %>'>
                </asp:TextBox><br />
                PostIP:
                <asp:TextBox ID="PostIPTextBox" runat="server" Text='<%# Bind("PostIP") %>'>
                </asp:TextBox><br />
                <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                    Text="삽입">
                </asp:LinkButton>
                <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                    Text="취소">
                </asp:LinkButton>
            </InsertItemTemplate>
            <ItemTemplate>
                <table>
                    <tr>
                        <td style="width: 100px">
                            이름:</td>
                        <td style="width: 100px">
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                        </td>
                        <td style="width: 100px">
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 100px">
                        </td>
                        <td style="width: 100px">
                        </td>
                    </tr>
                </table>
                <br />
                Num:
                <asp:Label ID="NumLabel" runat="server" Text='<%# Eval("Num") %>'></asp:Label><br />
                Name:
                <asp:Label ID="NameLabel" runat="server" Text='<%# Bind("Name") %>'></asp:Label><br />
                Email:
                <asp:Label ID="EmailLabel" runat="server" Text='<%# Bind("Email") %>'></asp:Label><br />
                Title:
                <asp:Label ID="TitleLabel" runat="server" Text='<%# Bind("Title") %>'></asp:Label><br />
                PostDate:
                <asp:Label ID="PostDateLabel" runat="server" Text='<%# Bind("PostDate") %>'></asp:Label><br />
                PostIP:
                <asp:Label ID="PostIPLabel" runat="server" Text='<%# Bind("PostIP") %>'></asp:Label><br />
                <asp:LinkButton ID="EditButton" runat="server" CausesValidation="False" CommandName="Edit"
                    Text="편집"></asp:LinkButton>
                <asp:LinkButton ID="DeleteButton" runat="server" CausesValidation="False" CommandName="Delete"
                    Text="삭제"></asp:LinkButton>
                <asp:LinkButton ID="NewButton" runat="server" CausesValidation="False" CommandName="New"
                    Text="새로 만들기"></asp:LinkButton>
            </ItemTemplate>
        </asp:FormView>
    
    </div>
    </form>
</body>
</html>
