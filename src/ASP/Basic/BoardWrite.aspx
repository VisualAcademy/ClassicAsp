<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BoardWrite.aspx.cs" Inherits="Basic_BoardWrite" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>제목 없음</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FormView ID="FormView1" runat="server" DataKeyNames="Num" DataSourceID="SqlDataSource1"
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
                PostIP:
                <asp:TextBox ID="PostIPTextBox" runat="server" Text='<%# Bind("PostIP") %>'>
                </asp:TextBox><br />
                Content:
                <asp:TextBox ID="ContentTextBox" runat="server" Text='<%# Bind("Content") %>'>
                </asp:TextBox><br />
                Password:
                <asp:TextBox ID="PasswordTextBox" runat="server" Text='<%# Bind("Password") %>'>
                </asp:TextBox><br />
                Encoding:
                <asp:TextBox ID="EncodingTextBox" runat="server" Text='<%# Bind("Encoding") %>'>
                </asp:TextBox><br />
                Homepage:
                <asp:TextBox ID="HomepageTextBox" runat="server" Text='<%# Bind("Homepage") %>'>
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
                PostIP:
                <asp:TextBox ID="PostIPTextBox" runat="server" Text='<%# Bind("PostIP") %>'>
                </asp:TextBox><br />
                Content:
                <asp:TextBox ID="ContentTextBox" runat="server" Text='<%# Bind("Content") %>'>
                </asp:TextBox><br />
                Password:
                <asp:TextBox ID="PasswordTextBox" runat="server" Text='<%# Bind("Password") %>'>
                </asp:TextBox><br />
                Encoding:
                <asp:TextBox ID="EncodingTextBox" runat="server" Text='<%# Bind("Encoding") %>'>
                </asp:TextBox><br />
                Homepage:
                <asp:TextBox ID="HomepageTextBox" runat="server" Text='<%# Bind("Homepage") %>'>
                </asp:TextBox><br />
                <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                    Text="삽입">
                </asp:LinkButton>
                <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                    Text="취소">
                </asp:LinkButton>
            </InsertItemTemplate>
            <ItemTemplate>
                Num:
                <asp:Label ID="NumLabel" runat="server" Text='<%# Eval("Num") %>'></asp:Label><br />
                Name:
                <asp:Label ID="NameLabel" runat="server" Text='<%# Bind("Name") %>'></asp:Label><br />
                Email:
                <asp:Label ID="EmailLabel" runat="server" Text='<%# Bind("Email") %>'></asp:Label><br />
                Title:
                <asp:Label ID="TitleLabel" runat="server" Text='<%# Bind("Title") %>'></asp:Label><br />
                PostIP:
                <asp:Label ID="PostIPLabel" runat="server" Text='<%# Bind("PostIP") %>'></asp:Label><br />
                Content:
                <asp:Label ID="ContentLabel" runat="server" Text='<%# Bind("Content") %>'></asp:Label><br />
                Password:
                <asp:Label ID="PasswordLabel" runat="server" Text='<%# Bind("Password") %>'></asp:Label><br />
                Encoding:
                <asp:Label ID="EncodingLabel" runat="server" Text='<%# Bind("Encoding") %>'></asp:Label><br />
                Homepage:
                <asp:Label ID="HomepageLabel" runat="server" Text='<%# Bind("Homepage") %>'></asp:Label><br />
            </ItemTemplate>
        </asp:FormView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:BasicConnectionString %>"
            DeleteCommand="DELETE FROM [Basic] WHERE [Num] = @Num" InsertCommand="INSERT INTO [Basic] ([Name], [Email], [Title], [PostIP], [Content], [Password], [Encoding], [Homepage]) VALUES (@Name, @Email, @Title, @PostIP, @Content, @Password, @Encoding, @Homepage)"
            SelectCommand="SELECT [Num], [Name], [Email], [Title], [PostIP], [Content], [Password], [Encoding], [Homepage] FROM [Basic]"
            UpdateCommand="UPDATE [Basic] SET [Name] = @Name, [Email] = @Email, [Title] = @Title, [PostIP] = @PostIP, [Content] = @Content, [Password] = @Password, [Encoding] = @Encoding, [Homepage] = @Homepage WHERE [Num] = @Num">
            <DeleteParameters>
                <asp:Parameter Name="Num" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="Name" Type="String" />
                <asp:Parameter Name="Email" Type="String" />
                <asp:Parameter Name="Title" Type="String" />
                <asp:Parameter Name="PostIP" Type="String" />
                <asp:Parameter Name="Content" Type="String" />
                <asp:Parameter Name="Password" Type="String" />
                <asp:Parameter Name="Encoding" Type="String" />
                <asp:Parameter Name="Homepage" Type="String" />
                <asp:Parameter Name="Num" Type="Int32" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="Name" Type="String" />
                <asp:Parameter Name="Email" Type="String" />
                <asp:Parameter Name="Title" Type="String" />
                <asp:Parameter Name="PostIP" Type="String" />
                <asp:Parameter Name="Content" Type="String" />
                <asp:Parameter Name="Password" Type="String" />
                <asp:Parameter Name="Encoding" Type="String" />
                <asp:Parameter Name="Homepage" Type="String" />
            </InsertParameters>
        </asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
