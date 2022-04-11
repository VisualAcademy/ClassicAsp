<table border="1" width="500">
    <tr>
        <td colspan="2">
        <% Server.Execute("Navigator.asp") %>
        </td>
    </tr>
    <tr style="height:250px;">
        <td>
        <% Server.Execute("Category.inc") %>
        </td>
        <td>
        <!-- #include file="./Content.asp" -->
        </td>
    </tr>
    <tr>
        <td colspan="2">
        <!-- #include virtual="/ASP/SSI/Copyright.txt" -->
        </td>
    </tr>
</table>