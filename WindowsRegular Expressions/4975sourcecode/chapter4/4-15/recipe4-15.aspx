<%@ Page Language="vb" AutoEventWireup="false" %> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html> 
<head><title></title>
</head>
<body>
    <form Id="Form1" RunAt="server">
    <asp:TextBox id="txtInput" runat="server"></asp:TextBox>
    <asp:RegularExpressionValidator Id="revInput" RunAt="server" 
        ControlToValidate="txtInput"
        ErrorMessage="Please enter a valid value"
        ValidationExpression="^(0?2/(0?[1-9]|[12][0-9])|(0?[469]|11)/(0?[1-9]|[12][0-9]|30)|(0?[13578]|1[02])/(0?[1-9]|[12][0-9]|3[01]))/\d{4}$">
    </asp:RegularExpressionValidator>
    <asp:Button Id="btnSubmit" RunAt="server" CausesValidation="True"
        Text="Submit"></asp:Button>
    </form>
</body>
