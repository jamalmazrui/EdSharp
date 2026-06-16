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
		ValidationExpression="^(00[1-9]|0[1-9]\d|[1-5]\d{2}|6[0-5]\d|66[0-5]|66[7-9]|6[7-8]\d|690|7[0-2]\d|73[0-3]|750|76[4-9]|77[0-2])-(?!00)\d{2}-(?!0000)\d{4}$"></asp:RegularExpressionValidator>
	<asp:Button Id="btnSubmit" RunAt="server" CausesValidation="True"
		Text="Submit"></asp:Button>
	</form>
</body>
