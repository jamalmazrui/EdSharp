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
		ValidationExpression="((1 )?\(\d{3}\) \d{3}-\d{4})|(1-)?(\d{3}-){2}\d{4}|(1\.)?(\d{3}\.){2}\d{4}|1?\d{10}"></asp:RegularExpressionValidator>
	<asp:Button Id="btnSubmit" RunAt="server" CausesValidation="True"
		Text="Submit"></asp:Button>
	</form>
</body>
