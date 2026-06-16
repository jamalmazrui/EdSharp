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
		ValidationExpression="^(\d{4}-){3}|(\d{4} ){3}\d{4}|\d{15,16}|\d{4} \d{2} \d{4} \d{5}|\d{4}-\d{2}-\d{4}-\d{5}$"></asp:RegularExpressionValidator>
	<asp:Button Id="btnSubmit" RunAt="server" CausesValidation="True"
		Text="Submit"></asp:Button>
	</form>
</body>
