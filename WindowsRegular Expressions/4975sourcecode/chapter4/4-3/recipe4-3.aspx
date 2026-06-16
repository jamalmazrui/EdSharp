<%@ Page Language="vb" AutoEventWireup="false" Debug="true"%> 
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
		ValidationExpression="((0[1-9]|[12][0-9]|30)-(Sep|Apr|Jun|Nov)|(0[1-9]|[12][0-9])-Feb|(0[1-9]|[12][0-9]|3[01])-(Jan|Mar|May|Jul|Aug|Oct|Dec))-\d{4}">
		</asp:RegularExpressionValidator>
	<asp:Button Id="btnSubmit" RunAt="server" CausesValidation="True"
		Text="Submit"></asp:Button>
	</form>
</body>
