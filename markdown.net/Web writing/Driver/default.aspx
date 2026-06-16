<%@ Page language="c#" Codebehind="default.aspx.cs" AutoEventWireup="false" Inherits="anrControls.MarkdownDriver._default" ValidateRequest="false" EnableViewState="false" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<body>
<form id="Form1" method="post" runat="server">

<h2>Markdown.NET</h2>

<p>To learn more about Markdown.NET, please read 
   <a href="http://www.aspnetresources.com/blog/markdown_announced.aspx">Announcing Markdown.NET</a>.
</p>

<div>
    <p>Markdown source:</p>
    <asp:TextBox Rows="20" style="width: 490px; overflow: auto;" runat="server" TextMode="MultiLine" id="MarkdownSource" />
</div>

<p>
    Filter:
    <asp:DropDownList id="FilterOptions" Runat="server">
        <asp:ListItem>Markdown</asp:ListItem>
        <asp:ListItem>SmartyPants</asp:ListItem>
        <asp:ListItem>Both</asp:ListItem>
    </asp:DropDownList>
    
    &nbsp;
    
    Results: 
    <asp:DropDownList id="OutputOptions" Runat="server">
        <asp:ListItem>Source</asp:ListItem>
        <asp:ListItem>Preview</asp:ListItem>
        <asp:ListItem>Source &amp; preview</asp:ListItem>
    </asp:DropDownList>
    <asp:Button id="Convert" Runat="server" Text="Convert" />
</p>

<asp:PlaceHolder id="plhHtmlsource" Runat="server" Visible="false">
<div>
    <p>HTML source:</p>
    <asp:TextBox Rows="20" style="width: 490px; overflow: auto; " runat="server" TextMode="MultiLine" id="HtmlSource" />
</div>
</asp:PlaceHolder>

<asp:PlaceHolder id="plhHtmlPreview" Runat="server" Visible="false">
<div>
    <p>HTML preview:</p>
    <asp:Label Runat="server" style="width: 480px; border: 1px solid #ccc; display: block; background: #fff; padding: 5px;" id="Preview" />
</div>
</asp:PlaceHolder>

<p style="font-size: 0.85em; margin-top: 4em;">
  Markdown 1.0.1 / SmartyPants 1.5.1<br />
  Copyright © 2004 John Gruber
</p>

<p style="font-size: 0.85em;">
   Markdown.NET 1.0.1<br />
   Copyright © 2004-2005 Milan Negovan
</p>

</form>
</body>
</html>