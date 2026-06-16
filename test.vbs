Set o = CreateObject("MSHTML.HTMLDocument")
Set oDoc = o.createDocumentFromUrl(sUrl)
Do until o.readyState = "complete"
sleep 1
loop
