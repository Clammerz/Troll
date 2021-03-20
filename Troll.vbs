Set Troll = CreateObject("InternetExplorer.Application")

With Troll
    .Navigate "about:blank"
    .ToolBar = 0
    .StatusBar = 0
    .Left = 100
    .Top = 100
    .Width = 200
    .Height = 200
    .Visible = 1
    .Document.Title = "Trolled XD !!!"
    .Document.Body.InnerHTML = _
        "<img src='https://i.imgur.com/aSwhVbb.png' height=100 width=100>"
End With
