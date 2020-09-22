Attribute VB_Name = "modProg"
' feel free to use this anyway you see fit, I do ask
' that if you modify the code in anyway to enhance it
' please email me a copy so I can check it out. no need
' to give me credit if you use this
' voting for me would be nice  :)


Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)


' make sure that when the form is resized that the
' statusbar is rsized before we continue
sb.Align = 2
sb.Refresh

' set the properties of the progressbar
' flat with no border seems to look the best
' also set the progressbar to the top of the zorder
pb.ZOrder 0
pb.Appearance = ccFlat
pb.BorderStyle = ccNone

' now resize the progressbar1 to fit in the statusbar panel
pb.Left = sb.Panels(lPan).Left + 25
pb.Width = sb.Panels(lPan).Width - 45
pb.Top = sb.Top + 45
pb.Height = sb.Height - 75


End Function
