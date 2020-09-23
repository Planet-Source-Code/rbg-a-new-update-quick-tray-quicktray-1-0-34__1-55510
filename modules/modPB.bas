Attribute VB_Name = "modPB"
Option Explicit

Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)
 sb.Align = 2
 sb.Refresh
 pb.ZOrder 0
 pb.Appearance = ccFlat
 pb.BorderStyle = ccNone
 pb.Left = sb.Panels(lPan).Left + 25
 pb.Width = sb.Panels(lPan).Width - 45
 pb.Top = sb.Top + 45
 pb.Height = sb.Height - 75
 pb.Visible = True
End Function
