Sub MyCrop()
Dim oshp As Shape
Set oshp = ActiveWindow.Selection.ShapeRange(1)
If oshp.Type = msoPicture Then
'is it a picture?
With oshp.PictureFormat
.CropLeft = in2Points(0#)
.CropRight = in2Points(4.5)
.CropTop = in2Points(0)
.CropBottom = in2Points(0#)
End With
End If
'is it a placeholder with a picture in it?
If oshp.Type = msoPlaceholder Then
If oshp.PlaceholderFormat.ContainedType = msoPicture Then
With oshp.PictureFormat
.CropLeft = in2Points(0.28)
.CropRight = in2Points(0.47)
.CropTop = in2Points(1.34)
.CropBottom = in2Points(0.22)
End With
End If
End If
oshp.Width = in2Points(7.9)
oshp.Height = in2Points(6.17)
'you will probably want to set the left and top too
End Sub

Function in2Points(inVal As Single) As Single
in2Points = inVal * 72
End Function
