# Recolor-Picture-to-black-and-white-25-by-word-VBA
Recolor Picture to black and white 25% by word VBA
The code is as follows: 
Sub changeImageToBalckAndWhite()
'
' changeImageToBalckAndWhite 宏
'
'
For Each InlineShape In ActiveDocument.InlineShapes
InlineShape.PictureFormat.ColorType = msoPictureGrayscale

InlineShape.PictureFormat.IncrementContrast 0.1
InlineShape.PictureFormat.IncrementBrightness 0.1
InlineShape.PictureFormat.IncrementContrast 0.2
InlineShape.PictureFormat.IncrementBrightness 0.1

Next InlineShape


End Sub 
