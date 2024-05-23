# Auto_PPTX_Total_Number

Đây là code để hiển thị tổng số slide trong pptx, yêu cầu: ban đầu cần sửa tất cả các slide master sao cho nó bắt đầu bằng dấu gạch dưới, giống như này: `_<#>` tại ô số trang.
Có thể thay bằng cái khác (thay trong cả code).

- Cách chạy: Bấm ALT_F11, tạo 1 module, dán code này vào, rồi khi nào cần chạy thì bấm `view>macro`, chọn hàm `UpdateSlideNumberTextBox_AllSlides` bấm run
- 

```vba
Sub UpdateSlideNumberTextBox_AllSlides()

  Dim oSlide As Slide
  Dim oShape As Shape
  Dim foundTextBox As Boolean

  For Each oSlide In ActivePresentation.Slides ' Loop through ALL slides
    foundTextBox = False ' Flag to track if a text box found

    For Each oShape In oSlide.Shapes ' Loop through shapes on the slide
      On Error Resume Next ' Handle potential errors gracefully

      ' Check if the shape is a text box or placeholder
      If oShape.Type = 17 Or oShape.Type = 14 Then
        ' Check for existing text containing an underscore ("_")
        If InStr(oShape.TextFrame.TextRange.Text, "_") = 1 Then
          foundTextBox = True ' Mark text box found
          oShape.TextFrame.TextRange.Text = "_" & oSlide.SlideIndex & "/" & ActivePresentation.Slides.Count & "_"
          Exit For ' Exit loop after finding and updating the text box
        End If
      End If

      On Error GoTo 0 ' Reset error handling for the next shape

    Next oShape

    ' If no text box is found, insert a new one (optional)
    If Not foundTextBox Then
      ' Insert new text box code (replace with your desired position and size)
      ' Set oShape = oSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 490, 323, 28)
      ' oShape.TextFrame.TextRange.Text = "Slide " & oSlide.SlideIndex & " of " & ActivePresentation.Slides.Count
    End If

  Next oSlide

End Sub
```
