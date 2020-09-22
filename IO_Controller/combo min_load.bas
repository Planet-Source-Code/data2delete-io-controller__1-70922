Attribute VB_Name = "Module1"
Public Sub Loadup_Minutes(combo As ComboBox)
combo.AddItem "00"
combo.AddItem "01"
combo.AddItem "02"
combo.AddItem "03"
combo.AddItem "04"
combo.AddItem "05"
combo.AddItem "06"
combo.AddItem "07"
combo.AddItem "08"
combo.AddItem "09"
Dim z As Integer
For z = 10 To 59
combo.AddItem z
Next z
End Sub
