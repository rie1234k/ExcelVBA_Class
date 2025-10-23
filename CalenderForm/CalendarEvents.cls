VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents DayLabel As MSForms.Label
Attribute DayLabel.VB_VarHelpID = -1
Private DayIndex As Long



Public Sub SetLabel(NewLabel As MSForms.Label, NewIndex As Long)

    Set DayLabel = NewLabel
    
    DayIndex = NewIndex

End Sub


Private Sub DayLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    CalendarForm.Controls("Label" & DayIndex).SpecialEffect = fmSpecialEffectSunken
    


End Sub

Private Sub DayLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim iDate As Date

 With CalendarForm
    
        
        '“ú•tŽæ“¾
        iDate = DateSerial(.CB_Year, .CB_Month, .Controls("Label" & DayIndex).Caption)
        
        .TB_Date.Text = iDate
        
        .Controls("Label" & DayIndex).SpecialEffect = fmSpecialEffectRaised
     
        .TB_Date.SetFocus
        
    End With

End Sub
