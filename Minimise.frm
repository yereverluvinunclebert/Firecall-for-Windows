VERSION 5.00
Begin VB.Form MinimiseForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "minimiseForm"
   ClientHeight    =   1380
   ClientLeft      =   150
   ClientTop       =   9000
   ClientWidth     =   1155
   Icon            =   "Minimise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Minimise.frx":000C
   ScaleHeight     =   1380
   ScaleWidth      =   1155
   ShowInTaskbar   =   0   'False
   Begin VB.Timer pulseTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   390
      Top             =   510
   End
End
Attribute VB_Name = "MinimiseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : MinimiseForm
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit
'@ModuleAttribute VB_Creatable, False
'@ModuleAttribute VB_Exposed, False
'@PredeclaredId
'@ModuleAttribute VB_Name, "MinimiseForm"
'@ModuleAttribute VB_GlobalNameSpace, False

Private minFormPositionX As Single
Private minFormPositionY As Single

Private imageCounter As Integer


Private Sub Form_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
            FCWMinimiseFormX = Str$(MinimiseForm.Left)
            FCWMinimiseFormY = Str$(MinimiseForm.Top)
            
            PutINISetting "Software\FireCallWin", "minimiseFormX", FCWMinimiseFormX, FCWSettingsFile
            PutINISetting "Software\FireCallWin", "minimiseFormY", FCWMinimiseFormY, FCWSettingsFile

End Sub

Private Sub pulseTimer_Timer()
    Dim fullPath As String
    
    If MinimiseForm.Visible = True Then
        If inputDataChangedFlag = True Then
            imageCounter = imageCounter + 1
            If imageCounter >= 30 Then imageCounter = 1
            
            fullPath = App.Path & "\resources\images\" & "fireCall" & imageCounter & ".jpg"
            
            If fFExists(fullPath) Then
                MinimiseForm.Picture = LoadPicture(fullPath)
            End If
            
        Else
            imageCounter = 0
            fullPath = App.Path & "\resources\images\" & "fireCall1.jpg"
            If fFExists(fullPath) Then
                MinimiseForm.Picture = LoadPicture(fullPath)
            End If
            pulseTimer.Enabled = False
        End If
    End If
    
End Sub

Private Sub Form_DblClick()

    Call FormDblClickSub
    
End Sub




Public Sub FormDblClickSub()

    ' if the program is minimised, maximise it
    FireCallMain.opacityFadeInTimer.Enabled = True
    'If FireCallMain.Visible = False Then
    If FireCallMain.WindowState = vbMinimized Then
        FireCallMain.opacityFadeInTimer.Enabled = True
        MinimiseForm.Visible = False
        
        
        If Val(FCWIconiseDelay) > 0 Then
                If fInIDE Then
                    FireCallMain.iconiseTimer.Enabled = True 'restart the VB6 timer to iconise the program when needed
                Else
                    Call initiateIconiseTimerInCode 'restart the code timer to iconise the program when needed
                End If
        End If
    End If


'    FireCallMain.picTextChangeBright.Visible = False
'    FireCallMain.picTextChangeDull.Visible = True

    inputDataChangedFlag = False
    
    
    
End Sub

Private Sub Form_Load()
    imageCounter = 0

    pulseTimer.Enabled = False
   
    If FCWMinimiseFormX = "0" Then
        MinimiseForm.Left = FireCallMain.Left - 300
    Else
        MinimiseForm.Left = Val(FCWMinimiseFormX)
    End If
    
    If FCWMinimiseFormY = "0" Then
        MinimiseForm.Top = FireCallMain.Top
    Else
        MinimiseForm.Top = Val(FCWMinimiseFormY)
    End If
    
End Sub

Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    minFormPositionX = x
    minFormPositionY = y
    
    If Button = 2 Then
        ' use the menu from the specialised menu form to avoid generating a title bar
        PopupMenu MinimiseMenuForm.minMenuPopUp, vbPopupMenuRightButton
    End If
End Sub

Private Sub Form_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If Button = 1 Then
        With Me
            .Left = .Left - (minFormPositionX - x)
            .Top = .Top - (minFormPositionY - y)

        End With
    End If
End Sub
