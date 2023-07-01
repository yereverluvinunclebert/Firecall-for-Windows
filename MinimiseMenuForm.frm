VERSION 5.00
Begin VB.Form MinimiseMenuForm 
   BorderStyle     =   0  'None
   ClientHeight    =   3345
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "MinimiseMenuForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu minMenuPopUp 
      Caption         =   "Minimise Menu Holder"
      Begin VB.Menu mnuOpenProgram 
         Caption         =   "Re-open the Main Program Window"
      End
      Begin VB.Menu mnuBringToCentre 
         Caption         =   "Re-open Program and Centre on Main Monitor"
      End
      Begin VB.Menu mnuCloseProgram 
         Caption         =   "Close the Main Program"
      End
   End
End
Attribute VB_Name = "MinimiseMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : MinimiseMenuForm
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------

'@ModuleAttribute VB_Creatable, False
'@ModuleAttribute VB_Name, "MinimiseMenuForm"
'@PredeclaredId
'@ModuleAttribute VB_GlobalNameSpace, False
'@ModuleAttribute VB_Exposed, False

Option Explicit




Private Sub mnuBringToCentre_Click()
    Call centreMainScreen
    Call MinimiseForm.FormDblClickSub

    
End Sub

Private Sub mnuCloseProgram_Click()
    'Call Form_Unload_Sub
    Unload FireCallMain
End Sub


Private Sub mnuOpenProgram_Click()
    Call MinimiseForm.FormDblClickSub
End Sub
