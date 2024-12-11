VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSwitchboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ******************************
' Created by: Nick Bywater
' Created for: National Park Service, CAKN
' Application created on: 2020-June
' License: Public Domain
' ******************************

Private Sub cmdOpenTransformController_Click()
    DoCmd.OpenForm "frmTransformLabData", acNormal
End Sub

