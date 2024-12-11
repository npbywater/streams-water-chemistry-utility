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
