VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServiceCenter 
   Caption         =   "CenterPoint Standard Offer Service Center"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   OleObjectBlob   =   "frmServiceCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServiceCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdProgram_Click()
    Me.Hide
    frmProgram.Show vbModeless
End Sub

Private Sub cmdReport_Click()
    Me.Hide
    frmReport.Show vbModeless
End Sub

Private Sub cmdSponsor_Click()
    Me.Hide
    frmSponsor.Show vbModeless
End Sub

Private Sub cmdTools_Click()
    Me.Hide
    frmTools.Show vbModeless
End Sub

Private Sub cmdTrakSmart_Click()
    Me.Hide
    frmTrakSmart.Show vbModeless
End Sub


