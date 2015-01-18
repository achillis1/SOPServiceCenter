VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSponsor 
   Caption         =   "Sponsor"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11100
   OleObjectBlob   =   "frmSponsor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSponsor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdServiceCenter_Click()
    Me.Hide
    frmServiceCenter.Show vbModeless
End Sub
