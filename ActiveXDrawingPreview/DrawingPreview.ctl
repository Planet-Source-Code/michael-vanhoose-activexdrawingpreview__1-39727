VERSION 5.00
Begin VB.UserControl DrawingPreview 
   BackColor       =   &H80000009&
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ScaleHeight     =   2325
   ScaleWidth      =   3030
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   360
         Top             =   1800
      End
   End
End
Attribute VB_Name = "DrawingPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim sTemp As String
Public Property Get YourDrawingFile() As String
    YourDrawingFile = sTemp
End Property

Public Property Let YourDrawingFile(ByVal vNewValue As String)
    sTemp = vNewValue
    PropertyChanged "YourDrawingFile"
End Property
Private Sub Timer1_Timer()
    If YourDrawingFile <> "" Then MdlDwgPreview.PaintPreview YourDrawingFile, Pic1
        Timer1.Enabled = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   YourDrawingFile = PropBag.ReadProperty("YourDrawingFile", Extender.Name)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "YourDrawingFile", YourDrawingFile, Extender.Name
End Sub
Private Sub UserControl_Resize()
   Pic1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

