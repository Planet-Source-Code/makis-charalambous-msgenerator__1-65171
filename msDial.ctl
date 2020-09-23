VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl msDial 
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   ScaleHeight     =   765
   ScaleWidth      =   780
   Begin VB.PictureBox knob1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   0
      Width           =   525
   End
   Begin PicClip.PictureClip picKnob 
      Left            =   2310
      Top             =   180
      _ExtentX        =   979
      _ExtentY        =   72628
      _Version        =   393216
      Rows            =   61
      Picture         =   "msDial.ctx":0000
   End
End
Attribute VB_Name = "msDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim oldX As Single, oldY As Single
Dim i As Integer, a As Single, lastVal As Integer, INN As Integer, OTN As Integer

Event DialChange(nValue As Integer)
'Default Property Values:
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Integer

Private Sub Knob1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    
        If Y < knob1.ScaleHeight / 2 Then
            If X > oldX Then m_Value = m_Value + 1
            If X < oldX Then m_Value = m_Value - 1
        Else
            If X > oldX Then m_Value = m_Value - 1
            If X < oldX Then m_Value = m_Value + 1
        End If
        
        If X < knob1.ScaleWidth / 2 Then
            If Y > oldY Then m_Value = m_Value - 1
            If Y < oldY Then m_Value = m_Value + 1
        Else
            If Y > oldY Then m_Value = m_Value + 1
            If Y < oldY Then m_Value = m_Value - 1
        End If
           
        If m_Value > 100 Then m_Value = 100
        If m_Value < 0 Then m_Value = 0
        
        RaiseEvent DialChange(m_Value)
        
        knob1.Picture = picKnob.GraphicCell(m_Value * 60 / 100)
        
        oldX = X
        oldY = Y
        
    End If
End Sub

Private Sub UserControl_Initialize()
    
    knob1.Picture = picKnob.GraphicCell(0)
    UserControl.Width = 555
    UserControl.Height = 555
    knob1.Height = 555
    knob1.Width = 555
    
End Sub

Private Sub UserControl_Resize()
    
    UserControl.Width = 555
    UserControl.Height = 555
    knob1.Height = 555
    knob1.Width = 555
    
End Sub
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    If m_Value > 100 Then m_Value = 0
    If m_Value < 0 Then m_Value = 100
    knob1.Picture = picKnob.GraphicCell(m_Value * 60 / 100)
    PropertyChanged "Value"
End Property
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

