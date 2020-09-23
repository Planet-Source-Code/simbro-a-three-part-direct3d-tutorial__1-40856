VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   1680
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DX As New DirectX8             'DirectX8 is the main object. Everything
                                   'else is based from this.
Dim D3D As Direct3D8               'Covers everything that is 3D.
Dim D3DDevice As Direct3DDevice8   'Stands for hardware that renders the
                                   'project.

Private Sub Form_Load()
    Dim b As Boolean
    
    Me.Show
    DoEvents
    
    b = InitD3D(Form1.hWnd)
    If Not b Then
        MsgBox "Unable to Create Device see InitD3D() source for comments)"
        End
    End If
    
    Timer1.Enabled = True
    
End Sub

Public Function InitD3D(hWnd As Long) As Boolean
On Error GoTo ErrHandler            'If something goes wrong go to errhandler

    Set D3D = DX.Direct3DCreate()   'creates D3D Object
    If D3D Is Nothing Then Exit Function 'if object isn't created exit sub

    Dim mode As D3DDISPLAYMODE      'Dim variable to find display mode
    'Retrieve the Display mode format
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode

    'Dims d3dpp as a D3DPRESENT_PARAMETERS object. Each field is filled in to set
    'up for a new Direct3D Device.
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1              'Windowed state set to 1 (true)
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = mode.Format

    'Creates D3DDevice
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If D3DDevice Is Nothing Then Exit Function

    InitD3D = True                  'Everything went fine
    Exit Function
ErrHandler:                         'Something went wrong
    InitD3D = False                 'Set function to false so program stops running
End Function

Private Sub Render()
    
    If D3DDevice Is Nothing Then Exit Sub
    'Clears the current frame
    'This is what makes the screen black when you run it. The "&H0&" is HTML code for black
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0&, 1#, 0

    'We begin the new scene
    D3DDevice.BeginScene

    'Here is where all the rendering would take place.
    
    'We end the Scene
    D3DDevice.EndScene

    'Here is where we send the rendered scene onto the screen
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cleanup
    End
End Sub

Private Sub Cleanup()
    Set D3D = Nothing                   'Set D3D variable to Nothing
    Set D3DDevice = Nothing             'Set D3DDevice to nothing
End Sub


Private Sub Timer1_Timer()
    Render
End Sub

