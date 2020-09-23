VERSION 5.00
Begin VB.Form frmMatrices 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "frmMatrices"
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
Dim VB As Direct3DVertexBuffer8

Private Type CUSTOMVERTEX
    x As Single
    y As Single
    z As Single
    color As Long
End Type

Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)
Const pi = 3.1415

Private Sub Form_Load()
    Dim b As Boolean
    
    Me.Show
    DoEvents
    
    b = InitD3D(frmMatrices.hWnd)
    If Not b Then
        MsgBox "Unable to Create Device see InitD3D() source for comments)"
        End
    End If
    
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
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
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    InitD3D = True                  'Everything went fine
    Exit Function
ErrHandler:                         'Something went wrong
    InitD3D = False                 'Set function to false so program stops running
End Function

Sub SetupMatrices()
    
        Dim matWorld As D3DMATRIX
        D3DXMatrixRotationY matWorld, Timer * 4
        D3DDevice.SetTransform D3DTS_WORLD, matWorld
        
        Dim matView As D3DMATRIX
        D3DXMatrixLookAtLH matView, vec3(0#, 3#, -5#), _
                                     vec3(0#, 0#, 0#), _
                                     vec3(0#, 1#, 0#)
        
        D3DDevice.SetTransform D3DTS_VIEW, matView
        
        Dim matProj As D3DMATRIX
        D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 1, 1000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
        
        
End Sub

Function InitGeometry() As Boolean

    Dim Vertices(2) As CUSTOMVERTEX
    Dim VertexSizeInBytes As Long

    VertexSizeInBytes = Len(Vertices(0))

    With Vertices(0): .x = -1: .y = -1: .z = 0: .color = &HFFFF0000: End With
    With Vertices(1): .x = 1: .y = -1: .z = 0: .color = &HFF00FF00: End With
    With Vertices(2): .x = 0: .y = 1: .z = 0: .color = &HFF00FFFF: End With

    Set VB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If VB Is Nothing Then Exit Function

    D3DVertexBuffer8SetData VB, 0, VertexSizeInBytes * 3, 0, Vertices(0)

    InitGeometry = True
End Function

Private Sub Render()
    
    Dim v As CUSTOMVERTEX
    Dim sizeofVertex As Long
    
    
    If D3DDevice Is Nothing Then Exit Sub
    'Clears the current frame
    'This is what makes the screen black when you run it. The "&H0&" is HTML code for black
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0&, 1#, 0

    'We begin the new scene
    D3DDevice.BeginScene

    SetupMatrices

    'Here is where all the rendering would take place.
    sizeofVertex = Len(v)
    D3DDevice.SetStreamSource 0, VB, sizeofVertex
    D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
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
    Set VB = Nothing
    Set D3D = Nothing                   'Set D3D variable to Nothing
    Set D3DDevice = Nothing             'Set D3DDevice to nothing
End Sub


Private Sub Timer1_Timer()
    Render
End Sub

Function vec3(x As Single, y As Single, z As Single) As D3DVECTOR
    vec3.x = x
    vec3.y = y
    vec3.z = z
End Function
