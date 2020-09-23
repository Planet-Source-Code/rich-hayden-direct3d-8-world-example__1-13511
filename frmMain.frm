VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "D3DScene"
   ClientHeight    =   13065
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   13065
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   14775
      Left            =   0
      ScaleHeight     =   14715
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.Timer tmrFps 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5760
         Top             =   3120
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS PROGRAM MAY NOT WORK ON MACHINES WITHOUT GOOD GRAPHICS CARDS, IF IT DOESN'T THEN YOU CAN TRY CHANGING, THE FOLLOWING LINE:
'Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, d3dpp)
'TO
'Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
'THIS SHOULD FIX IT, ALTHOUGH PERFORMANCE WILL BE VERY POOR, AND VERY BUGGY
'
'D3DScene Copyright (c) 2000 Richard Hayden. All Rights Reserved.
'If you use any of this code in your programs, please acknowledge me in your code.
'
'I must also acknowledge Simon Price (http://www.vbgames.co.uk) who has introduced me to D3D in Vb, with his excellent tutorials, and has also helped me on some aspects of this program,
'
'Cheers, Simon!
'
'If anyone can help me with Collision Detection and/or transparency in textures, ie. colour keys etc., please e-mail me on r_hayden@breathemail.net

Option Explicit

Dim g_DX As New DirectX8            ' mother of it all
Dim g_D3DX As New D3DX8
Dim g_D3D As Direct3D8              ' used to create the D3DDevice
Dim g_D3DDevice As Direct3DDevice8  ' rendering device
Dim g_Grass, g_Sky1, g_Sky2, g_Sky3, g_Sky4, g_Sky5, g_Sky6, g_House1, g_House2, g_House3, g_House4, g_HouseRoof1, g_HouseRoof2, g_HouseRoof3, g_HouseRoof4, g_Floor, g_Gorilla, g_Road, g_RoadSide1, g_RoadSide2, g_RoadMiddle As Direct3DVertexBuffer8   ' holds vertex data
Dim g_TGrass, g_TBricks, g_TSky, g_TRoof, g_TFloor, g_TGorilla, g_TRoad As Direct3DTexture8   ' textures

Dim jumping As Boolean 'boolean to tell whether the camera is jumping or not
Dim jUP As Boolean 'which way is the camera going, in terms of jumping; up or down

Dim di As DirectInput8 'this is DirectInput, used to monitor the keys on the keyboard in my case
Dim diDEV As DirectInputDevice8 'this device will be the keyboard
Dim diState As DIKEYBOARDSTATE 'to check the state of keys

Dim fps As Integer 'frames/sec

Dim Angle As Single 'holds the angle, at which the camera is pointing
Dim pitch As Single 'holds the pitch of the camera (this is where the camera is pointing in terms of the y axis, ie. up and down etc.)

Dim camz, camx, camy As Single 'hold the position of the camera on the x, y and z axis

' a structure for custom vertex type
Private Type CUSTOMVERTEX
    position As D3DVECTOR    '3d position for vertex
    Color As Long           'color of the vertex
    tu As Single            'texture map coordinate
    tv As Single            'texture map coordinate
End Type

' custom FVF, which describes our custom vertex structure
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Const g_pi As Single = 3.141592653 'pi
Const g_90d As Single = g_pi / 2 '90 degrees in radians
Const g_180d As Single = g_pi '180 degrees in radians
Const g_270d As Single = (g_pi / 2) * 3 '270 degrees in radians
Const g_360d As Single = g_pi * 2 '360 degrees in radians

Const TURN_SPEED = g_90d / 18 'camera turning speed
Const MOVE_SPEED = 0.5 'camera moving speed
Const FAST_MOVE_SPEED = 0.8 'fast camera moving speed
Const JUMP_MOVE_SPEED = 1.2 'jumping camera moving speed
Const JUMP_SPEED = 1 'jumping speed
Const PITCH_SPEED = 0.2 'look up and down speed

Private Sub Form_Resize()
    'resize the picbox to utilise full size of form
    Picture1.Width = frmMain.Width
    Picture1.Height = frmMain.Height
End Sub

Private Sub SortPix(strWhat As String)
    On Error Resume Next
    'create and delete pix from pic boxes
    If strWhat = "create" Then
        SavePicture frmTextures.picGrass.Picture, App.Path & "\grass.bmp"
        SavePicture frmTextures.picBricks.Picture, App.Path & "\bricks.bmp"
        SavePicture frmTextures.picSky.Picture, App.Path & "\sky.bmp"
        SavePicture frmTextures.picRoof.Picture, App.Path & "\roof.bmp"
        SavePicture frmTextures.picTile.Picture, App.Path & "\tile.bmp"
        SavePicture frmTextures.picGorilla.Picture, App.Path & "\gorilla.bmp"
        SavePicture frmTextures.picAsphalt.Picture, App.Path & "\asphalt.bmp"
    ElseIf strWhat = "delete" Then
        Kill App.Path & "\grass.bmp"
        Kill App.Path & "\bricks.bmp"
        Kill App.Path & "\sky.bmp"
        Kill App.Path & "\roof.bmp"
        Kill App.Path & "\tile.bmp"
        Kill App.Path & "\gorilla.bmp"
        Kill App.Path & "\asphalt.bmp"
    End If
    Err.Number = 0
End Sub

Private Sub Form_Load()
    Dim b As Boolean
    
    ' Allow the form to become visible
    DoEvents
    'make the pix
    SortPix "create"
    'starting position + angle
    camx = 0
    camy = 10
    camz = -1
    Angle = g_360d
    'maximising form
    frmMain.Width = Screen.Width
    frmMain.Height = Screen.Height
    frmMain.Top = 0
    frmMain.Left = 0
    Picture1.Width = frmMain.Width
    Picture1.Height = frmMain.Height
    Picture1.Top = 0
    Picture1.Left = 0
    Me.Show
    Form1.Show
    'create directinput object
    Set di = g_DX.DirectInputCreate()
        
    If Err.Number <> 0 Then
        MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
        End
    End If
        
    'create keyboard device
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    'set common data format to keyboard
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    diDEV.Acquire
        
    
    
    ' Initialize D3D and D3DDevice
    b = InitD3D(Picture1.hWnd)
    If Not b Then
        MsgBox "Unable to CreateDevice (see InitD3D() source for comments)"
        End
    End If
    
    
    ' Initialize vertex buffer with geometry and load texture
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
        End
    End If
    
    
    'enabled fps timer to get the frames/second
    tmrFps.Enabled = True
    Do While 1
        DoEvents
        Render
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'let's cleanup
    Cleanup
    End
End Sub

Function InitD3D(hWnd As Long) As Boolean
    On Local Error Resume Next
    
    ' Create the D3D object
    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then Exit Function
    
    ' Get The current Display Mode format
    Dim Mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
         
    ' Set up the structure used to create the D3DDevice. Since we are now
    ' using more complex geometry, we will create a device with a zbuffer.
    ' the D3DFMT_D16 indicates we want a 16 bit z buffer.
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = Mode.Format
    d3dpp.BackBufferCount = 1
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16

    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, _
                                      D3DCREATE_HARDWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Exit Function
    
    ' Device state is set here
    ' Turn off culling, so we see the front and back
    g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Turn on the zbuffer
    g_D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    
    ' Turn off lighting
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    InitD3D = True
End Function

Sub SetupMatrices()
    Dim matView As D3DMATRIX
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matWorld As D3DMATRIX
    Dim matProj As D3DMATRIX
    Dim AngleConv As Single
    
    'setup world matrix
    D3DXMatrixIdentity matWorld
    g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    'get the state of the keyboard to distate
    diDEV.GetDeviceStateKeyboard diState
    
    'if turning left or right, then change angle accordingly
    If diState.Key(205) <> 0 Then
        Angle = Angle - TURN_SPEED
        If Angle < 0 Then
            Angle = g_360d - (-Angle)
        End If
    End If
    If diState.Key(203) <> 0 Then
        Angle = Angle + TURN_SPEED
        If Angle > g_360d Then
            Angle = 0 + (Angle - g_360d)
        End If
    End If
    'convert to correct angle system
    AngleConv = g_360d - Angle
    'move forward or backward if up or down keys are active
    If diState.Key(200) <> 0 Then
        'if shift is pressed, then move faster
        If Not jumping Then
            If (diState.Key(42) <> 0) Or (diState.Key(54) <> 0) Then
                camx = camx + (Sin(AngleConv) * FAST_MOVE_SPEED)
                camz = camz + (Cos(AngleConv) * FAST_MOVE_SPEED)
            Else
                camx = camx + (Sin(AngleConv) * MOVE_SPEED)
                camz = camz + (Cos(AngleConv) * MOVE_SPEED)
            End If
        Else
            camx = camx + (Sin(AngleConv) * JUMP_MOVE_SPEED)
            camz = camz + (Cos(AngleConv) * JUMP_MOVE_SPEED)
        End If
    End If
    If diState.Key(208) <> 0 Then
        'if shift is pressed, then move faster
        If Not jumping Then
            If (diState.Key(42) <> 0) Or (diState.Key(54) <> 0) Then
                camx = camx - (Sin(AngleConv) * FAST_MOVE_SPEED)
                camz = camz - (Cos(AngleConv) * FAST_MOVE_SPEED)
            Else
                camx = camx - (Sin(AngleConv) * MOVE_SPEED)
                camz = camz - (Cos(AngleConv) * MOVE_SPEED)
            End If
        Else
            camx = camx - (Sin(AngleConv) * JUMP_MOVE_SPEED)
            camz = camz - (Cos(AngleConv) * JUMP_MOVE_SPEED)
        End If
    End If
    'if pressing page up or down then look up or down, respectively
    If diState.Key(201) <> 0 Then
        pitch = pitch + PITCH_SPEED
    End If
    If diState.Key(209) <> 0 Then
        pitch = pitch - PITCH_SPEED
    End If
    If diState.Key(57) <> 0 Then
        If Not jumping Then
            jumping = True
            jUP = True
        End If
    End If
    
    'this all does the jumping stuff, quite simple......
    If jumping Then
        If jUP = False Then
            camy = camy - JUMP_SPEED
            If camy <= 10 Then
                jumping = False
                camy = 10
            End If
        Else
            camy = camy + JUMP_SPEED
            If camy >= 20 Then
                jUP = False
            End If
        End If
    End If

    'make them identity matrices
    D3DXMatrixIdentity matView
    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matRotation
    D3DXMatrixIdentity matLook
    'rotate around x and y, for angle and pitch
    D3DXMatrixRotationY matRotation, Angle
    D3DXMatrixRotationX matPitch, pitch
    'multiply angle and pitch matrices together to create one 'look' matrix
    D3DXMatrixMultiply matLook, matRotation, matPitch
    'put the position of the camera into the translation matrix, matPos
    D3DXMatrixTranslation matPos, -camx, -camy, -camz
    'multiply that with the look matrix to make the complete view matrix
    D3DXMatrixMultiply matView, matPos, matLook
    'which we can then set as the view matrix:
    g_D3DDevice.SetTransform D3DTS_VIEW, matView
    'update debug form
    Form1.Label1.Caption = "camx: " & camx & Chr(13) & "camy: " & camy & Chr(13) & "camz: " & camz & Chr(13) & "angleconv: " & AngleConv & Chr(13) & "cos(AngleConv) = " & Cos(AngleConv) & Chr(13) & "sin(AngleConv) = " & Sin(AngleConv)
    'let events happen
    DoEvents

    'setup the projection matrix
    D3DXMatrixPerspectiveFovLH matProj, g_pi / 3, 1, 1, 10000
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

Sub SetupLights()
     
    Dim col As D3DCOLORVALUE
    
    
    ' Set up a material. The material here just has the diffuse and ambient
    ' colors set to yellow. Note that only one material can be used at a time.
    Dim mtrl As D3DMATERIAL8
    With col:    .r = 1: .g = 1: .b = 0: .a = 1:   End With
    mtrl.diffuse = col
    mtrl.Ambient = col
    g_D3DDevice.SetMaterial mtrl
    
    ' Set up a white, directional light, with an oscillating direction.
    ' Note that many lights may be active at a time (but each one slows down
    ' the rendering of our scene). However, here we are just using one. Also,
    ' we need to set the D3DRS_LIGHTING renderstate to enable lighting
    
    Dim light As D3DLIGHT8
    light.Type = D3DLIGHT_DIRECTIONAL
    light.diffuse.r = 1#
    light.diffuse.g = 1#
    light.diffuse.b = 1#
    light.Direction.x = Cos(Timer * 2)
    light.Direction.y = 1#
    light.Direction.z = Sin(Timer * 2)
    light.Range = 1000#
    
    g_D3DDevice.SetLight 0, light                   'let d3d know about the light
    g_D3DDevice.LightEnable 0, 1                    'turn it on
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 1    'make sure lighting is enabled

    ' Finally, turn on some ambient light.
    ' Ambient light is light that scatters and lights all objects evenly
    g_D3DDevice.SetRenderState D3DRS_AMBIENT, &H202020
    
End Sub

Function InitGeometry() As Boolean
    Dim i As Long
    
    Set g_TBricks = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\bricks.bmp")
    If g_TBricks Is Nothing Then Exit Function
    Set g_TGrass = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\grass.bmp")
    If g_TGrass Is Nothing Then Exit Function
    Set g_TSky = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\sky.bmp")
    If g_TSky Is Nothing Then Exit Function
    Set g_TRoof = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\roof.bmp")
    If g_TSky Is Nothing Then Exit Function
    Set g_TFloor = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\tile.bmp")
    If g_TFloor Is Nothing Then Exit Function
    Set g_TGorilla = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\gorilla.bmp")
    If g_TGorilla Is Nothing Then Exit Function
    Set g_TRoad = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path + "\asphalt.bmp")
    If g_TRoad Is Nothing Then Exit Function
    
    'create an array to hold the vertex values temporarily, until added to buffer
    Dim Vertices(0 To 3) As CUSTOMVERTEX
    Dim VertexSizeInBytes As Long
    'get the size of a vertex
    VertexSizeInBytes = Len(Vertices(0))

    'create the grass or floor vertex buffer
    Vertices(0).position = vec3(-1000, 2, 1000)
    Vertices(1).position = vec3(1000, 2, 1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, -1000)
    Vertices(0).Color = &H8080FF
    Vertices(1).Color = &HFF6FFFFF
    Vertices(2).Color = &HFFFF80
    Vertices(3).Color = &HFFC0FF
    Vertices(0).tu = 0
    Vertices(1).tu = 100
    Vertices(2).tu = 0
    Vertices(3).tu = 100
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 100
    Vertices(3).tv = 100

    ' Create the vertex buffer.
    Set g_Grass = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Grass Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Grass, 0, VertexSizeInBytes * 4, 0, Vertices(0)

    'create the first sky vertex buffers
    Vertices(0).position = vec3(-1000, 1000, 1000)
    Vertices(1).position = vec3(1000, 1000, 1000)
    Vertices(2).position = vec3(-1000, 2, 1000)
    Vertices(3).position = vec3(1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Sky1 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Sky1 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Sky1, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1000, 1000, -1000)
    Vertices(1).position = vec3(1000, 1000, -1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, -1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Sky2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Sky2 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Sky2, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(1000, 1000, -1000)
    Vertices(1).position = vec3(1000, 1000, 1000)
    Vertices(2).position = vec3(1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Sky3 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Sky3 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Sky3, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1000, 1000, -1000)
    Vertices(1).position = vec3(-1000, 1000, 1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(-1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Sky4 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Sky4 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Sky4, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(1000, 1000, -1000)
    Vertices(1).position = vec3(-1000, 1000, -1000)
    'if the below '1000000 values are made 1000 as they should be then, it causes weird things to happen, maybe a bug. Making the number massive seems to make the side-affect so small that it is hardly noticeable.
    Vertices(2).position = vec3(1000, 1000, 1000)
    Vertices(3).position = vec3(-1000, 1000, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Sky5 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Sky5 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_Sky5, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the house vertex buffers
    Vertices(0).position = vec3(10, 20, 10)
    Vertices(1).position = vec3(40, 20, 10)
    Vertices(2).position = vec3(10, 2, 10)
    Vertices(3).position = vec3(40, 2, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_House1 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_House1 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_House1, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 20, 40)
    Vertices(1).position = vec3(40, 20, 40)
    Vertices(2).position = vec3(10, 2, 40)
    Vertices(3).position = vec3(40, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_House2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_House2 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_House2, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 20, 10)
    Vertices(1).position = vec3(10, 20, 40)
    Vertices(2).position = vec3(10, 2, 10)
    Vertices(3).position = vec3(10, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_House3 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_House3 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_House3, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(40, 20, 10)
    Vertices(1).position = vec3(40, 20, 40)
    Vertices(2).position = vec3(40, 2, 10)
    Vertices(3).position = vec3(40, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_House4 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_House4 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_House4, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the roof vertex buffers
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(40, 30, 25)
    Vertices(2).position = vec3(10, 20, 10)
    Vertices(3).position = vec3(40, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_HouseRoof1 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_HouseRoof1 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_HouseRoof1, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(40, 30, 25)
    Vertices(2).position = vec3(10, 20, 40)
    Vertices(3).position = vec3(40, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_HouseRoof2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_HouseRoof2 Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_HouseRoof2, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(10, 20, 10)
    Vertices(2).position = vec3(10, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_HouseRoof3 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_HouseRoof3 Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_HouseRoof3, 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    Vertices(0).position = vec3(40, 30, 25)
    Vertices(1).position = vec3(40, 20, 10)
    Vertices(2).position = vec3(40, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_HouseRoof4 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_HouseRoof4 Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_HouseRoof4, 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    Vertices(0).position = vec3(40, 2.1, 10)
    Vertices(1).position = vec3(10, 2.1, 10)
    Vertices(2).position = vec3(40, 2.1, 40)
    Vertices(3).position = vec3(10, 2.1, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_Floor = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Floor Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_Floor, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(39.9, 15, 22.5)
    Vertices(1).position = vec3(39.9, 15, 27.5)
    Vertices(2).position = vec3(39.9, 9, 22.5)
    Vertices(3).position = vec3(39.9, 9, 27.5)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_Gorilla = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Gorilla Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_Gorilla, 0, VertexSizeInBytes * 4, 0, Vertices(0)

    Vertices(0).position = vec3(-20, 2.1, -1000)
    Vertices(1).position = vec3(0, 2.1, -1000)
    Vertices(2).position = vec3(-20, 2.1, 1000)
    Vertices(3).position = vec3(0, 2.1, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 5
    Vertices(2).tu = 0
    Vertices(3).tu = 5
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1000
    Vertices(3).tv = 1000

    ' Create the vertex buffer.
    Set g_Road = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_Road Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_Road, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-20, 2.2, -1000)
    Vertices(1).position = vec3(-19, 2.2, -1000)
    Vertices(2).position = vec3(-20, 2.2, 1000)
    Vertices(3).position = vec3(-19, 2.2, 1000)
    Vertices(0).Color = &H0&
    Vertices(1).Color = &H0&
    Vertices(2).Color = &H0&
    Vertices(3).Color = &H0&

    ' Create the vertex buffer.
    Set g_RoadSide1 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_RoadSide1 Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_RoadSide1, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1, 2.2, -1000)
    Vertices(1).position = vec3(0, 2.2, -1000)
    Vertices(2).position = vec3(-1, 2.2, 1000)
    Vertices(3).position = vec3(0, 2.2, 1000)
    Vertices(0).Color = &H0&
    Vertices(1).Color = &H0&
    Vertices(2).Color = &H0&
    Vertices(3).Color = &H0&

    ' Create the vertex buffer.
    Set g_RoadSide2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_RoadSide2 Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_RoadSide2, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-9.5, 2.2, -1000)
    Vertices(1).position = vec3(-10.5, 2.2, -1000)
    Vertices(2).position = vec3(-9.5, 2.2, 1000)
    Vertices(3).position = vec3(-10.5, 2.2, 1000)
    Vertices(0).Color = &HC0C0C0
    Vertices(1).Color = &HC0C0C0
    Vertices(2).Color = &HC0C0C0
    Vertices(3).Color = &HC0C0C0

    ' Create the vertex buffer.
    Set g_RoadMiddle = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_RoadMiddle Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_RoadMiddle, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    InitGeometry = True
End Function

Sub Cleanup()
    'release all components
    Set g_TGrass = Nothing
    Set g_TBricks = Nothing
    Set g_TSky = Nothing
    Set g_TRoof = Nothing
    Set g_TFloor = Nothing
    Set g_TGorilla = Nothing
    Set g_TRoad = Nothing
    Set g_Grass = Nothing
    Set g_Sky1 = Nothing
    Set g_Sky2 = Nothing
    Set g_Sky3 = Nothing
    Set g_Sky4 = Nothing
    Set g_Sky5 = Nothing
    Set g_House1 = Nothing
    Set g_House2 = Nothing
    Set g_House3 = Nothing
    Set g_House4 = Nothing
    Set g_HouseRoof1 = Nothing
    Set g_HouseRoof2 = Nothing
    Set g_HouseRoof3 = Nothing
    Set g_HouseRoof4 = Nothing
    Set g_Floor = Nothing
    Set g_Gorilla = Nothing
    Set g_Road = Nothing
    Set g_RoadSide1 = Nothing
    Set g_RoadSide2 = Nothing
    Set g_RoadMiddle = Nothing
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
    diDEV.Unacquire
    SortPix "delete"
End Sub

Sub Render()

    Dim v As CUSTOMVERTEX
    Dim sizeOfVertex As Long
    
    
    If g_D3DDevice Is Nothing Then Exit Sub

    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
    ' Clear the z buffer to 1
    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF&, 1#, 0
    
     
    ' Begin the scene
    g_D3DDevice.BeginScene
    
    g_D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    g_D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    
    ' Setup the world, view, and projection matrices
    SetupMatrices
    
    g_D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    ' Draw the triangles in the vertex buffer
    ' Note we are now using a triangle strip of vertices
    ' instead of a triangle list
    sizeOfVertex = Len(v)

    g_D3DDevice.SetTexture 0, g_TGrass
    g_D3DDevice.SetStreamSource 0, g_Grass, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_TSky
    g_D3DDevice.SetStreamSource 0, g_Sky1, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_Sky2, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_Sky3, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_Sky4, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_Sky5, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    g_D3DDevice.SetTexture 0, g_TBricks
    g_D3DDevice.SetStreamSource 0, g_House1, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_House2, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_House3, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_House4, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_TRoof
    g_D3DDevice.SetStreamSource 0, g_HouseRoof1, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_HouseRoof2, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_HouseRoof3, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetStreamSource 0, g_HouseRoof4, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetTexture 0, g_TFloor
    g_D3DDevice.SetStreamSource 0, g_Floor, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_TGorilla
    g_D3DDevice.SetStreamSource 0, g_Gorilla, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_TRoad
    g_D3DDevice.SetStreamSource 0, g_Road, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, Nothing
    g_D3DDevice.SetStreamSource 0, g_RoadSide1, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_RoadSide2, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_RoadMiddle, sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.EndScene
    
    
     
    ' Present the backbuffer contents to the front buffer (screen)
    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    'update fps
    fps = fps + 1
End Sub

Function vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    'vector creation helper function
    vec3.x = x
    vec3.y = y
    vec3.z = z
End Function

Private Sub tmrFps_Timer()
    'display fps
    Form1.lblFps.Caption = fps & " frames per second."
    'reset fps
    fps = 0
End Sub
