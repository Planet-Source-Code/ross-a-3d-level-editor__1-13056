VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   6240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dx As New DirectX7
Dim dd As DirectDraw4
Dim clip As DirectDrawClipper
Dim d3drm As Direct3DRM3
Dim scene As Direct3DRMFrame3
Dim cam As Direct3DRMFrame3
Dim dev As Direct3DRMDevice3
Dim MYmyview As Direct3DRMViewport2
Dim mesh As Direct3DRMMeshBuilder3
Dim ddsp As DirectDrawSurface7
Dim ddsb As DirectDrawSurface7
Dim yes As Boolean
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub Form_Load()
'Set the Variables
Dim X1 As Single, Y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, x4 As Single, y4 As Single, z4 As Single, TexFile As String, TileX As Single, TileY As Single, r As Single, g As Single, b As Single
Dim counter, count2 As Integer
Dim f As Direct3DRMFace2
Dim t As Direct3DRMTexture3
Dim pos As D3DVECTOR
'Set 3d variables

    Set dd = dx.DirectDraw4Create("")
    Set clip = dd.CreateClipper(0)
    clip.SetHWnd Me.hWnd
    Set d3drm = dx.Direct3DRMCreate()
    Set scene = d3drm.CreateFrame(Nothing)
    Set cam = d3drm.CreateFrame(scene)
    scene.AddLight d3drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 240, 240, 0.5)
    Set dev = d3drm.CreateDeviceFromClipper(clip, "IID_IDirect3DHALDevice", Me.ScaleWidth, Me.ScaleHeight)
    dev.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    dev.SetDither D_TRUE
    Set MYmyview = d3drm.CreateViewport(dev, cam, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    MYmyview.SetBack 5000!
    Set mesh = d3drm.CreateMeshBuilder()
    mesh.SetPerspective D_TRUE
    scene.AddVisual mesh

'Build the scene from 2d map

For counter = 1 To 225
DoEvents
TexFile = tex(counter)


If ((TexFile = 9) Or (TexFile = 8)) Then
For count2 = 1 To 5
DoEvents
If count2 = 1 Then
X1 = X(counter)
z1 = z(counter)
Y1 = 0
x2 = X1 + 200
z2 = z1 + 200
y2 = 0
x3 = x2
z3 = z2
y3 = 300
x4 = X1
z4 = z1
y4 = 300
End If

If count2 = 2 Then
X1 = X(counter) + 200
z1 = z(counter) + 200
Y1 = 0
x2 = X(counter)
z2 = z(counter)
y2 = 0
x3 = x2
z3 = z2
y3 = 300
x4 = X1
z4 = z1
y4 = 300
End If

If count2 = 3 Then
X1 = X(counter)
z1 = z(counter) + 200
Y1 = 0
x2 = X(counter) + 200
z2 = z(counter)
y2 = 0
x3 = x2
z3 = z2
y3 = 300
x4 = X1
z4 = z1
y4 = 300
End If

If count2 = 4 Then
X1 = X(counter) + 200
z1 = z(counter)
Y1 = 0
x2 = X(counter)
z2 = z(counter) + 200
y2 = 0
x3 = x2
z3 = z2
y3 = 300
x4 = X1
z4 = z1
y4 = 300
End If

If count2 = 5 Then
TexFile = "0"
X1 = X(counter)
z1 = z(counter)
Y1 = 0
x2 = X(counter)
z2 = z(counter) + 200
y2 = 0
x3 = x2 + 200
z3 = z2
y3 = 0
x4 = X1
z4 = z1
y4 = 0
End If

    'create face
Set f = d3drm.CreateFace()
    ' add vertexs
f.AddVertex X1, Y1, z1
f.AddVertex x2, y2, z2
f.AddVertex x3, y3, z3
f.AddVertex x4, y4, z4
    ' get type of file
If TexFile = "" Then
    ' set colors
f.SetColorRGB r, g, b
Else
    ' create textuere
Set t = d3drm.LoadTexture(App.Path & "\" & TexFile & ".bmp")
    ' set u and v values
f.SetTextureCoordinates 3, 0, 0
f.SetTextureCoordinates 2, TileX, 0
f.SetTextureCoordinates 1, TileX, TileY
f.SetTextureCoordinates 0, 0, TileY
    ' set the texture
t.SetDecalTransparency D_TRUE
f.SetTexture t
End If



' add face to mesh
mesh.AddFace f

Next count2
TexFile = "0"
End If
    
X1 = X(counter)
z1 = z(counter)
x2 = X1
z2 = z1 + 200
x3 = x2 + 200
z3 = z2
x4 = x3
z4 = z1
Y1 = 0
y2 = 0
y3 = 0
y4 = 0
TileX = 1
TileY = 1
r = 0
g = 0
b = 0
    
    
    
    
    'create face
Set f = d3drm.CreateFace()
    ' add vertexs
f.AddVertex X1, Y1, z1
f.AddVertex x2, y2, z2
f.AddVertex x3, y3, z3
f.AddVertex x4, y4, z4
    ' get type of file
If TexFile = "" Then
    ' set colors
f.SetColorRGB r, g, b
Else
    ' create textuere
Set t = d3drm.LoadTexture(App.Path & "\" & TexFile & ".bmp")
    ' set u and v values
f.SetTextureCoordinates 0, 0, 0
f.SetTextureCoordinates 1, TileX, 0
f.SetTextureCoordinates 2, TileX, TileY
f.SetTextureCoordinates 3, 0, TileY
    ' set the texture
f.SetTexture t
End If
    ' add face to mesh
mesh.AddFace f


Next counter


' create sky around
Call MakeWall(d3drm, mesh, -500, -5, -500, 3500, -5, -500, 3500, 1000, -500, -500, 1000, -500, "sky", 20, 20, 0, 0, 0)
Call MakeWall(d3drm, mesh, 3500, -5, -500, 3500, -5, 3500, 3500, 1000, 3500, 3500, 1000, -500, "sky", 20, 20, 0, 0, 0)
Call MakeWall(d3drm, mesh, 3500, -5, 3500, -500, -5, 3500, -500, 1000, 3500, 3500, 1000, 3500, "sky", 20, 20, 0, 0, 0)
Call MakeWall(d3drm, mesh, -500, -5, 3500, -500, -5, -500, -500, 1000, -500, -500, 1000, 3500, "sky", 20, 20, 0, 0, 0)
Call MakeWall(d3drm, mesh, -500, 1000, -500, 3500, 1000, -500, 3500, 1000, 3500, -500, 1000, 3500, "sky", 20, 20, 0, 0, 0)
' create sea around
Call MakeWall(d3drm, mesh, -500, -5, -500, -500, -5, 3500, 3500, -5, 3500, 3500, -5, -500, "1", 1, 1, 0, 0, 0)

cam.GetPosition scene, pos ' GET THE CAMERA'S POSITION IN SCENE
pos.X = 1500
pos.z = 1500
pos.y = Module1.height
cam.SetPosition scene, pos.X, pos.y, pos.z ' SET CAMERAS NEW POSITION
       
    Me.Show
    Me.Refresh
    DoEvents

Do While DoEvents()
    'MOVE FORWRD!!!
If GetKeyState(vbKeyUp) < -1 Then cam.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, (Module1.speed / 10) ' 2 unit foward in z direction
    'MOVE BACK
If GetKeyState(vbKeyDown) < -1 Then cam.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, ((Module1.speed - (Module1.speed * 2)) / 10) ' 2 unit back in z direction
    'ROTRATE LEFT
If GetKeyState(vbKeyLeft) < -1 Then cam.AddRotation D3DRMCOMBINE_BEFORE, 0, -2, 0, (Module1.turning / 100) ' rotate 2 radian to the left
    'ROTRATE RIGHT
If GetKeyState(vbKeyRight) < -1 Then cam.AddRotation D3DRMCOMBINE_BEFORE, 0, 2, 0, (Module1.turning / 100) ' rotate 2 radian to the right
cam.GetPosition scene, pos ' GET THE CAMERA'S POSITION IN SCENE

If pos.X > 2985 Then pos.X = 2985
If pos.X < 15 Then pos.X = 15
If pos.z > 2985 Then pos.z = 2985
If pos.z < 15 Then pos.z = 15

cam.SetPosition scene, pos.X, pos.y, pos.z ' SET CAMERAS NEW POSITION
    ' render the scene
MYmyview.Clear D3DRMCLEAR_ALL
MYmyview.Render scene
dev.Update
    'IF WE PRESS ESCAPE THEN QUIT!!
If GetKeyState(vbKeyEscape) < -5 Then
yes = False
Timer1.Enabled = True
Exit Sub
End If
Loop
    
End Sub

Private Sub MakeWall(d3drm As Direct3DRM3, mesh As Direct3DRMMeshBuilder3, X1 As Single, Y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, x4 As Single, y4 As Single, z4 As Single, TexFile As String, TileX As Single, TileY As Single, r As Single, g As Single, b As Single)
    Dim f As Direct3DRMFace2
    Dim t As Direct3DRMTexture3
    Set f = d3drm.CreateFace()
    f.AddVertex X1, Y1, z1
    f.AddVertex x2, y2, z2
    f.AddVertex x3, y3, z3
    f.AddVertex x4, y4, z4
    If TexFile = "" Then
        f.SetColorRGB r, g, b
    Else
        Set t = d3drm.LoadTexture(App.Path & "\" & TexFile & ".bmp")
        f.SetTextureCoordinates 0, 0, 0
        f.SetTextureCoordinates 1, TileX, 0
        f.SetTextureCoordinates 2, TileX, TileY
        f.SetTextureCoordinates 3, 0, TileY
        f.SetTexture t
    End If
    mesh.AddFace f
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
