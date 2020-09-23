VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form frmMain 
   Caption         =   "FlightSim - by Filip"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   96
      TabIndex        =   3
      Top             =   1440
      Width           =   1500
   End
   Begin RMControl7.RMCanvas drmMain 
      Height          =   2655
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Ground
Dim Ground As Direct3DRMMeshBuilder3
Dim GroundFrame As Direct3DRMFrame3
' Airport Boxes
Dim Bighangar As Direct3DRMMeshBuilder3
Dim Runway As Direct3DRMMeshBuilder3
Dim TaxiRoad1 As Direct3DRMMeshBuilder3
Dim TaxiRoad2 As Direct3DRMMeshBuilder3
Dim Airport As Direct3DRMMeshBuilder3
Dim Concrete As Direct3DRMMeshBuilder3
Dim Stripe As Direct3DRMMeshBuilder3
Dim Smallbox As Direct3DRMMeshBuilder3
Dim Lights As Direct3DRMMeshBuilder3
' Airport frames
Dim StripeFrame(0 To 4, 0 To 20) As Direct3DRMFrame3
Dim AirportFrame As Direct3DRMFrame3
Dim ConcreteFrame As Direct3DRMFrame3
Dim Taxi1Frame As Direct3DRMFrame3
Dim Taxi2Frame As Direct3DRMFrame3
Dim RunwayFrame As Direct3DRMFrame3
Dim BigHangarFrame As Direct3DRMFrame3
Dim CameraBoxFrame As Direct3DRMFrame3
Dim LightsFrame(0 To 10) As Direct3DRMFrame3
' Lights
Dim Light As Direct3DRMLight
Dim LightFrame As Direct3DRMFrame3
Dim Position As D3DVECTOR
Dim Vector As D3DVECTOR
Dim rotvalue As Single
Dim Z As Integer
Dim t As Double
Dim rotate As Single
Dim yrotate As Single
Dim zrotate As Single
Dim Thrust
Dim altit As Single

Private Sub drmMain_KeyDown(keyCode As Integer, Shift As Integer)
    Select Case keyCode
        Case vbKeyLeft
            zrotate = 0.02
            rotate = -1
        Case vbKeyRight
            zrotate = -0.02
            rotate = 1
        Case vbKeyDown
            altit = -3.14159265358797 / 4
        Case vbKeyUp
            altit = 3.14159265358797 / 4
        Case vbKeyHome
            Thrust = Round(Thrust, 3) + 0.003
        Case vbKeyEnd
            If Thrust > 0 Then
            Thrust = Round(Thrust, 3) - 0.003
            End If
    End Select
End Sub

Private Sub drmMain_KeyUp(keyCode As Integer, Shift As Integer)
    zrotate = 0
    rotate = 0
    altit = 0
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    Dim k As Single
    Dim l As Single
    rotate = 0
    MsgBox "Use <Home> to give more thrust, <end> to give less, and the arrow keys to move the plane", , "FlightSim"
    drmMain.StartWindowed
    Show
    Z = 0
    drmMain.Viewport.SetBack 9000
    drmMain.SceneFrame.SetSceneBackgroundRGB 0, 1, 1
    ' Set Frames
    For i = 0 To 4
        For j = 0 To 20
            Set StripeFrame(i, j) = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
        Next j
    Next i
    Set AirportFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set ConcreteFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set Taxi1Frame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set Taxi2Frame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set RunwayFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set CameraBoxFrame = drmMain.D3DRM.CreateFrame(drmMain.CameraFrame)
    Set BigHangarFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set GroundFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Set LightFrame = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    For i = 0 To 10
        Set LightsFrame(i) = drmMain.D3DRM.CreateFrame(drmMain.SceneFrame)
    Next i
    Set Ground = drmMain.CreateBoxMesh(5000, 0.01, 5000)
    Set Light = drmMain.D3DRM.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
    Ground.GetFace(4).SetColorRGB 0.5, 0.9, 0.5
    Ground.GetFace(5).SetColorRGB 0.5, 0.9, 0.5
    Light.SetRange 900
    ' Create Airport boxes
    Set Bighangar = drmMain.CreateBoxMesh(10, 5, 15)
    Set Runway = drmMain.CreateBoxMesh(5, 0.05, 190)
    Set Stripe = drmMain.CreateBoxMesh(0.4, 0.05, 5)
    Set TaxiRoad1 = drmMain.CreateBoxMesh(30, 0.05, 5)
    Set TaxiRoad2 = drmMain.CreateBoxMesh(5, 0.05, 20)
    Set Concrete = drmMain.CreateBoxMesh(35, 0.05, 35)
    Set Airport = drmMain.CreateBoxMesh(50, 0.025, 200)
    Set Smallbox = drmMain.CreateBoxMesh(0.1, 0.1, 0.1)
    Set Lights = drmMain.CreateBoxMesh(2, 0.5, 1)
    Airport.GetFace(4).SetColorRGB 0.25, 0.25, 0.25
    Airport.GetFace(5).SetColorRGB 0.25, 0.25, 0.25
    Runway.GetFace(4).SetColorRGB 0.5, 0.5, 0.5
    Runway.GetFace(5).SetColorRGB 0.5, 0.5, 0.5
    Lights.GetFace(0).SetColorRGB 1, 1, 0
    Lights.GetFace(1).SetColorRGB 1, 1, 0
    TaxiRoad1.GetFace(4).SetColorRGB 0.5, 0.5, 0.5
    TaxiRoad1.GetFace(5).SetColorRGB 0.5, 0.5, 0.5
    TaxiRoad2.GetFace(4).SetColorRGB 0.5, 0.5, 0.5
    TaxiRoad2.GetFace(5).SetColorRGB 0.5, 0.5, 0.5
    ' Add meshes to frames
    RunwayFrame.AddVisual Runway
    RunwayFrame.SetPosition Nothing, -17.5, -59.975, 0
    CameraBoxFrame.AddVisual Smallbox
    CameraBoxFrame.SetPosition Nothing, 0, 0, 0.25
    AirportFrame.AddVisual Airport
    AirportFrame.SetPosition Nothing, 0, -60, 0
    Taxi1Frame.AddVisual TaxiRoad1
    Taxi1Frame.SetPosition Nothing, 0, -60, -42.5
    Taxi2Frame.AddVisual TaxiRoad2
    Taxi2Frame.SetPosition Nothing, 17.5, -60, -35
    ConcreteFrame.AddVisual Concrete
    ConcreteFrame.SetPosition Nothing, 7.5, -60, -7.5
    BigHangarFrame.AddVisual Bighangar
    BigHangarFrame.SetPosition Nothing, 15, -57.5, -12.5
    k = -92.5
    l = -19.2
    i = 0
    j = 0
    Do
        Do
            DoEvents
            StripeFrame(i, j).AddVisual Stripe
            StripeFrame(i, j).SetPosition Nothing, l, -59.95, k
            i = i + 1
            l = l + 0.8
            If i = 5 Then Exit Do
        Loop
        i = 0
        l = -19.2
        j = j + 1
        k = k + 10
        If j = 20 Then Exit Do
    Loop
    For i = 0 To 10
        LightsFrame(i).AddVisual Lights
        LightsFrame(i).SetPosition Nothing, -17.5, -58, 180 + i * 40
    Next i
    ' Ground
    GroundFrame.AddVisual Ground
    GroundFrame.SetPosition Nothing, 0, -61, 25
    ' Lights
    LightFrame.AddLight Light
    LightFrame.SetPosition Nothing, 1000, 1000, 1000
    drmMain.CameraFrame.SetPosition drmMain.CameraFrame, -17.5, -59.5, -80
    Vector.Z = 1
    ' Here starts game cycle
    Do
        t = t + 0.01
        DoEvents
        drmMain.CameraFrame.GetPosition Nothing, Position
        drmMain.CameraFrame.GetRotation Nothing, Vector, rotvalue
        Label4 = "Thrust:" & Thrust
        yrotate = Round(yrotate, 3) + 0.001 * rotate
        If Thrust < 0.75 Then
        If Position.y > -59.5 Then
            drmMain.CameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, -0.25 * (-Thrust + 0.75) * ((Position.y + 59.6) / 4), 0
        End If
        End If
        If Position.y < -59.5 Then
            drmMain.CameraFrame.SetPosition Nothing, Position.x, -59.5, Position.Z
        End If
        If Position.y > -59.4 Then
            drmMain.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, zrotate
            drmMain.CameraFrame.SetRotation Nothing, 0, 1, 0, yrotate
        ElseIf Position.y <= -59.4 Then
            drmMain.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, yrotate
        End If
        drmMain.CameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, Thrust
        drmMain.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, altit / 100
        drmMain.Update
        Me.Caption = drmMain.FPS & " fps"
        Label2 = "Alt:" & Round(Position.y + 59.5, 2)
        Label3 = "Z:" & Position.Z
        picPos.Cls
        picPos.Line (49, 46)-(50, 54), &H888888, BF
        picPos.PSet (Position.x / 50 + 50, Position.Z / 50 + 50), vbBlack
    Loop
End Sub

Private Sub Form_Resize()
    drmMain.Width = Me.ScaleWidth - drmMain.Left
    drmMain.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
End Sub
