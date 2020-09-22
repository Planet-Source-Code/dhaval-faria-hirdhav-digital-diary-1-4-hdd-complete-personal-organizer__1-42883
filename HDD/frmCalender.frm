VERSION 5.00
Begin VB.Form frmCalender 
   BorderStyle     =   0  'None
   Caption         =   "Hirdhav Digital Diary - Scheduler"
   ClientHeight    =   5775
   ClientLeft      =   -45
   ClientTop       =   0
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmCalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSchList 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   600
      TabIndex        =   61
      ToolTipText     =   "Double click on the item to get the full inforamtion."
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox txtDate 
      Height          =   330
      Left            =   1800
      TabIndex        =   60
      Top             =   240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   1440
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   36
      Left            =   1240
      Top             =   3780
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   36
      Left            =   1240
      Top             =   3560
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   35
      Left            =   640
      Top             =   3780
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   35
      Left            =   640
      Top             =   3560
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   34
      Left            =   4240
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   34
      Left            =   4240
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   33
      Left            =   3640
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   33
      Left            =   3640
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   32
      Left            =   3040
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   32
      Left            =   3040
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   31
      Left            =   2440
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   31
      Left            =   2440
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   30
      Left            =   1840
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   30
      Left            =   1840
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   29
      Left            =   1240
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   29
      Left            =   1240
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   28
      Left            =   640
      Top             =   3300
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   28
      Left            =   640
      Top             =   3080
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   27
      Left            =   4240
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   27
      Left            =   4240
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   26
      Left            =   3640
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   26
      Left            =   3640
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   25
      Left            =   3040
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   25
      Left            =   3040
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   24
      Left            =   2440
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   24
      Left            =   2440
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   23
      Left            =   1840
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   23
      Left            =   1840
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   22
      Left            =   1240
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   22
      Left            =   1240
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   21
      Left            =   640
      Top             =   2820
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   21
      Left            =   640
      Top             =   2600
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   20
      Left            =   4240
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   20
      Left            =   4240
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   19
      Left            =   3640
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   19
      Left            =   3640
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   18
      Left            =   3040
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   18
      Left            =   3040
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   17
      Left            =   2440
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   17
      Left            =   2440
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   16
      Left            =   1840
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   16
      Left            =   1840
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   15
      Left            =   1240
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   15
      Left            =   1240
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   14
      Left            =   640
      Top             =   2340
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   14
      Left            =   640
      Top             =   2120
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   13
      Left            =   4240
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   13
      Left            =   4240
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   12
      Left            =   3640
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   12
      Left            =   3640
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   11
      Left            =   3040
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   11
      Left            =   3040
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   10
      Left            =   2440
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   10
      Left            =   2440
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   9
      Left            =   1840
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   9
      Left            =   1840
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   8
      Left            =   1240
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   8
      Left            =   1240
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   7
      Left            =   640
      Top             =   1860
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   7
      Left            =   640
      Top             =   1640
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   6
      Left            =   4240
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   6
      Left            =   4240
      Top             =   1160
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   5
      Left            =   3640
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   5
      Left            =   3640
      Top             =   1160
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   4
      Left            =   3045
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   4
      Left            =   3045
      Top             =   1155
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   3
      Left            =   2440
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   3
      Left            =   2440
      Top             =   1160
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   2
      Left            =   1840
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   2
      Left            =   1840
      Top             =   1160
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   1
      Left            =   1240
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   1
      Left            =   1240
      Top             =   1160
      Width           =   60
   End
   Begin VB.Shape shapeAP2 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   0
      Left            =   640
      Top             =   1380
      Width           =   60
   End
   Begin VB.Shape shapeAP1 
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   0
      Left            =   640
      Top             =   1160
      Width           =   60
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click on the date to Add, Edit or Delete Record. Double click on date to Get Full Information."
      Height          =   495
      Left            =   720
      TabIndex        =   59
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   58
      Top             =   3555
      Width           =   660
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "December"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   57
      Top             =   3555
      Width           =   1785
   End
   Begin VB.Label lblMainMenuSupport 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      Height          =   225
      Left            =   120
      TabIndex        =   55
      Top             =   5340
      Width           =   4995
   End
   Begin VB.Shape shapeMainMenu 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Line Line28 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   54
      Top             =   4800
      Width           =   525
   End
   Begin VB.Line Line27 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Line Line26 
      BorderWidth     =   2
      X1              =   1800
      X2              =   4800
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   36
      Left            =   1200
      TabIndex        =   53
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   35
      Left            =   600
      TabIndex        =   52
      Top             =   3600
      Width           =   615
   End
   Begin VB.Line Line25 
      BorderWidth     =   2
      X1              =   600
      X2              =   1800
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line24 
      BorderWidth     =   2
      X1              =   1200
      X2              =   1200
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Line Line23 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   34
      Left            =   4200
      TabIndex        =   51
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   33
      Left            =   3600
      TabIndex        =   50
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   32
      Left            =   3000
      TabIndex        =   49
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   31
      Left            =   2400
      TabIndex        =   48
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   30
      Left            =   1800
      TabIndex        =   47
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   29
      Left            =   1200
      TabIndex        =   46
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   28
      Left            =   600
      TabIndex        =   45
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   27
      Left            =   4200
      TabIndex        =   44
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   26
      Left            =   3600
      TabIndex        =   43
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   25
      Left            =   3000
      TabIndex        =   42
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   24
      Left            =   2400
      TabIndex        =   41
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   23
      Left            =   1800
      TabIndex        =   40
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   22
      Left            =   1200
      TabIndex        =   39
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   21
      Left            =   600
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   20
      Left            =   4200
      TabIndex        =   37
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   19
      Left            =   3600
      TabIndex        =   36
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   18
      Left            =   3000
      TabIndex        =   35
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   17
      Left            =   2400
      TabIndex        =   34
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   16
      Left            =   1800
      TabIndex        =   33
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   14
      Left            =   600
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   13
      Left            =   4200
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   29
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   28
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   27
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   26
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   24
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   21
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   6000
   End
   Begin VB.Label lblFri 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fri"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblSat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sat"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblThu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thu"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblTue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tue"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblWed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wed"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblMon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mon"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   800
      Width           =   615
   End
   Begin VB.Label lblSun 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sun"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   800
      Width           =   615
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   3600
      X2              =   3600
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Label lblNMSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "Next Month"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblNM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   330
      Width           =   255
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape shapeNM 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4080
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3960
      X2              =   4080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblCMSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Present Month"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3590
      TabIndex        =   6
      Top             =   330
      Width           =   255
   End
   Begin VB.Shape shapeCM 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   3480
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   3360
      X2              =   3480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblPMSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Previous Month"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblPM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   330
      Width           =   255
   End
   Begin VB.Shape shapePM 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2880
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1200
      X2              =   1200
      Y1              =   720
      Y2              =   3480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   3480
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   600
      X2              =   4800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   600
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1680
      X2              =   2880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblNowDateSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblNowDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now Date"
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   375
      Width           =   1710
   End
   Begin VB.Shape shapeNowDate 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCaptionSupport 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   240
      Y2              =   6000
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hirdhav Digital Diary  -  Scheduler"
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   10
      Width           =   2805
   End
   Begin VB.Shape shapeCaption 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   255
      Left            =   15
      Top             =   15
      Width           =   5160
   End
   Begin VB.Shape shapeBack 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   600
      Top             =   480
      Width           =   4215
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   600
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   1200
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   1800
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   2400
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   3000
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3600
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   4200
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   13
      Left            =   4200
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   12
      Left            =   3600
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   11
      Left            =   3000
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   10
      Left            =   2400
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   9
      Left            =   1800
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   1200
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   600
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   20
      Left            =   4200
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   19
      Left            =   3600
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   18
      Left            =   3000
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   17
      Left            =   2400
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   16
      Left            =   1800
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   15
      Left            =   1200
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   14
      Left            =   600
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   27
      Left            =   4200
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   26
      Left            =   3600
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   25
      Left            =   3000
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   24
      Left            =   2400
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   23
      Left            =   1800
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   22
      Left            =   1200
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   21
      Left            =   600
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   34
      Left            =   4200
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   33
      Left            =   3600
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   32
      Left            =   3000
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   31
      Left            =   2400
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   30
      Left            =   1800
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   29
      Left            =   1200
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   28
      Left            =   600
      Top             =   3000
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   36
      Left            =   1200
      Top             =   3480
      Width           =   615
   End
   Begin VB.Shape shapeDate 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   35
      Left            =   600
      Top             =   3480
      Width           =   615
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SelDate
Public strUsername
Public AMPM1
Public AMPM2
Public SelDateTmp
Public SelYearTmp

Private Sub Form_Load()
For I = 0 To 36
    shapeDate(I).BackColor = RGB(145, 155, 100)
Next I
For j = 0 To 36
    shapeAP1(j).BackColor = vbBlack
    shapeAP2(j).BackColor = vbBlack
    shapeAP1(j).Visible = False
    shapeAP2(j).Visible = False
Next j
SelDate = Day(Now)
Me.BackColor = RGB(145, 155, 100)
lstSchList.BackColor = RGB(145, 155, 100)
shapeCaption.BackColor = vbBlack
lblCaption.ForeColor = RGB(145, 155, 100)
shapeBack.BackColor = vbBlack
shapeNowDate.BackColor = RGB(145, 155, 100)
lblNowDate.Caption = Format(Date, "D/MMM/YYYY(DDD)")
shapePM.BackColor = RGB(145, 155, 100)
shapeCM.BackColor = RGB(145, 155, 100)
shapeNM.BackColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = RGB(145, 155, 100)

strUsername = frmMain.lblUsername.Caption
CurrentMonth = Format(Date, "MMM")
CurrentYear = Year(Now)
IDs
lblYear.Caption = Year(Now)
lblMonth.Caption = Format(Date, "MMMM")

For j = 0 To 36
    If lblDate(j).Caption = Day(Now) Then
        shapeDate(j).BackColor = vbBlack
        lblDate(j).ForeColor = RGB(145, 155, 100)
    End If
Next j

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan
        txtDate.Text = ReS("Date")
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        End If
        If Len(txtDate.Text) = 10 Then
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
Do
    txtDate.Text = ReS("Date")
    If Len(txtDate.Text) = 11 Then
        txtDate.SelStart = 0
        txtDate.SelLength = 2
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 7
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    Else
        txtDate.SelStart = 0
        txtDate.SelLength = 1
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 6
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    End If
    AMPM1 = ReS("AP1")
    AMPM2 = ReS("AP2")
    
    If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
        lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    End If
    
    For I = 0 To 36
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        Else
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
    If txtDate.SelText = CurrentYear Then
        If lblDate(I).Caption = SelDateTmp Then
            If AMPM1 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM1 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
            If AMPM2 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM2 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
        End If
    End If
    Next I
    ReS.MoveNext
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing


ErrHan:
If Err.Number = 3021 Then
    Exit Sub
End If
End Sub

Private Sub lblCaptionSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub lblCMSupport_Click()
For I = 0 To 36
    shapeAP1(I).Visible = False
    shapeAP2(I).Visible = False
Next I
lstSchList.Clear
CurrentMonth = Format(Date, "MMM")
lblMonth.Caption = Format(Date, "MMMM")
CurrentYear = Year(Now)
lblYear.Caption = CurrentYear
IDs
For I = 0 To 36
    If lblDate(I).Caption = Day(Now) Then
        shapeDate(I).BackColor = vbBlack
        lblDate(I).ForeColor = RGB(145, 155, 100)
    End If
Next I
SelDate = Day(Now)

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan
        txtDate.Text = ReS("Date")
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        Else
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
        If txtDate.SelText <> CurrentYear Then
            ReS.Close
            db.Close
            Set ReS = Nothing
            Set db = Nothing
            Exit Sub
        End If
Do
    txtDate.Text = ReS("Date")
    If Len(txtDate.Text) = 11 Then
        txtDate.SelStart = 0
        txtDate.SelLength = 2
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 7
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    Else
        txtDate.SelStart = 0
        txtDate.SelLength = 1
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 6
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    End If
    
    If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
        lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    End If
    
    For I = 0 To 36
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        Else
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
    If txtDate.SelText = CurrentYear Then
        If lblDate(I).Caption = SelDateTmp Then
            If AMPM1 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM1 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
            If AMPM2 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM2 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
        End If
    End If
    Next I
    ReS.MoveNext
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

ErrHan:
If Err.Number = 3021 Then
    Exit Sub
End If
End Sub

Private Sub lblCMSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCM.ForeColor = RGB(145, 155, 100)
shapeCM.BackColor = vbBlack
End Sub

Private Sub lblCMSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeCM.BackColor = RGB(145, 155, 100)
lblCM.ForeColor = vbBlack
End Sub

Private Sub lblDate_Click(Index As Integer)
If lblDate(Index).Caption = "" Then
    Exit Sub
End If
SelDate = lblDate(Index).Caption
For I = 0 To 36
    lblDate(I).ForeColor = vbBlack
Next I
If CurrentMonth & CurrentYear = Format(Date, "MMM") & Year(Now) Then
For j = 0 To 36
    If lblDate(j).Caption = Day(Now) Then
        lblDate(j).ForeColor = RGB(145, 155, 100)
    End If
Next j
End If

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan
lstSchList.Clear
Do

txtDate.Text = ReS("Date")

If Len(txtDate.Text) = 11 Then
    txtDate.SelStart = 0
    txtDate.SelLength = 2
    SelDateTmp = txtDate.SelText
    txtDate.SelStart = 7
    txtDate.SelLength = 4
    SelYearTmp = txtDate.SelText
End If
If Len(txtDate.Text) = 10 Then
    txtDate.SelStart = 0
    txtDate.SelLength = 1
    SelDateTmp = txtDate.SelText
    txtDate.SelStart = 6
    txtDate.SelLength = 4
    SelYearTmp = txtDate.SelText
End If
    
If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
    lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    ReS.MoveNext
Else
    ReS.MoveNext
End If
Loop

ErrHan:
    If Err.Number = 3021 Then
        Exit Sub
    End If
End Sub

Private Sub lblDate_DblClick(Index As Integer)
SchDate = lblDate(Index).Caption
SchMonth = CurrentMonth
SchYear = CurrentYear
frmFullSch.Show
Me.Hide
End Sub

Private Sub lblDate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    SchMonth = CurrentMonth
    SchDate = lblDate(Index).Caption
    SchYear = CurrentYear
    PopupMenu frmMenu.mnuDates
End If
End Sub

Private Sub lblMainMenuSupport_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lblMainMenuSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMainMenu.ForeColor = RGB(145, 155, 100)
shapeMainMenu.BackColor = vbBlack
End Sub

Private Sub lblMainMenuSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeMainMenu.BackColor = RGB(145, 155, 100)
lblMainMenu.ForeColor = vbBlack
End Sub

Private Sub lblNMSupport_Click()
lstSchList.Clear
For j = 0 To 36
    shapeAP1(j).Visible = False
    shapeAP2(j).Visible = False
Next j
For I = 0 To 36
    shapeDate(I).BackColor = RGB(145, 155, 100)
    lblDate(I).ForeColor = vbBlack
Next I
If CurrentMonth = "Jan" Then
    CurrentMonth = "Feb"
    lblMonth.Caption = "February"
ElseIf CurrentMonth = "Feb" Then
    CurrentMonth = "Mar"
    lblMonth.Caption = "March"
ElseIf CurrentMonth = "Mar" Then
    CurrentMonth = "Apr"
    lblMonth.Caption = "April"
ElseIf CurrentMonth = "Apr" Then
    CurrentMonth = "May"
    lblMonth.Caption = "May"
ElseIf CurrentMonth = "May" Then
    CurrentMonth = "Jun"
    lblMonth.Caption = "June"
ElseIf CurrentMonth = "Jun" Then
    CurrentMonth = "Jul"
    lblMonth.Caption = "July"
ElseIf CurrentMonth = "Jul" Then
    CurrentMonth = "Aug"
    lblMonth.Caption = "August"
ElseIf CurrentMonth = "Aug" Then
    CurrentMonth = "Sep"
    lblMonth.Caption = "September"
ElseIf CurrentMonth = "Sep" Then
    CurrentMonth = "Oct"
    lblMonth.Caption = "October"
ElseIf CurrentMonth = "Oct" Then
    CurrentMonth = "Nov"
    lblMonth.Caption = "November"
ElseIf CurrentMonth = "Nov" Then
    CurrentMonth = "Dec"
    lblMonth.Caption = "December"
ElseIf CurrentMonth = "Dec" Then
    CurrentMonth = "Jan"
    lblMonth.Caption = "January"
    CurrentYear = CurrentYear + 1
    lblYear.Caption = CurrentYear
End If
IDs
If CurrentMonth = Format(Date, "MMM") Then
    If CurrentYear = Year(Now) Then
        For I = 0 To 36
            If lblDate(I).Caption = Day(Now) Then
                shapeDate(I).BackColor = vbBlack
                lblDate(I).ForeColor = RGB(145, 155, 100)
            End If
        Next I
    End If
End If

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan

Do
    txtDate.Text = ReS("Date")
    If Len(txtDate.Text) = 11 Then
        txtDate.SelStart = 0
        txtDate.SelLength = 2
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 7
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    Else
        txtDate.SelStart = 0
        txtDate.SelLength = 1
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 6
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    End If
    AMPM1 = ReS("AP1")
    AMPM2 = ReS("AP2")
    
    If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
        lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    End If
    
    For I = 0 To 36
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        Else
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
    If txtDate.SelText = CurrentYear Then
        If lblDate(I).Caption = SelDateTmp Then
            If AMPM1 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM1 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
            If AMPM2 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM2 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
        End If
    End If
    Next I
    ReS.MoveNext
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

ErrHan:
If Err.Number = 3021 Then
    Exit Sub
End If
End Sub

Private Sub lblNMSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNM.ForeColor = RGB(145, 155, 100)
shapeNM.BackColor = vbBlack
End Sub

Private Sub lblNMSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNM.BackColor = RGB(145, 155, 100)
lblNM.ForeColor = vbBlack
End Sub

Private Sub lblNowDateSupport_Click()
frmJumpDate.Show vbModal
End Sub

Private Sub lblNowDateSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNowDate.ForeColor = RGB(145, 155, 100)
shapeNowDate.BackColor = vbBlack
End Sub

Private Sub lblNowDateSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapeNowDate.BackColor = RGB(145, 155, 100)
lblNowDate.ForeColor = vbBlack
End Sub

Private Sub lblPMSupport_Click()
lstSchList.Clear
For j = 0 To 36
    shapeAP1(j).Visible = False
    shapeAP2(j).Visible = False
Next j
For I = 0 To 36
    shapeDate(I).BackColor = RGB(145, 155, 100)
    lblDate(I).ForeColor = vbBlack
Next I
If CurrentMonth = "Jan" Then
    CurrentMonth = "Dec"
    lblMonth.Caption = "December"
    CurrentYear = CurrentYear - 1
    lblYear.Caption = CurrentYear
ElseIf CurrentMonth = "Feb" Then
    CurrentMonth = "Jan"
    lblMonth.Caption = "January"
ElseIf CurrentMonth = "Mar" Then
    CurrentMonth = "Feb"
    lblMonth.Caption = "February"
ElseIf CurrentMonth = "Apr" Then
    CurrentMonth = "Mar"
    lblMonth.Caption = "March"
ElseIf CurrentMonth = "May" Then
    CurrentMonth = "Apr"
    lblMonth.Caption = "April"
ElseIf CurrentMonth = "Jun" Then
    CurrentMonth = "May"
    lblMonth.Caption = "May"
ElseIf CurrentMonth = "Jul" Then
    CurrentMonth = "Jun"
    lblMonth.Caption = "June"
ElseIf CurrentMonth = "Aug" Then
    CurrentMonth = "Jul"
    lblMonth.Caption = "July"
ElseIf CurrentMonth = "Sep" Then
    CurrentMonth = "Aug"
    lblMonth.Caption = "August"
ElseIf CurrentMonth = "Oct" Then
    CurrentMonth = "Sep"
    lblMonth.Caption = "September"
ElseIf CurrentMonth = "Nov" Then
    CurrentMonth = "Oct"
    lblMonth.Caption = "October"
ElseIf CurrentMonth = "Dec" Then
    CurrentMonth = "Nov"
    lblMonth.Caption = "November"
End If
IDs
If CurrentMonth = Format(Date, "MMM") Then
    If CurrentYear = Year(Now) Then
        For I = 0 To 36
            If lblDate(I).Caption = Day(Now) Then
                shapeDate(I).BackColor = vbBlack
                lblDate(I).ForeColor = RGB(145, 155, 100)
            End If
        Next I
    End If
End If

Dim db As Database
Dim ReS As Recordset

Set db = OpenDatabase(App.Path + "\Data\" + strUsername + "\Sch.dat")
Set ReS = db.OpenRecordset(CurrentMonth)

On Error GoTo ErrHan
Do
    txtDate.Text = ReS("Date")
    If Len(txtDate.Text) = 11 Then
        txtDate.SelStart = 0
        txtDate.SelLength = 2
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 7
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    Else
        txtDate.SelStart = 0
        txtDate.SelLength = 1
        SelDateTmp = txtDate.SelText
        txtDate.SelStart = 6
        txtDate.SelLength = 4
        SelYearTmp = txtDate.SelText
    End If
    AMPM1 = ReS("AP1")
    AMPM2 = ReS("AP2")
    
    If SelDateTmp & SelYearTmp = SelDate & CurrentYear Then
        lstSchList.AddItem ReS("TF") & ReS("AP1") & "  " & ReS("Description")
    End If
    
    For I = 0 To 36
        If Len(txtDate.Text) = 11 Then
            txtDate.SelStart = 7
            txtDate.SelLength = 4
        Else
            txtDate.SelStart = 6
            txtDate.SelLength = 4
        End If
    If txtDate.SelText = CurrentYear Then
        If lblDate(I).Caption = SelDateTmp Then
            If AMPM1 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM1 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
            If AMPM2 = "AM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP1(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP1(I).Visible = True
            End If
            If AMPM2 = "PM" Then
                If shapeDate(I).BackColor = vbBlack Then
                    shapeAP2(I).BackColor = RGB(145, 155, 100)
                End If
                shapeAP2(I).Visible = True
            End If
        End If
    End If
    Next I
    ReS.MoveNext
Loop

ReS.Close
db.Close

Set ReS = Nothing
Set db = Nothing

ErrHan:
If Err.Number = 3021 Then
    Exit Sub
End If
End Sub

Private Sub lblPMSupport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPM.ForeColor = RGB(145, 155, 100)
shapePM.BackColor = vbBlack
End Sub

Private Sub lblPMSupport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shapePM.BackColor = RGB(145, 155, 100)
lblPM.ForeColor = vbBlack
End Sub

Private Sub lstSchList_DblClick()
frmDetailSch.Show
Me.Hide
End Sub

Private Sub lstSchList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu frmMenu.mnuLB
End If
End Sub

Private Sub Timer1_Timer()
For I = 0 To 36
    If lblDate(I).Caption = SelDate Then
        If lblDate(I).ForeColor = RGB(145, 155, 100) Then
            lblDate(I).ForeColor = vbBlack
            Exit Sub
        End If
        If lblDate(I).ForeColor = vbBlack Then
            lblDate(I).ForeColor = RGB(145, 155, 100)
            Exit Sub
        End If
    End If
Debug.Print lblDate(I).Caption
Next I
End Sub
