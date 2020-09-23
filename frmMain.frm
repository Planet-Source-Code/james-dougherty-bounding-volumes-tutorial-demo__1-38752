VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bounding Volumes Tutorial"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   10680
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bounding Sphere"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2630
         TabIndex        =   31
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBSRadius 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtBSCenter 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdComputeBS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Compute  Object 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtBSCenter 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   38
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Center      X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   1440
            TabIndex        =   36
            Top             =   720
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bounding Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2655
         Begin VB.CommandButton cmdComputeBB 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Compute Object 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtBBMin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBBMax 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtBBMin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBBMax 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min            X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max           X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   1440
            TabIndex        =   28
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   1440
            TabIndex        =   27
            Top             =   1320
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bounding Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   5400
         TabIndex        =   11
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBBMax 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtBBMin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBBMax 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtBBMin 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdComputeBB 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Compute Object 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   210
            Left            =   1440
            TabIndex        =   20
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   210
            Left            =   1440
            TabIndex        =   19
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max           X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   210
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min            X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   210
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   945
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bounding Sphere"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   8025
         TabIndex        =   3
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBSCenter 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdComputeBS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Compute  Object 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtBSCenter 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBSRadius 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                   Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Center      X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   555
         End
      End
   End
   Begin VB.PictureBox Canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   709
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3900
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '<------A Must Have - AND DEFINE EVERY VARIABLE :)

'Please Read...
'----------------------------------------------------------------
'I hope this helps someone. I figured theres a lot of collision
'detection tutorials, but what good are they if you don't
'understand the bounding volumes? So maybe this will help someone
'take the next leap.
'
'Hope you enjoy!
'-James-
'----------------------------------------------------------------

'I didn't use DirectX because I want to show actually
'what happening here and not be to confusing. I hate commenting
'this much because I think it messes up the code. To cluttered!
'BUT, I wanted everyone to understand whats going on and not get
'lost. Shall we begin!

'This will be the type definition to represent a point on  the screen
Private Type Point2D
 X As Single
 Y As Single
End Type

'Basic simple object

'Vertices (The white dots)
'What are vertices?
' -Plural for vertex. A defined point or
'  defined points in space. (my definition at least?)

'Sphere_Center (The center of the object)
'Sphere_Radius (The radius of the object)
'Box_Min (Min X(Left) and the Min Y(Top) of the object)
'Box_Max (Max X(Left) and the Max Y(Top) of the object)
Private Type Object2D
 Vertices(3) As Point2D
 Sphere_Center As Point2D
 Sphere_Radius As Single
 Sphere_Computed As Boolean
 Box_Min As Point2D
 Box_Max As Point2D
End Type

'I'll use two objects for this tutorial
Private Object(1) As Object2D

Private Sub Form_Load()
 
 'Here I just defined where I want the points to be located
 
 'V0 X,Y           V1 X,Y
 '   50,56            200,56
 '
 '   ¤----------------¤
 '   |                |
 '   |                |
 '   |                |
 '   |                |
 '   |                |
 '   |                |
 '   |                |
 '   ¤----------------¤
 '
 'V2 X,Y           V3 X,Y
 '   50,200           200,200
 
 With Object(0)
  .Vertices(0).X = 50:  .Vertices(0).Y = 56
  .Vertices(1).X = 200: .Vertices(1).Y = 56
  .Vertices(2).X = 50:  .Vertices(2).Y = 200
  .Vertices(3).X = 200: .Vertices(3).Y = 200
 End With
 
 'Draw the vertices(Dots) of object 0 on the picture box
 DrawVertices 0
 
 'Same as object 0 but moved to the left and smaller
 With Object(1)
  .Vertices(0).X = 550:  .Vertices(0).Y = 75
  .Vertices(1).X = 670: .Vertices(1).Y = 75
  .Vertices(2).X = 550:  .Vertices(2).Y = 190
  .Vertices(3).X = 670: .Vertices(3).Y = 190
 End With
 
 'Draw the vertices(Dots) of object 1 on the picture box
 DrawVertices 1
  
End Sub

Private Sub cmdClear_Click()
 Dim i As Long
 
 'This just clears the picture box and all the text fields
 Canvas.Cls
 For i = 0 To 3
  txtBBMin(i).Text = ""
  txtBBMax(i).Text = ""
  txtBSCenter(i).Text = ""
 Next
 txtBSRadius(0).Text = ""
 txtBSRadius(1).Text = ""
 
 'reset the spheres computed flags
 Object(0).Sphere_Computed = False
 Object(1).Sphere_Computed = False
 
 'Then redraw the objects
 DrawVertices 0
 DrawVertices 1
 
End Sub

Private Sub cmdComputeBB_Click(Index As Integer)
 
 'Note - I hate using integers thats why theres a CLng()
 'See ComputeBoundingBox() for details
 ComputeBoundingBox CLng(Index), Object(Index).Box_Min, Object(Index).Box_Max
 
 'Which object did we compute?
 Select Case Index
  Case 0
   'After calling ComputeBoundingBox() it returned Box_Min and
   'Box_Max back to use. I just show the data here of what it returned
   txtBBMin(0).Text = Object(0).Box_Min.X
   txtBBMin(1).Text = Object(0).Box_Min.Y
   txtBBMax(0).Text = Object(0).Box_Max.X
   txtBBMax(1).Text = Object(0).Box_Max.Y
  Case 1
   'Same as Case 0 but with object 1
   txtBBMin(2).Text = Object(1).Box_Min.X
   txtBBMin(3).Text = Object(1).Box_Min.Y
   txtBBMax(2).Text = Object(1).Box_Max.X
   txtBBMax(3).Text = Object(1).Box_Max.Y
 End Select
 
 'This draws the bounding box so you can see what it computed
 ShowBoundingBox CLng(Index), Object(Index).Box_Min, Object(Index).Box_Max
 
End Sub

Private Sub cmdComputeBS_Click(Index As Integer)
 
 'Note - I hate using integers thats why theres a CLng()
 'See ComputeBoundingSphere() for details
 ComputeBoundingSphere CLng(Index), Object(Index).Sphere_Center, Object(Index).Sphere_Radius
 
 Select Case Index
  Case 0
   'After calling ComputeBoundingSphere() it returned the spheres center
   'and the spheres radius back to use.
   'I just show the data here of what it returned
   txtBSCenter(0).Text = Object(0).Sphere_Center.X
   txtBSCenter(1).Text = Object(0).Sphere_Center.Y
   txtBSRadius(0).Text = Object(0).Sphere_Radius
   Object(0).Sphere_Computed = True
  Case 1
   'Same as Case 0 but with object 1
   txtBSCenter(2).Text = Object(1).Sphere_Center.X
   txtBSCenter(3).Text = Object(1).Sphere_Center.Y
   txtBSRadius(1).Text = Object(1).Sphere_Radius
   Object(1).Sphere_Computed = True
 End Select
 
 'This draws the bounding sphere so you can see what it computed
 ShowBoundingSphere CLng(Index), Object(Index).Sphere_Center, Object(Index).Sphere_Radius
 
 'If both spheres were computer the show the distance
 'I put this in for the fun of it.
 ShowDistance
 
End Sub

Private Sub DrawVertices(ObjectID As Long)
 
 'I put the DrawWidth to 5 to make the dots big
 Canvas.DrawWidth = 5
 
 'These take the vetex positions and makes a dot where they are
 'located at
 With Object(ObjectID)
  Canvas.Line (.Vertices(0).X, .Vertices(0).Y)-(.Vertices(0).X, .Vertices(0).Y), , BF
  Canvas.Line (.Vertices(1).X, .Vertices(1).Y)-(.Vertices(1).X, .Vertices(1).Y), , BF
  Canvas.Line (.Vertices(2).X, .Vertices(2).Y)-(.Vertices(2).X, .Vertices(2).Y), , BF
  Canvas.Line (.Vertices(3).X, .Vertices(3).Y)-(.Vertices(3).X, .Vertices(3).Y), , BF
 End With
 
 'then resizr down the dots
 Canvas.DrawWidth = 2
 
End Sub

'This function finds the distance between two givin points
Private Function ComputeDistance(Point1 As Point2D, Point2 As Point2D) As Single
 
 
 '   ¤----------------¤                 ¤----------------¤
 '   |                |                 |                |
 '   |                |                 |                |
 '   |      P1        |                 |       P2       |
 '   |      ¤---------|-----------------|--------¤       |
 '   |                |                 |                |
 '   |                |                 |                |
 '   |                |                 |                |
 '   ¤----------------¤                 ¤----------------¤
 
 
 ComputeDistance = (Sqr(((Point1.X - Point2.X) * (Point1.X - Point2.X)) + _
                        ((Point1.Y - Point2.Y) * (Point1.Y - Point2.Y))))
End Function

'Same as above but instead of finding the distance from
'center-center it finds the distance from sphere edge-sphere edge
Private Function ComputeCollisionDistance(Point1 As Point2D, Point2 As Point2D) As Single
 Dim Distance As Single
 
 '   ¤----------------¤                 ¤----------------¤
 '   |                |                 |                |
 '   |                |                 |                |
 '   |              \ |P1             P2| /              |
 '   |   Sphere----> ||¤---------------¤|| <----Sphere   |
 '   |   Edge       / |                 | \     Edge     |
 '   |                |                 |                |
 '   |                |                 |                |
 '   ¤----------------¤                 ¤----------------¤
 
 Distance = ComputeDistance(Point1, Point2)
 Distance = (Distance - ((Object(0).Sphere_Radius) + (Object(1).Sphere_Radius)))
 ComputeCollisionDistance = Distance
 
End Function

Private Sub ShowDistance()
 Dim CollisionDistance As Single
 Dim Distance As Single
 Dim XValue As Point2D
 
 'Make sure both bounding spheres were computed
 If Not Object(0).Sphere_Computed Or Not Object(1).Sphere_Computed Then Exit Sub
 
 Canvas.DrawWidth = 2
 
 'Draw a line to the center of sphere 0 and the center of sphere 1
 Canvas.Line (Object(0).Sphere_Center.X, Object(0).Sphere_Center.Y)- _
             (Object(1).Sphere_Center.X, Object(1).Sphere_Center.Y), &HFFFFFF
 
 'Get the distance between the centers (see ComputeDistance())
 Distance = ComputeDistance(Object(0).Sphere_Center, Object(1).Sphere_Center)
 
 'Get the distance between the two edges (see ComputeCollisionDistance())
 CollisionDistance = ComputeCollisionDistance(Object(0).Sphere_Center, Object(1).Sphere_Center)
 
 'I used this to find the center of the distance line
 'its used to position the text
 XValue = Point2D_Average(Object(0).Sphere_Center, Object(1).Sphere_Center)
 
 'take off 50 because of the length of the text
 Canvas.CurrentX = XValue.X - 50
 
 'add 20 to put it right below the line
 Canvas.CurrentY = Object(1).Sphere_Center.Y - 20
 
 'Show the distance
 Canvas.Print "Distance - " & Distance
 
 'take off 50 because of the length of the text
 Canvas.CurrentX = XValue.X - 50
 Canvas.CurrentY = Object(1).Sphere_Center.Y
 
 'Show the collision distance
 Canvas.Print "Collision Distance - " & CollisionDistance
  
End Sub

Private Sub ComputeBoundingBox(ObjectID As Long, rBox_Min As Point2D, rBox_Max As Point2D)
 Dim Min As Point2D
 Dim Max As Point2D
 Dim i As Long
 
 'give the tmp Min and Max crazy value that you know will
 'be overwritten no matter what
 Min.X = 1E+35
 Min.Y = 1E+35
 Max.X = -1E+35
 Max.Y = -1E+35
 
 'Loop through all the vertices
 For i = 0 To 3
  With Object(ObjectID)
   
   'its not really a big deal to use this algorithm here, but
   'if you have 1000 vertices its awesome
   
   'simply check the extents
   'If vertex.x < min.x | min.x = vertex.x
   'If vertex.y > min.y | min.y = vertex.y
   'If vertex.x < max.x | max.x = vertex.x
   'If vertex.y < max.y | max.y = vertex.y
   
   If .Vertices(i).X < Min.X Then Min.X = .Vertices(i).X
   If .Vertices(i).Y < Min.Y Then Min.Y = .Vertices(i).Y
   If .Vertices(i).X > Max.X Then Max.X = .Vertices(i).X
   If .Vertices(i).Y > Max.Y Then Max.Y = .Vertices(i).Y
  End With
 Next
 
 'thats it! Now just return the values
 rBox_Min = Min
 rBox_Max = Max
 
End Sub

'This draws the bounding box of the object
Private Sub ShowBoundingBox(ObjectID As Long, Box_Min As Point2D, Box_Max As Point2D)
 Dim Color As Long
 
 'make the draw width a tid-bit bigger
 Canvas.DrawWidth = 3
 
 'Change the color depending on the object
 If ObjectID = 0 Then Color = &H8000&
 If ObjectID = 1 Then Color = &H808000
 
 'which bounding box do we draw
 With Object(ObjectID)
  
  'this draws the box
  Canvas.Line (.Box_Min.X, .Box_Min.Y)-(.Box_Max.X, .Box_Max.Y), Color, B
 
  Canvas.ForeColor = &HFFFFFF 'White
  'position the text
  Canvas.CurrentX = .Box_Min.X - 20
  Canvas.CurrentY = .Box_Min.Y - 15
  'and display Box Min
  Canvas.Print "Box Min"
  
  'reposition the text
  Canvas.CurrentX = .Box_Max.X - 20
  Canvas.CurrentY = .Box_Max.Y + 5
  'and display Box Max
  Canvas.Print "Box Max"
  
  'Redraw the vertices so the don't get drawn over.
  DrawVertices ObjectID
 End With
 
 'reset the draw width
 Canvas.DrawWidth = 2
 
End Sub

Private Sub ComputeBoundingSphere(ObjectID As Long, rSphere_Center As Point2D, rSphere_Radius As Single)
 Dim Min As Point2D
 Dim Max As Point2D
 
 'same as bounding box but we have to compute the sphere instead
 ComputeBoundingBox ObjectID, Min, Max
 
 'get the max point between the sphere and / 2
 rSphere_Radius = Point2D_Max(Max.X - Min.X, Max.Y - Min.Y) / 2
 
 'finds the center of the sphere by averaging the min and max points
 rSphere_Center = Point2D_Average(Max, Min)
 
End Sub

'This draws the bounding circle of the object
Private Sub ShowBoundingSphere(ObjectID As Long, Sphere_Center As Point2D, Sphere_Radius As Single)
 Dim Color As Long
 
 Canvas.DrawWidth = 3
 'Change the color depending on the object
 If ObjectID = 0 Then Color = &H80&    'Dark Red
 If ObjectID = 1 Then Color = &H40C0&  'Dark Orange
 
 With Object(ObjectID)
  'Change the color
  Canvas.ForeColor = Color
  
  'Draw the bounding circle based on the computed sphere center and radius
  Canvas.Circle (.Sphere_Center.X, .Sphere_Center.Y), .Sphere_Radius
  
  'Show the center of the bounding sphere
  '(same as the vertex drawing, but with the sphere's center)
  Canvas.DrawWidth = 5
  Canvas.Line (.Sphere_Center.X, .Sphere_Center.Y)-(.Sphere_Center.X, .Sphere_Center.Y), &HFFFF&
  
  'Change the color to yellow
  Canvas.ForeColor = &H8000000E
  'Display "Center" right below the sphere's center point
  Canvas.CurrentX = .Sphere_Center.X
  Canvas.CurrentY = .Sphere_Center.Y + 5
  Canvas.Print "Center"
  
  'Redraw the vertices so the don't get drawn over.
  DrawVertices ObjectID
 End With
 'Reset the draw width
 Canvas.DrawWidth = 2
 
End Sub

'utility functions to save typing :)
'DP = Destination Point
'P1 = Point1
'P2 = Point2

'DP = P1 + P2
Private Function Point2D_Add(Point1 As Point2D, Point2 As Point2D) As Point2D
 Point2D_Add.X = Point1.X + Point2.X
 Point2D_Add.Y = Point1.Y + Point2.Y
End Function

'Adds two point together the divides by two giving the center (or average of the two)
'
'DP = P1 + P2
'DP.x = DP.x / 2
'DP.y = DP.y / 2
Private Function Point2D_Average(Point1 As Point2D, Point2 As Point2D) As Point2D
 Point2D_Average = Point2D_Add(Point1, Point2)
 Point2D_Average.X = Point2D_Average.X / 2
 Point2D_Average.Y = Point2D_Average.Y / 2
End Function

'Finds which point has the greatest value

'P1 > P2 | DP = P1
'P2 > P1 | DP = P2
Private Function Point2D_Max(Point1 As Single, Point2 As Single) As Single
 If Point1 > Point2 Then Point2D_Max = Point1: Exit Function
 If Point2 > Point1 Then Point2D_Max = Point2: Exit Function
End Function
