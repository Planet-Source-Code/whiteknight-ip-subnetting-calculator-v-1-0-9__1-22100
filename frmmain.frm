VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Information"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9180
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tipback 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   50
      Top             =   5520
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Label lbltip 
         BackColor       =   &H80000018&
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   4875
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   120
      TabIndex        =   48
      Top             =   3120
      Width           =   4575
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "IP Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton bttn_check 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compute &Now"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton bttn_defaultmask 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Default Mask"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton bttn_reset 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reset"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox picIP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         MousePointer    =   3  'I-Beam
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txtip 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   360
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "168"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtip 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   15
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "192"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtip 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   705
            MaxLength       =   3
            TabIndex        =   2
            Text            =   "100"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtip 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1050
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "0"
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   300
            TabIndex        =   37
            Top             =   0
            Width           =   60
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   645
            TabIndex        =   36
            Top             =   0
            Width           =   60
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   990
            TabIndex        =   35
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.PictureBox picSM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         MousePointer    =   3  'I-Beam
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
         Begin VB.TextBox txtsm 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1065
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "0"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtsm 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   720
            MaxLength       =   3
            TabIndex        =   6
            Text            =   "255"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtsm 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   15
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "255"
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox txtsm 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   360
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "255"
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1005
            TabIndex        =   33
            Top             =   0
            Width           =   60
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   645
            TabIndex        =   32
            Top             =   0
            Width           =   60
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   300
            TabIndex        =   31
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
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
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Subnet Mask:"
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
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblNetID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1320
         TabIndex        =   39
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Network ID:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Binary Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   4575
      Begin VB.Label lblbinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Mask:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblbinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1200
         TabIndex        =   25
         Top             =   600
         Width           =   3195
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Network ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblbinary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   3195
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subnetting Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   4800
      TabIndex        =   19
      Top             =   1320
      Width           =   4335
      Begin VB.ListBox lstNetIDs 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lblRange 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   47
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblnumhosts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblnumnetworks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Hosts:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Subnetworks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Network ID's                                         Broadcast ID's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   4065
      End
   End
   Begin VB.CommandButton bttn_close 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Close"
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
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Network Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   4800
      TabIndex        =   12
      Top             =   0
      Width           =   4335
      Begin VB.Label lblgood 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lbltype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblclass 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Good IP For Host:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP Address Class:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Const SHACF_DEFAULT  As Long = &H0

Const GWL_EXSTYLE = (-20)
' member variable for TipCaption property
Private m_TipCaption As String
' member variable for TipTop property
Private m_TipTop As Long
' member variable for TipLeft property
Private m_TipLeft As Long
' member variable for TipVisible property
Private m_TipVisible As Boolean
' The IsTipVisible property
Property Get IsTipVisible() As Boolean
  IsTipVisible = m_IsTipVisible
End Property
' The TipVisible property
Property Get TipVisible() As Boolean
  TipVisible = m_TipVisible
End Property
Property Let TipVisible(ByVal newValue As Boolean)
  m_TipVisible = newValue
  tipback.Visible = m_TipVisible
End Property
' The TipLeft property
Property Get TipLeft() As Long
  TipLeft = m_TipLeft
End Property
Property Let TipLeft(ByVal newValue As Long)
  m_TipLeft = newValue
  tipback.Left = newValue
End Property
' The TipTop property
Property Get TipTop() As Long
  TipTop = m_TipTop
End Property
Property Let TipTop(ByVal newValue As Long)
  m_TipTop = newValue
  tipback.Top = newValue
End Property
Private Sub Form_Initialize()
  m_TipCaption = ""
  m_IsTipVisible = False
  m_TipTop = 0
  m_TipLeft = 0
  m_TipVisible = False
End Sub



' The TipCaption property

Property Get TipCaption() As String
  TipCaption = m_TipCaption
End Property

Property Let TipCaption(ByVal newValue As String)
  m_TipCaption = newValue
  lbltip.Caption = "  " & m_TipCaption
End Property


Private Sub bttn_check_Click()
  
  'just incase an error does happen we will ignore it
  On Error Resume Next
  Dim x As Integer, tempsm As String, temprange As Integer, tempip As String
  Dim tindex As Integer
  'Hide Our Tip If its Visible
  If IsTipVisible = True Then HideTip
  Enable
  'hold our ip and mask for later use
  tempsm = txtsm(0).Text & "." & txtsm(1).Text & "." & txtsm(2).Text & "." & txtsm(3).Text
  tempip = txtip(0).Text & "." & txtip(1).Text & "." & txtip(2).Text & "." & txtip(3).Text
  lblbinary(0).Caption = ""
  lblbinary(1).Caption = ""
  'Check the mask incase it wasnt checked before
  For tindex = 0 To txtsm.Count - 1
    If checkmask(tindex) = False Then
      MsgBox "Number for mask must be:" & Chr(13) & "0, 128, 192, 224, 240, 248, 252, 254, or 255" & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
      txtsm(tindex).SetFocus
      'highlight the section
      SendKeys "{HOME}+{END}"
      Exit Sub
    End If
  Next tindex
  'Set the binary IP And Mask Labels
  For x = 0 To txtip.Count - 1
    lblbinary(0).Caption = lblbinary(0).Caption & "." & ConvertBin(CInt(txtip(x).Text))
    lblbinary(1).Caption = lblbinary(1).Caption & "." & ConvertBin(CInt(txtsm(x).Text))
  Next x
  lblbinary(0).Caption = Mid(lblbinary(0).Caption, 2)
  lblbinary(1).Caption = Mid(lblbinary(1).Caption, 2)
  'Set the binary Network ID Label
  lblbinary(2).Caption = GetBinNetID(lblbinary(0).Caption, lblbinary(1).Caption)
  'Set the Network ID by converting the Binary Network ID
  lblNetID.Caption = ConvertBinToIP(lblbinary(2).Caption)
  'Set the IP Class Label
  lblclass.Caption = GetIPClass(lblNetID.Caption)
  'Set the type of IP Label (Public, Private, Loopback, Multicast, or Experimental)
  lbltype.Caption = IPType
  'Get the range
  temprange = GetRange(tempsm)
  'If the range = 256 then the range = 1
  If temprange = 256 Then temprange = 1
  'Set the Range Label
  lblRange.Caption = temprange
  'Add the CIDR to the Network ID
  lblNetID.Caption = lblNetID.Caption & "/" & GetBits(tempsm)
  'Set the Total Number of networks label
  lblnumnetworks.Caption = GetPosNetworks(tempsm)
  'Set the total number of hosts allowed per network label
  lblnumhosts.Caption = GetPosHosts(lblbinary(1).Caption)
  'Load all Network ID's and Broadcast ID's to the list
  LoadNetID lblNetID.Caption, tempsm, iRange
  'Set If IP Can be assigned to a host
  lblgood.Caption = IsGoodIP(tempip)
  
  'Heighlight the network in the listbox
  HighlightNetworkID (Mid(lblNetID.Caption, 1, InStr(1, lblNetID.Caption, "/") - 1))
  'Return Focus to The IP Text Box
  txtip(0).SetFocus
End Sub

Private Sub bttn_check_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "Get the information from the supplied IP address && Mask.", bttn_check
    Exit Sub
  End If
End Sub

Private Sub bttn_close_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
  'we are done
  Unload Me
End Sub

Private Sub bttn_close_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  ShowTip "Close This Program.", bttn_close
End If
End Sub

Private Sub bttn_defaultmask_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
  'get the default mask for the IP
  defaultmask txtip(0).Text & "." & txtip(1).Text & "." & txtip(2).Text & "." & txtip(3).Text
  txtsm(0).SetFocus
End Sub

Private Sub bttn_defaultmask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "Sets the default Subnet Mask for the IP address.", bttn_defaultmask
    Exit Sub
End If
End Sub

Private Sub bttn_reset_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
  'This just sets the IP and Subnet Mask To a Default _
   and clear all other fields.
  txtip(0).Text = "192"
  txtip(1).Text = "168"
  txtip(2).Text = "100"
  txtip(3).Text = "1"
  txtsm(0).Text = "255"
  txtsm(1).Text = "255"
  txtsm(2).Text = "255"
  txtsm(3).Text = "0"
  lblbinary(0).Caption = ""
  lblbinary(1).Caption = ""
  lblbinary(2).Caption = ""
  lblclass.Caption = ""
  lblNetID.Caption = ""
  lblgood.Caption = ""
  lbltype.Caption = ""
  lblnumnetworks.Caption = ""
  lblnumhosts.Caption = ""
  lblRange.Caption = ""
  lstNetIDs.Clear
  Disable
  txtip(0).SetFocus
End Sub

Private Sub bttn_reset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "Resets the default Subnet Mask and the IP address.", bttn_reset
    Exit Sub
End If
End Sub

Private Sub Form_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub Form_Load()
Dim lhwnd As Long
  On Error Resume Next
  Dim style As Long
  Dim x As Integer
  Disable
  For x = 0 To txtip.Count - 1
    'Get the current style.
    style = GetWindowLong(txtip(x).hWnd, GWL_STYLE)
    style = GetWindowLong(txtsm(x).hWnd, GWL_STYLE)
    ' Add ES_NUMBER to the style.
    SetWindowLong txtip(x).hWnd, GWL_STYLE, style Or ES_NUMBER
    SetWindowLong txtsm(x).hWnd, GWL_STYLE, style Or ES_NUMBER
    ' Subclass to ignore the context menu.
    OldWindowProc = SetWindowLong(txtip(x).hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
    IPOldWindowProc(x) = OldWindowProc
    OldWindowProc = SetWindowLong(txtsm(x).hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
    SMOldWindowProc(x) = OldWindowProc
  Next x
  'Set The Controls To Look Like Office 2000
  'For x = 0 To frmmain.Count - 1
  '  lhwnd = Controls(x).hWnd
  '    If lhwnd <> 0 Then
  '     AddOfficeBorder lhwnd
  '    End If
    'On Error Resume Next
    AddOfficeBorder bttn_reset.hWnd
    AddOfficeBorder bttn_check.hWnd
    AddOfficeBorder bttn_defaultmask.hWnd
    AddOfficeBorder bttn_close.hWnd
    AddOfficeBorder picIP.hWnd
    AddOfficeBorder picSM.hWnd
    AddOfficeBorder lstNetIDs.hWnd
  'Next x
  'set the caption and the info label
  Me.Caption = "IP Calculator version " & App.Major & "." & App.Minor & "." & App.Revision & " by Ryan Conrad"
  lblinfo.Caption = "IP Calculator version " & App.Major & "." & App.Minor & "." & App.Revision & " by Ryan Conrad" & _
      vbCrLf & "Coded in Visual Basic 6 (sp4) on 03/25/01" & vbCrLf & "Visit http://camalot.virtualave.net"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Dim x As Integer
  'Reset the textboxes to the default
  For x = 0 To txtip.Count - 1
    SetWindowLong txtip(x).hWnd, GWL_WNDPROC, IPOldWindowProc(x)
    SetWindowLong txtsm(x).hWnd, GWL_WNDPROC, SMOldWindowProc(x)
  Next x
  'See You Later
  End
End Sub

Private Sub Frame1_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub Frame2_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub Frame3_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub Frame4_Click()
If IsTipVisible = True Then HideTip
End Sub

Private Sub Frame5_Click()
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub Label1_Click(Index As Integer)
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub lbl_Click(Index As Integer)
  'Hide our tip if its visible
  If IsTipVisible = True Then HideTip
End Sub

Private Sub lblbinary_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lbltip As String
  Select Case Index
    Case 0
      lbltip = "Displays the binary format of the IP address."
    Case 1
      lbltip = "Displays the binary format of the Subnet Mask."
    Case 2
      lbltip = "Displays the binary format of the Network ID."
  End Select
  If Button = 1 Then
    'Hide our tip if its visible
    If IsTipVisible = True Then HideTip
  Else
    ShowTip lbltip, lblbinary(Index)
  End If
End Sub

Private Sub lblclass_Click()
If IsTipVisible = True Then HideTip
End Sub

Private Sub lblclass_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "This is the Class of the IP address.", lblclass
    Exit Sub
End If
End Sub

Private Sub lblgood_Click()
If IsTipVisible = True Then HideTip
End Sub

Private Sub lblgood_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "Tells if the IP given is good for a host.", lblgood
    Exit Sub
End If
End Sub

Private Sub lblinfo_Click()
If IsTipVisible = True Then HideTip
End Sub

Private Sub lblNetID_Click()
If IsTipVisible = True Then HideTip
End Sub

Private Sub lblNetID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    ShowTip "Shows the Network ID for the IP address.", lblNetID
    Exit Sub
End If
End Sub

Private Sub picIP_Click()
  'if the IP picturebox is clicked set focus to the ip text box
  txtip(0).SetFocus
End Sub

Private Sub picSM_Click()
  'if the Mask picturebox is clicked set focus to the ip text box
  txtsm(0).SetFocus
End Sub

Private Sub txtip_Change(Index As Integer)
  On Error Resume Next
  'If the section = "" we need to put a value there
  If txtip(Index) = "" Then txtip(Index) = "0": SendKeys "{HOME}+{END}"
  'Now we need to set a range of numbers allowed.
  If CInt(txtip(Index).Text) > 255 Then
    MsgBox "Number must be between 0 - 255." & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
    SendKeys "{HOME}+{END}"
  End If
  If Len(txtip(Index).Text) = 3 Then
    If Index = txtip.Count - 1 Then
      txtsm(0).SetFocus
    Else
      txtip(Index + 1).SetFocus
    End If
  End If
End Sub

Private Sub txtip_Click(Index As Integer)
  'select the section
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtip_GotFocus(Index As Integer)
  'Select the section
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtip_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  Dim tindex As Integer
  'If the '.' or the 'Enter' Is pressed then goto the next section
  If KeyAscii = Asc(".") Or KeyAscii = 13 Then
    If Index = txtip.Count - 1 Then
      tindex = 0
      txtsm(tindex).SetFocus
    Else
      tindex = Index + 1
      txtip(tindex).SetFocus
    End If
  End If

End Sub

Private Sub txtip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  ShowTip "IP Octet #" & Index & " in IP address. Numbers range from 0 - 255.", txtip(Index)
End If
End Sub

'For the following txtsm_ See txtip_ for comments
Private Sub txtsm_Change(Index As Integer)
  On Error Resume Next
  If txtsm(Index) = "" Then txtsm(Index) = "0": SendKeys "{HOME}+{END}"
  If Len(txtsm(Index).Text) = 3 Then
    If Index = txtsm.Count - 1 Then
      txtip(0).SetFocus
    Else
      txtsm(Index + 1).SetFocus
    End If
  End If
End Sub

Private Sub txtsm_Click(Index As Integer)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtsm_GotFocus(Index As Integer)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtsm_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  Dim tindex As Integer
  If KeyAscii = Asc(".") Or KeyAscii = 13 Or KeyAscii = Asc(vbTab) Then
    If checkmask(Index) = True Then
      If Index = txtsm.Count - 1 Then
        tindex = 0
      Else
        tindex = Index + 1
      End If
    Else
      tindex = Index
      MsgBox "Number for mask must be:" & Chr(13) & "0, 128, 192, 224, 240, 248, 252, 254, or 255" & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
    End If
    txtsm(tindex).SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Function checkmask(Index As Integer) As Boolean
  'this returns true if the Mask is a valid mask
  If CInt(txtsm(Index).Text) <> 255 And CInt(txtsm(Index).Text) <> 0 And CInt(txtsm(Index).Text) <> 128 And _
      CInt(txtsm(Index).Text) <> 224 And CInt(txtsm(Index).Text) <> 240 And CInt(txtsm(Index).Text) <> 248 And _
      CInt(txtsm(Index).Text) <> 252 And CInt(txtsm(Index).Text) <> 254 Then
    checkmask = False
  Else
    checkmask = True
  End If
End Function

Public Function AddOfficeBorder(ByVal hWnd As Long)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function

