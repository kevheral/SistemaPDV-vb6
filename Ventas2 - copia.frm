VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Ventas2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10860
   ClientLeft      =   -360
   ClientTop       =   420
   ClientWidth     =   15270
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   73
      Top             =   1710
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   64
      Top             =   3075
      Width           =   3855
   End
   Begin VB.TextBox txt_Descuento 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   63
      Top             =   3570
      Width           =   3855
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Picture         =   "Ventas2 - copia.frx":0000
      TabIndex        =   62
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   61
      Top             =   1095
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   60
      Top             =   4065
      Width           =   3855
   End
   Begin VB.TextBox TXT_CANTIDAD 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   59
      Top             =   2580
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   58
      Top             =   2085
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXCENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   57
      Top             =   1215
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   6120
      TabIndex        =   7
      Top             =   1080
      Width           =   9015
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "V"
         Height          =   7455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8895
         Begin VB.CommandButton cmd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   120
            Width           =   1400
         End
         Begin VB.CommandButton cmd7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd10 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd11 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd12 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1020
            Width           =   1400
         End
         Begin VB.CommandButton cmd13 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd14 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd16 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd17 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd18 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1920
            Width           =   1400
         End
         Begin VB.CommandButton cmd19 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd20 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd21 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd22 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd23 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd24 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmd25 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd26 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd27 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd28 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd29 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd30 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   3720
            Width           =   1400
         End
         Begin VB.CommandButton cmd31 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd32 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd33 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd34 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd35 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd36 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4620
            Width           =   1400
         End
         Begin VB.CommandButton cmd37 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd38 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd39 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd40 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd41 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd42 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5520
            Width           =   1400
         End
         Begin VB.CommandButton cmd48 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   7320
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   6420
            Width           =   1400
         End
         Begin VB.CommandButton cmd47 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   5880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   6420
            Width           =   1400
         End
         Begin VB.CommandButton cmd46 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   6420
            Width           =   1400
         End
         Begin VB.CommandButton cmd45 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   6420
            Width           =   1400
         End
         Begin VB.CommandButton cmd44 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   6420
            Width           =   1400
         End
         Begin VB.CommandButton cmd43 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   6420
            Width           =   1400
         End
      End
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   14160
      Picture         =   "Ventas2 - copia.frx":50EA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "SALIR"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   12960
      Picture         =   "Ventas2 - copia.frx":9501
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "ELIMINAR"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   4560
      Picture         =   "Ventas2 - copia.frx":D650
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "MODIFICAR"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton BTN5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   4560
      Picture         =   "Ventas2 - copia.frx":11878
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "NUEVO"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Top             =   1590
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   955
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Ventas2 - copia.frx":15B8C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "MODIFICAR"
      Top             =   2280
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   120
      TabIndex        =   65
      Top             =   5520
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   66
      Top             =   8280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   72
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   71
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label LBL_VALOR 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   0
      TabIndex        =   70
      Top             =   4680
      Width           =   4455
   End
   Begin VB.Label Label17 
      Caption         =   "Conteo Factura "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   69
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6960
      TabIndex        =   68
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4320
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuskin 
      Caption         =   "Skin"
      Begin VB.Menu mnuskins 
         Caption         =   "[VBNet]"
         Index           =   0
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPYellow] "
         Index           =   1
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPBlue] "
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XP] "
         Index           =   3
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[XPPink] "
         Index           =   4
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[HotMail] "
         Index           =   5
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsMedia] "
         Index           =   6
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WMPCobre] "
         Index           =   7
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WMPVerde] "
         Index           =   8
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Word2000] "
         Index           =   9
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[RoyaleXP] "
         Index           =   10
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsLive] "
         Index           =   11
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[WindowsLive2] "
         Index           =   12
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Google Picasa]"
         Index           =   13
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 Blue TabStrip]"
         Index           =   14
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 Blue Button]"
         Index           =   15
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Office2007 check green]"
         Index           =   16
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Hotmail2]"
         Index           =   17
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Windows vista]"
         Index           =   18
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[CBblue]"
         Index           =   19
      End
      Begin VB.Menu mnuskins 
         Caption         =   "[Royale xp TaskBar]"
         Index           =   20
      End
      Begin VB.Menu mnuskins 
         Caption         =   "WinVista2"
         Index           =   21
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Opciones"
      Begin VB.Menu mnualign 
         Caption         =   "Alinear Menu"
         Begin VB.Menu mnualignMent 
            Caption         =   "No alinear"
            Index           =   0
         End
         Begin VB.Menu mnualignMent 
            Caption         =   "Arriba"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnualignMent 
            Caption         =   "Abajo"
            Index           =   2
         End
      End
      Begin VB.Menu mnuText 
         Caption         =   "Texto"
         Begin VB.Menu mnuCambiarFuente 
            Caption         =   "Cambiar fuente"
         End
         Begin VB.Menu mnuUnderLine 
            Caption         =   "Subrayar enlace"
            Begin VB.Menu mnuUnderLines 
               Caption         =   "Subrayar en MouseUP"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuUnderLines 
               Caption         =   "Subrayar en Check"
               Checked         =   -1  'True
               Index           =   1
            End
         End
         Begin VB.Menu mnuForeColor 
            Caption         =   "Color de fuente"
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorNormal"
               Index           =   0
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorUp"
               Index           =   1
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForecolorCheck"
               Index           =   2
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "ForeColorDisabled"
               Index           =   3
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuForeColors 
               Caption         =   "No usar Forecolor"
               Checked         =   -1  'True
               Index           =   5
            End
         End
      End
      Begin VB.Menu mnuScroll2 
         Caption         =   "Scroll"
         Begin VB.Menu mnuScrolls 
            Caption         =   "Aumentar"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuScrolls 
            Caption         =   "Disminuir"
            Index           =   1
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuScrolls 
            Caption         =   "Restaurar"
            Index           =   2
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuMar 
         Caption         =   "Margen"
         Begin VB.Menu mnuMargin 
            Caption         =   "Margen del texto"
            Index           =   0
         End
         Begin VB.Menu mnuMargin 
            Caption         =   "Margen del botón"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEna 
         Caption         =   "Enabled"
         Begin VB.Menu mnuEnabled 
            Caption         =   "Habilitar control"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuEnabled 
            Caption         =   "Habilitar/Deshabilitar botón"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDrag 
         Caption         =   "Drag"
         Begin VB.Menu mnuDragDrop 
            Caption         =   "Habilitar Drag Drop ( Debe estar el Align en 0)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuTam 
         Caption         =   "Tamaño del Menú"
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Chico"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Medio"
            Index           =   1
         End
         Begin VB.Menu mnuSizeMenu 
            Caption         =   "Mas grande"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLoadSkin 
         Caption         =   "Cargar Skin desde archivo"
         Begin VB.Menu mnuLoadSkins 
            Caption         =   "Skin rosa"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuScroll 
      Caption         =   "Scroll"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveScroll 
         Caption         =   "Ir al principio"
         Index           =   0
      End
      Begin VB.Menu mnuMoveScroll 
         Caption         =   "Ir al final"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Ventas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim RS_VENTAS As ADODB.Recordset
Dim RS_NoDoc As ADODB.Recordset
Dim RS_TOTAL As ADODB.Recordset
Dim RS_ELIMINAR As ADODB.Recordset
Dim RS_IMPUESTO As ADODB.Recordset
Dim RS_PRODUCTO As ADODB.Recordset
Dim RS_MODIFICAR As ADODB.Recordset
Dim RS_SALIDA As ADODB.Recordset
Dim RS_SALIDA1 As ADODB.Recordset
Dim DIM_SEGUIR As Boolean
'FIXIT: Declare 'DIM_CODIGO' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim DIM_CODIGO, a
Dim DIM_CODIGO_1
Dim DIM_INVENTARIO
Dim DIM_EXCENTO As Boolean
'FIXIT: Declare 'DIM_ITEM' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Dim DIM_ITEM
Dim DIM_ITEM1
Dim DIM_1
Dim DIM_NUM
Dim PUB_29, PUB_30, PUB_31, PUB_32, PUB_65, PUB_28, DIM_CANTIDAD
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQLITEM
'FIXIT: Declare 'DIM_PRODUCTOS' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_PRODUCTOS
'FIXIT: Declare 'DIM_PUNITARIO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_PUNITARIO
'FIXIT: Declare 'DIM_VALOR' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_VALOR
'FIXIT: Declare 'DIM_TOTAL' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_TOTAL
'FIXIT: Declare 'DIM_NODEI' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_NODEI
'FIXIT: Declare 'DIM_SELECT' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim DIM_SELECT
'FIXIT: Declare 'DIM_SELECT_1' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_SELECT_1
Dim RS_PRODUCTOS As ADODB.Recordset
'FIXIT: Declare 'DIM_SQL' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQL
'FIXIT: Declare 'DIM_VIEJO' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_VIEJO
Dim DIM_IMAGENP As Boolean
'FIXIT: Declare 'DIM_INT_1' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_1
'FIXIT: Declare 'DIM_INT_2' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_2
'FIXIT: Declare 'DIM_INT_3' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_3
'FIXIT: Declare 'DIM_INT_4' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_4
'FIXIT: Declare 'DIM_INT_5' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_5
'FIXIT: Declare 'DIM_INT_6' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_INT_6
'FIXIT: Declare 'DIM_RESULT' and 'DIM_MUESTRA' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_RESULT, DIM_MUESTRA
Dim DIM_INT_TIME_1 As Boolean
Dim DIM_INT_TIME_2 As Boolean
Dim DIM_INT_7 As Boolean
Dim DIM_INT_8 As Boolean
Dim DIM_INT_RS_2 As ADODB.Recordset
Dim nocli As ADODB.Recordset
Dim RS_VASIO As ADODB.Recordset
Dim RS_BORRAR As ADODB.Recordset
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Const TOP_MARGIN = 1
Const LEFT_MARGIN = 0
Private PRVT_1 As New ADODB.Connection
Private PRVT_2 As New ADODB.Command
Private PRVT_3 As New ADODB.Recordset

Option Explicit

Private Sub loadButtons()




End Sub

Private Sub BTN6_Click()

End Sub

Private Sub Command24_Click()

End Sub

Private Sub Command3_Click()
Text1.Text = ""
PUB_VALOR_C = ""
Calcular2.Show 1
End Sub

Private Sub Form_Load()
    
    Call mnuskins_Click(3)
    
       
DIM_SEGUIR = True
DIM_NUM = 1
lstvDatos_a_cero
AGREGAR_NUEVO
HOMBRE

End Sub

Private Sub mnualignMent_Click(Index As Integer)
    Call CheckMenuChange(mnualignMent, Index)

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSizeMenu_Click(Index As Integer)
    
    Call CheckMenuChange(mnuSizeMenu, Index)
    

End Sub

Private Sub mnuskins_Click(Index As Integer)
    
    Call CheckMenuChange(mnuskins, Index)
    
   
End Sub
Private Sub CheckMenuChange(pMenu As Object, lIndex As Integer)
    Dim xMenu As Variant
    For Each xMenu In pMenu
        xMenu.Checked = False
    Next
    pMenu(lIndex).Checked = True
End Sub
Private Sub txtFindBtn_GotFocus(Index As Integer)

End Sub

Private Sub txtFindBtn_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
       If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub ucBtn_Click(Index As Integer)


    
End Sub

Private Sub ucBtnFind_Click()


End Sub

Private Sub ucMenu1_ButtonClick( _
    ByVal ButtonIndex As Integer, _
    Button As cButton)
    
    With Button

        If .Caption = "VENTAS" Then
            Ventas.Show
        End If
        If .Caption = "REPORTES" Then
            Reportes.Show
        End If
        If .Caption = "EMPRESA" Then
            Ventas.Show
        End If
    End With
End Sub

Private Sub ucMenu1_ScrollContainerMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       Me.PopupMenu mnuScroll
    End If
End Sub
Private Sub ucMenu1_ButtonMouseOver(ByVal ButtonIndex As Integer)

End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub refrescar()

        Set RS_SALIDA = New Recordset
         RS_SALIDA.Open "SELECT Codigo,Producto,Salida,Descuento,punitario,ISV,Total,ClientE,NDVentas,fecha,Hora1,Descripcion,cliente,NoDE,DEI,FORMA,TARJETA,caja,COLOR,tipo,TAX FROM INVSalida1 WHERE NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
If DIM_INT_TIME_2 = True Then
'Text3.SetFocus
DIM_INT_TIME_2 = False
Else
DIM_INT_TIME_2 = False
End If
        Carga_lstvDatos
        
End Sub

Private Sub BTN1_Click()
On Error Resume Next
If RS_VENTAS.BOF = True And RS_VENTAS.EOF = True Then
Else
 RS_VENTAS.MoveFirst
 Label3 = RS_VENTAS.Fields("NoDoc4")
 refrescar
End If

End Sub

Private Sub BTN2_Click()
On Error Resume Next
    RS_VENTAS.MovePrevious
    
     'BTN3.Enabled = True
    ' BTN4.Enabled = True
       If RS_VENTAS.BOF = True Then
        RS_VENTAS.MoveFirst
        Label3 = RS_VENTAS.Fields("NoDoc4")
        refrescar

       Else
         refrescar
         Label3 = RS_VENTAS.Fields("NoDoc4")
       End If

End Sub

Private Sub BTN3_Click()
On Error Resume Next
 RS_VENTAS.MoveNext
 
 

       If RS_VENTAS.EOF = True Then
         RS_VENTAS.MoveLast
         Label3 = RS_VENTAS.Fields("NoDoc4")
         refrescar

       Else
        refrescar
        Label3 = RS_VENTAS.Fields("NoDoc4")
       End If

End Sub

Private Sub BTN4_Click()
On Error Resume Next
       RS_VENTAS.MoveLast
       Label3 = RS_VENTAS.Fields("NoDoc4")
       refrescar

End Sub

Private Sub BTN5_Click()
GUARDAR_NUEVO
'TXT_CODIGO.SetFocus
DIM_INT_TIME_1 = True
DIM_INT_TIME_2 = True
'DIM_NODOC = Label3
factura4.Show vbModal
Label4.Caption = ""
'Label2.Caption = ""
LBL_VALOR.Caption = ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_DEISUM As ADODB.Recordset
Set RS_DEISUM = New Recordset

RS_DEISUM.Open "Select sum(Total) from INVSalida1 WHERE NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'RS_CUENTAS_INGRESOS.Open "Select sum(Total) from INVSalida1 WHERE NDVentas like '" & RS_VENTAS.Fields("NoDoc4") & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'RS_SALIDA.Open "SELECT  FROM INVSalida1 WHERE NDVentas like '" & RS_VENTAS.Fields("NoDoc4") & "'", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic

'If IsNull(RS_DEISUM.Fields(0)) Then
DIM_SEGUIR = True
'Else

            If RS_DEISUM.Fields(0) >= 2000 Then
            Dim RS_DEI As ADODB.Recordset
            Set RS_DEI = New Recordset
            RS_DEI.Open "Select * from INVSalida1 WHERE NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
                 Do While Not RS_DEI.EOF
                 RS_DEI.Fields("node") = "0"
                 RS_DEI.Fields("dei") = False
                 RS_DEI.MoveNext
                 Loop
            RS_DEI.Close
            Set RS_DEI = Nothing
    
            End If
'End If
RS_DEISUM.Close
Set RS_DEISUM = Nothing

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

refrescar
AGREGAR_NUEVO

LBL_VALOR.Caption = ""
HOMBRE
LlenarInventarioTotal
DIM_CLIENTE = ""
DIM_RTNCIENTE = ""
End Sub


Private Sub BTN8_Click()
On Error Resume Next
factura.Show
End Sub

Private Sub BTN9_Click()
On Error Resume Next
If Text1.Text = "" Or TXT_CANTIDAD.Text = "" Then

  'RS_VENTAS.CancelUpdate
  RS_VENTAS.Requery
  RS_VENTAS.MoveLast
  RS_VENTAS.Delete
  RS_VENTAS.MovePrevious
  'PUBLIC_SUB_UNLOCK
 
    
Else
End If
Unload Me
End Sub

Private Sub CMD_NUEVO_Click()
On Error Resume Next
txt_Descuento.Text = "NUEVO"
txt_Descuento.Text = "VIEJO"
End Sub

Private Sub cm34_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd34.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cm35_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd35.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub CMD_VIEJO_Click()

End Sub



Private Sub cmd36_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd36.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd37_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd37.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
 
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd38_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd38.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        addcliente

End Sub

Private Sub cmd39_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd39.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd40_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd40.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub



Private Sub Command11_Click()
On Error Resume Next
Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQLITEM
DIM_SQLITEM = "SELECT * FROM Inventario01 where ID=" & 10
'DIM_SQLITEM = "SELECT * FROM Inventario01 where codigo = " & DIM_ITEM
RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly

        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente


End Sub

Private Sub cmd41_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd41.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd42_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd42.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd43_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd43.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd44_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd44.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd45_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd45.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
           LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd46_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd46.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd47_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd47.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd48_Click()

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd48.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
             LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub Command12_Click()
EliminarVentas2.Show 1

End Sub

Private Sub Command1_Click()
Text4.Text = "0.00"
DIM_EXCENTO = True
PUB_IMPUESTO = "0"
DIM_VIEJO_FORMA = True
End Sub

Private Sub cmd1_Click()
'On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd1.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

                
End Sub

Private Sub cmd10_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd10.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd11_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd11.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
           LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd12_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd12.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd13_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd13.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        LlenarDatosBoton
 
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd14_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd14.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd15_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd15.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd16_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd16.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
           LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd17_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd17.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd18_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd18.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd19_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd19.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton

        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd2_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd2.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
   
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd20_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd20.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd21_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd21.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd22_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd22.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd23_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd23.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd24_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd24.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
           LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd25_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd25.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd26_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd26.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd27_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd27.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd28_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd28.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd29_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd29.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd3_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd3.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd30_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd30.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
           LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd31_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd31.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd32_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd32.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd33_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd33.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd34_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd34.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd35_Click()
On Error Resume Next
        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd35.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub

Private Sub cmd4_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd4.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd5_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd5.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente
End Sub


Private Sub cmd6_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        
        DIM_SELECT = cmd6.Caption
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
            LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub

Private Sub cmd7_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd7.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        LlenarDatosBoton
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd8_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd8.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
          LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub cmd9_Click()
On Error Resume Next

        Set RS_PRODUCTO = New Recordset
'FIXIT: Declare 'DIM_SQLITEM' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
        Dim DIM_SQLITEM
        DIM_SELECT = cmd9.Caption
        
        DIM_SQLITEM = "SELECT * FROM   " & DIM_INVENTARIO & "  where nombre like '" & DIM_SELECT & "'"
        DIM_SQLITEM = DIM_SQLITEM & " AND tipo like '" & DIM_SELECT_1 & "'"
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, 3, 3
        
         LlenarDatosBoton
        
        CALCULAR_IMPUESTO

        addcliente

End Sub


Private Sub Command13_Click()

'
End Sub

Private Sub Command14_Click()
'On Error Resume Next

PUB_29 = Text3.Text
PUB_30 = Label4
PUB_31 = PUB_VALOR_C


'FIXIT: Declare 'DIM_1' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
'Dim DIM_1
DIM_IMAGENP = False
'FIXIT: Declare 'DIM_VALOR' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Dim DIM_VALOR
'NUEVA_CANTIDAD

'TXT_VALOR.Text = DIM_VALOR
lstvDatos_Ingresar

'Dim RS_TOTAL As ADODB.Recordset
Set RS_TOTAL = New Recordset
'like '" & DIM_NODOC & "'"
RS_TOTAL.Open "Select SUM(TOTAL),SUM(ISV),SUM(DESCUENTO) from INVSalida1 where NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'RS_TOTAL.Open "Select SUM(TOTAL),SUM(ISV),SUM(DESCUENTO) from INVSalida1 where NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
DIM_SUMTOTAL = RS_TOTAL.Fields(0)
DIM_SUMDESCUENTO = RS_TOTAL.Fields(2)
Set RS_TOTAL = Nothing
'DIM_SUMTOTAL = DIM_SUMTOTAL - DIM_SUMDESCUENTO
LBL_VALOR = "Total Lps." & Format(DIM_SUMTOTAL, "#,##0.00")
DIM_SEGUIR = True

'Carga_lstvDatos
Poner_datos
Command2.Enabled = True
DIM_EXCENTO = False
LlenarInventarioTotal
End Sub




Private Sub Command2_Click()
Cliente.Show vbModal
Text5.Text = DIM_CLIENTE
Text3.Text = DIM_RTNCIENTE
addcliente
End Sub

Private Sub Command23_Click()
On Error Resume Next
Poner_datos
DIM_IMAGENP = False
DIM_SEGUIR = True
'HABILITAR_VALOR
End Sub



Private Sub Command25_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 7
Else
txt_Descuento = txt_Descuento + "7"
End If
End Sub

Private Sub Command26_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 8
Else
txt_Descuento = txt_Descuento + "8"
End If
End Sub

Private Sub Command27_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 9
Else
txt_Descuento = txt_Descuento + "9"
End If
End Sub

Private Sub Command28_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 4
Else
txt_Descuento = txt_Descuento + "4"
End If
End Sub

Private Sub Command29_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 5
Else
txt_Descuento = txt_Descuento + "5"
End If
End Sub



Private Sub Command30_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 6
Else
txt_Descuento = txt_Descuento + "6"
End If
End Sub

Private Sub Command31_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 1
Else
txt_Descuento = txt_Descuento + "1"
End If
End Sub

Private Sub Command32_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 2
Else
txt_Descuento = txt_Descuento + "2"
End If
End Sub

Private Sub Command33_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 3
Else
txt_Descuento = txt_Descuento + "3"
End If
End Sub

Private Sub Command34_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 0
Else
txt_Descuento = txt_Descuento + "00"
End If
End Sub

Private Sub Command35_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = 0
Else
txt_Descuento = txt_Descuento + "0"
End If
End Sub

Private Sub Command36_Click()
On Error Resume Next
If txt_Descuento = "" Then
txt_Descuento = "000"
Else
txt_Descuento = txt_Descuento + "000"
End If
End Sub




Private Sub Command5_Click()
On Error Resume Next
If Text1.Text = "" Or TXT_CANTIDAD.Text = "" Then

  'RS_VENTAS.CancelUpdate
  RS_VENTAS.Requery
  RS_VENTAS.MoveLast
  RS_VENTAS.Delete
  RS_VENTAS.MovePrevious
  'PUBLIC_SUB_UNLOCK
 
    
Else
End If
Unload Me
End Sub

Private Sub Command8_Click()

End Sub



Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
DIM_ITEM = Item.SubItems(5)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim Borrar_Codigo
Dim Borrar_NoDoc
Dim Borrar_Hora
Dim i

If KeyCode = vbKeyDelete Then

    Dim Mens As Integer

           
            'Borrar_Codigo = LV.ListItems(0).Text
            Borrar_NoDoc = LV.SelectedItem.ListSubItems.Item(4).Text
            Borrar_Hora = LV.SelectedItem.ListSubItems.Item(6).Text
            
  
            
                    Dim DIM_SQLITEM As String
                    
                    Set RS_ELIMINAR = New Recordset
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Set RS_ELIMINAR = New Recordset
                    DIM_SQLITEM = "DELETE * FROM INVSalida1 where ID = " & DIM_ITEM
                    RS_ELIMINAR.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
                    Set RS_ELIMINAR = Nothing
                    
           
            
          For i = LV.ListItems.Count To 1 Step -1
                ' si está seleccionado
                If LV.ListItems(i).Selected Then
                    ' lo borramos
                    LV.ListItems.Remove i
                    
                End If
          Next
          
        End If



End Sub




Private Sub Salir_Click()
On Error Resume Next

End Sub

'FIXIT: Declare 'lstvDatos_Ingresar' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function lstvDatos_Ingresar()
'On Error Resume Next
PUB_32 = PUB_VALOR_C
'PUB_28 = Text3.Text
''''''''''''''''''''''''''''''''''''''''''''''
PUB_31 = PUB_VALOR_C
DIM_SUMTOTAL = Text2.Text
''''''''''''''''''''''''''''''''''''''''''''''
'PUB_41 = cmbtalla.Text
PUB_29 = Text3.Text
'PUB_41 = Combo1.Text
PUB_65 = 1


Set RS_SALIDA1 = New Recordset
'        RS_SALIDA.Open "SELECT Codigo,Producto,salida,Descuento,Valor,Saldo,NDVentas,Periodot,fecha,Hora1,caja FROM INVSalida1 WHERE NDVentas like'" & RS_VENTAS.Fields("NDOC"), PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
RS_SALIDA1.Open "SELECT * FROM INVSalida1", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic


With RS_SALIDA1
.AddNew
.Fields("Producto") = PUB_29
If PUB_30 = "" Then
.Fields("Salida") = 1
Else
.Fields("Salida") = PUB_30
End If
.Fields("PUnitario") = PUB_31

.Fields("Total") = DIM_1
.Fields("producto") = DIM_CODIGO
.Fields("caja") = "1"
.Fields("NDVentas") = DIM_NODOC
.Fields("codigo") = DIM_CODIGO_1
.Fields("Fecha") = Date
'.Fields("Tienda") = DIM_TIENDA
.Fields("Hora1") = Format(Time, "Long Time")
'.Fields("PUnitario") = PUB_70
'.Fields("Descripcion") = PUB_28
.Fields("Cliente") = PUB_28

If IsEmpty(PUB_IMPUESTO) Then
.Fields("ISV") = "0"
Else
.Fields("ISV") = PUB_IMPUESTO
End If

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
If DIM_VIEJO_FORMA = False Then
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
.Fields("tipo") = 0
.Fields("TAX") = "GRABADO"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
.Fields("tipo") = 1
.Fields("TAX") = "EXCENTO"
DIM_VIEJO_FORMA = False
End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            If PUB_CANTIDAD = False Then
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            .Fields("NoDE") = DIM_NODEI
            .Fields("DEI") = 1
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
            Else
            .Fields("NoDE") = "0"
            .Fields("DEI") = 0
            End If
            

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
DIM_VIEJO_FORMA = False
.Update
refrescar
Carga_lstvDatos
End With

Set RS_SALIDA1 = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
End Function
'FIXIT: Declare 'Limpiar_lstvDatos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Limpiar_lstvDatos()
On Error Resume Next
            LV.ListItems.Clear
End Function
'FIXIT: Declare 'Carga_lstvDatos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Carga_lstvDatos()


Set RS_SALIDA1 = New Recordset
'        RS_SALIDA.Open "SELECT Codigo,Producto,salida,Descuento,Valor,Saldo,NDVentas,Periodot,fecha,Hora1,caja FROM INVSalida1 WHERE NDVentas like'" & RS_VENTAS.Fields("NDOC"), PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
RS_SALIDA1.Open "SELECT codigo,producto,salida,total,ndventas,id FROM INVSalida1 WHERE NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic


 With RS_SALIDA1
        If .RecordCount <> 0 Then
        If RS_SALIDA.BOF = True And RS_SALIDA.EOF = True Then
        LV.ListItems.Clear
        Else
            LV.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Set Items = LV.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                Items.SubItems(3) = .Fields(3) & ""
                Items.SubItems(4) = .Fields(4) & ""
                Items.SubItems(5) = .Fields(5) & ""
                .MoveNext
            Loop
        End If
         End If
    End With
    
Set RS_SALIDA1 = Nothing
'    TXT_CODIGO.SetFocus
End Function
'FIXIT: Declare 'lstvDatos_a_cero' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function lstvDatos_a_cero()
On Error Resume Next
    With LV
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Codigo", 2000
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Total", 1000
        .ColumnHeaders.Add , , "N Doc", 1300
        .ColumnHeaders.Add , , "ID", 1300

        
    End With
End Function

Private Sub Poner_datos()
On Error Resume Next
'
'cmbtalla.Clear
'TXT_ID.Text = ""
'TXT_BANCO.Text = ""
'TXT_CUENTA.Text = ""
txt_Descuento.Text = ""
TXT_CANTIDAD.Text = ""

'Label2.Caption = ""
'Label7.Caption = ""
Text1.Text = ""
Text4.Text = ""
Text2.Text = ""
Text3.Text = ""
'Combo1.Text = "ANONIMO"

End Sub

Private Sub Command2_LostFocus()
On Error Resume Next
If DIM_INT_TIME_1 = True Then
DIM_INT_TIME_1 = False
Else
'Text3.SetFocus
End If
 Command2.BackColor = &HE0E0E0
End Sub


Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
End Sub

Private Sub Salir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
End Sub

Private Sub Text1_Change()
On Error Resume Next
If TXT_CANTIDAD.Text = "" Then
Else
                If PUB_VALOR_C = "" Then


                PUB_VALOR_C = Text1.Text

                'FRM_VENTASDIGITAL.Text1 = FRM_VENTASDIGITAL.Text1 + "5"
                DIM_SUBTOTAL = PUB_VALOR_C
                DIM_1 = PUB_VALOR_C * Label4
                LBL_VALOR = "Lps." & Format(DIM_1, "###,###,##0.00")
                Else
                'PUB_VALOR_C = Text1.Text

                'FRM_VENTASDIGITAL.Text1 = FRM_VENTASDIGITAL.Text1 + "5"
                DIM_SUBTOTAL = PUB_VALOR_C
                DIM_1 = PUB_VALOR_C * Label4
                LBL_VALOR = "Lps." & Format(DIM_1, "###,###,##0.00")
End If
        Text2.Text = Text1 * Label4
        LlenarDatosBoton
        CALCULAR_IMPUESTO
End If
End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
If Text1.Text = "" Then
    Text1.Text = "1"
Else
End If
End Sub

Private Sub TEXT1_LostFocus()
On Error Resume Next
'Text1.Text = Format(DIM_VALOR, "###,###,##0.00")
End Sub

Private Sub TXT_CANTIDAD_Change()
On Error Resume Next
If TXT_CANTIDAD.Text = "" Then
Else
If txt_Descuento = "" Or txt_Descuento = "0" Then
DIM_SUBTOTAL = DIM_PUNITARIO
Else
DIM_SUBTOTAL = PUB_VALOR_C
End If
Text2.Text = PUB_VALOR_C * Label4
'CALCULAR_IMPUESTO
End If


End Sub

Private Sub TXT_CANTIDAD_Click()
On Error Resume Next
TXT_CANTIDAD.Text = ""

End Sub

Private Sub TXT_CANTIDAD_GotFocus()
On Error Resume Next
If TXT_CANTIDAD.Text = "" Then
    TXT_CANTIDAD.Text = "1"
Else
End If

End Sub

Private Sub TXT_CANTIDAD_LostFocus()
On Error Resume Next
'lstvDatos_Ingresar

End Sub

'FIXIT: Declare 'VASIO' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Public Function VASIO()
On Error Resume Next
'FIXIT: Declare 'DIM_SQL_VASIO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQL_VASIO
'FIXIT: Declare 'DIM_SQL' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim DIM_SQL
'FIXIT: Declare 'DIM_A' and '_B' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim DIM_A, DIM_B

DIM_B = "0"
Set nocli = New Recordset
Set RS_VASIO = New Recordset
Set RS_BORRAR = New Recordset
'    DIM_SQL = "SELECT * FROM Ventas"
'    nocli.Open DIM_SQL, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly '
'Do While Not nocli.EOF
'DIM_A = nocli.Fields("nodoc")
'DIM_SQL_VASIO = "SELECT * FROM INVSalida1 WHERE ndventas = " & DIM_A
'       RS_VASIO.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
'If RS_VASIO.EOF Then
'    DIM_SQL_VASIO = "DELETE * FROM Ventas WHERE nodoc= " & DIM_A
'    RS_BORRAR.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockOptimistic
'     If nocli.EOF = True Then
'            Exit Function
'        End If
'        If nocli.BOF Then
'            Exit Function
'        End If
''
'    RS_BORRAR.Close
 ''   End If
 '       RS_VASIO.Close'''

       '
        'RS_VENTAS.MoveFirst
       'RS_VENTAS.MoveLast
'nocli.MoveNext
'Loop
'nocli.Close
    DIM_A = "0"
    DIM_SQL_VASIO = "DELETE * FROM INVSalida1 WHERE VALOR = " & DIM_A
    'DIM_SQL_VASIO = DIM_SQL_VASIO & " AND CODIGO = " & DIM_A
    RS_VASIO.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
    
    DIM_A = "0"
    DIM_SQL_VASIO = "DELETE * FROM INVSalida1 WHERE CODIGO = " & DIM_A
    'DIM_SQL_VASIO = DIM_SQL_VASIO & " AND CODIGO = " & DIM_A
    RS_VASIO.Open DIM_SQL_VASIO, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$'''''''''''''''''''''''''''''''''''''''''
End Function
'FIXIT: Declare 'AGREGAR_NUEVO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function AGREGAR_NUEVO()
'On Error Resume Next
'On Error Resume Next

Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventasdos Order by NoDoc4", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic

If RS_VENTAS.EOF = True Or RS_VENTAS.BOF = True Then
With RS_VENTAS

    
            DIM_NODOC = "1"
        Label16.Caption = Format(DIM_NODOC, "0000000#")
        DIM_NODOC = Format(DIM_NODOC, "0000000#")
       ' RS_VENTAS.Fields("NoDoc4") = DIM_NODOC
        

        
        Set RS_SALIDA = New Recordset
'        RS_SALIDA.Open "SELECT Codigo,Producto,salida,Descuento,Valor,Saldo,NDVentas,Periodot,fecha,Hora1,caja FROM INVSalida WHERE NDVentas like'" & RS_VENTAS.Fields("NDOC"), PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
        RS_SALIDA.Open "SELECT Producto,Salida,PUnitario,NDVentas,fecha,Hora1,Descripcion,Cliente,Id,caja,punitario,total,tipo,TAX,codigo FROM INVSalida", PUB_CONEXION_EASY, adOpenDynamic, adLockOptimistic
        'Set DIM_INT_RS_2 = New Recordset
        'DIM_INT_RS_2.Open "SELECT Codigo,Producto,Tipo,Egresos,Descuento,Valor,NDVentas2,fecha,Hora1,talla,Descripcion FROM INVSalida WHERE NDVentas2= " & RS_VENTAS.Fields("NDoc"), PUB_CONEXION_EASY_A, adOpenDynamic, adLockOptimistic

        Carga_lstvDatos
        
            Label16.Caption = Format(DIM_NODOC, "0000000#")
    DIM_DOCFIN = 12000
    DIM_REST = DIM_DOCFIN - DIM_NODOC
'    Label9.Caption = DIM_REST
    Label3.Caption = DIM_REST
    DIM_NODOC = Format(DIM_NODOC, "0000000#")
End With
Else


    If RS_VENTAS.BOF = False And RS_VENTAS.EOF = False Then
            RS_VENTAS.MoveFirst
            RS_VENTAS.MoveLast
    End If

RS_VENTAS.MoveLast
'DIM_NODOC = RS_VENTAS.Fields("NoDoc4") + 1
Limpiar_lstvDatos
    DIM_NODOC = RS_VENTAS.Fields("NoDoc4") + 1
    
        Label16.Caption = Format(DIM_NODOC, "0000000#")
    DIM_DOCFIN = 12000
    DIM_REST = DIM_DOCFIN - DIM_NODOC
'    Label9.Caption = DIM_REST
    Label3.Caption = DIM_REST
    DIM_NODOC = Format(DIM_NODOC, "0000000#")
    
End If

Set RS_VENTAS = Nothing


End Function
Public Function GUARDAR_NUEVO()

Set RS_VENTAS = New Recordset
RS_VENTAS.Open "select * from Ventasdos Order by NoDoc4", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
        
If RS_VENTAS.EOF = True Or RS_VENTAS.BOF = True Then

            RS_VENTAS.AddNew
    
            DIM_NODOC = "1"
            RS_VENTAS.AddNew
            
                        RS_VENTAS.Fields("NODOC1") = "000"
                        RS_VENTAS.Fields("NODOC2") = "001"
                        RS_VENTAS.Fields("NODOC3") = "01"
                        RS_VENTAS.Fields("NODOC4") = Format(DIM_NODOC, "0000000#")
                        DIM_NODOC = Format(DIM_NODOC, "0000000#")
            
            
            RS_VENTAS.Update
Else
        
            RS_VENTAS.MoveFirst
            RS_VENTAS.MoveLast


            DIM_NODOC = RS_VENTAS.Fields("NoDoc4") + 1
            RS_VENTAS.AddNew

            RS_VENTAS.Fields("NODOC1") = "000"
            RS_VENTAS.Fields("NODOC2") = "001"
            RS_VENTAS.Fields("NODOC3") = "01"
            RS_VENTAS.Fields("NODOC4") = Format(DIM_NODOC, "0000000#")
            DIM_NODOC = Format(DIM_NODOC, "0000000#")


RS_VENTAS.Update

End If
Set RS_VENTAS = Nothing


End Function

'FIXIT: Declare 'NUEVA_CANTIDAD' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function NUEVA_CANTIDAD()
On Error Resume Next
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset
RS_CUENTAS_INGRESOS.Open "Select * from CANTIDAD", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
PUB_CANTIDAD1 = RS_CUENTAS_INGRESOS.Fields("CANTIDAD")
RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'FIXIT: Declare 'DIM_VALOR1' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim DIM_VALOR1
'Dim RS_CUENTAS_INGRESOS As ADODB.Recordset
Set RS_CUENTAS_INGRESOS = New Recordset

RS_CUENTAS_INGRESOS.Open "Select SUM(Total) from INVSalida1 WHERE Fecha like '" & Date & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'like '" & RS_VENTAS.Fields("NoDoc4") & "'"

If IsNull(RS_CUENTAS_INGRESOS.Fields(0)) Then
DIM_VALOR1 = "0"
Else
DIM_VALOR1 = RS_CUENTAS_INGRESOS.Fields(0)
'Label12.Caption = Format(DIM_VALOR1, "#,##0.00")
End If

If Val(PUB_CANTIDAD1) >= Val(DIM_VALOR1) Then
PUB_CANTIDAD = False
Else
PUB_CANTIDAD = True
End If

RS_CUENTAS_INGRESOS.Close
Set RS_CUENTAS_INGRESOS = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
'FIXIT: Declare 'HABILITAR_VALOR' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE



Private Sub txt_Descuento_GotFocus()

On Error Resume Next
Frame1.Enabled = True
End Sub
Public Sub TextSelected()
On Error Resume Next
Dim i As Integer
'FIXIT: Declare 'oMyTextBox' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Dim oMyTextBox As Object


Set oMyTextBox = Screen.ActiveControl
If TypeName(oMyTextBox) = "TextBox" Then
i = Len(oMyTextBox.Text)
oMyTextBox.SelStart = 0
oMyTextBox.SelLength = i
End If

End Sub
'FIXIT: Declare 'CALCULAR_IMPUESTO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function CALCULAR_IMPUESTO()
On Error Resume Next

If DIM_EXCENTO = True Then
Else
        
        DIM_SUBTOTAL = (PUB_VALOR_C * Label4) / 1.15
        DIM_SUBTOTAL = Format(DIM_SUBTOTAL, "#,##0.00")
    
        PUB_IMPUESTO = DIM_SUBTOTAL
        PUB_IMPUESTO = PUB_IMPUESTO * 15 / 100
        Text4.Text = "Impuesto = " + Format(PUB_IMPUESTO, "#,##0.00")
        DIM_VIEJO = Text4
End If

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
'FIXIT: Declare 'CALCULAR_DESCUENTO' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE


Public Function HOMBRE()
On Error Resume Next
DIM_SELECT_1 = "ZAPATOS"
DIM_INVENTARIO = "inventario01"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario01 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function MUJER()
On Error Resume Next
DIM_SELECT_1 = "MUJER"
DIM_INVENTARIO = "Inventario02"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario02 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function INVIERNO()
On Error Resume Next
DIM_SELECT_1 = "INVIERNO"
DIM_INVENTARIO = "inventario03"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario03 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function VERANO()
On Error Resume Next
DIM_SELECT_1 = "VERANO"
DIM_INVENTARIO = "inventario04"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario04 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function HOGAR()
On Error Resume Next
DIM_SELECT_1 = "HOGAR"
DIM_INVENTARIO = "inventario05"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario05 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function NIÑO()
On Error Resume Next
DIM_SELECT_1 = "NIÑO"
DIM_INVENTARIO = "inventario06"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario06 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function ACCESORIO()
On Error Resume Next
DIM_SELECT_1 = "ACCESORIO"
DIM_INVENTARIO = "inventario07"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario07 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Public Function OTROS()
On Error Resume Next
DIM_SELECT_1 = "OTROS"
DIM_INVENTARIO = "inventario08"

                    'cmd1.Picture = LoadPicture("")k
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = ""
                    cmd1.BackColor = vbWhite
                    cmd2.Caption = ""
                    cmd2.BackColor = vbWhite
                    cmd3.Caption = ""
                    cmd3.BackColor = vbWhite
                    cmd4.Caption = ""
                    cmd4.BackColor = vbWhite
                    cmd5.Caption = ""
                    cmd5.BackColor = vbWhite
                    cmd6.Caption = ""
                    cmd6.BackColor = vbWhite
                    cmd7.Caption = ""
                    cmd7.BackColor = vbWhite
                    cmd8.Caption = ""
                    cmd8.BackColor = vbWhite
                    cmd9.Caption = ""
                    cmd9.BackColor = vbWhite
                    cmd10.Caption = ""
                    cmd10.BackColor = vbWhite
                    cmd11.Caption = ""
                    cmd11.BackColor = vbWhite
                    cmd12.Caption = ""
                    cmd12.BackColor = vbWhite
                    cmd13.Caption = ""
                    cmd13.BackColor = vbWhite
                    cmd14.Caption = ""
                    cmd14.BackColor = vbWhite
                    cmd15.Caption = ""
                    cmd15.BackColor = vbWhite
                    cmd16.Caption = ""
                    cmd16.BackColor = vbWhite
                    cmd17.Caption = ""
                    cmd17.BackColor = vbWhite
                    cmd18.Caption = ""
                    cmd18.BackColor = vbWhite
                    cmd19.Caption = ""
                    cmd19.BackColor = vbWhite
                    cmd20.Caption = ""
                    cmd20.BackColor = vbWhite
                    cmd21.Caption = ""
                    cmd21.BackColor = vbWhite
                    cmd22.Caption = ""
                    cmd22.BackColor = vbWhite
                    cmd23.Caption = ""
                    cmd23.BackColor = vbWhite
                    cmd24.Caption = ""
                    cmd24.BackColor = vbWhite
                    cmd25.Caption = ""
                    cmd25.BackColor = vbWhite
                    cmd26.Caption = ""
                    cmd26.BackColor = vbWhite
                    cmd27.Caption = ""
                    cmd27.BackColor = vbWhite
                    cmd28.Caption = ""
                    cmd28.BackColor = vbWhite
                    cmd29.Caption = ""
                    cmd29.BackColor = vbWhite
                    cmd30.Caption = ""
                    cmd30.BackColor = vbWhite
                    cmd31.Caption = ""
                    cmd31.BackColor = vbWhite
                    cmd32.Caption = ""
                    cmd32.BackColor = vbWhite
                    cmd33.Caption = ""
                    cmd33.BackColor = vbWhite
                    cmd34.Caption = ""
                    cmd34.BackColor = vbWhite
                    cmd35.Caption = ""
                    cmd35.BackColor = vbWhite
                    cmd36.Caption = ""
                    cmd36.BackColor = vbWhite
                    cmd37.Caption = ""
                    cmd37.BackColor = vbWhite
                    cmd38.Caption = ""
                    cmd38.BackColor = vbWhite
                    cmd39.Caption = ""
                    cmd39.BackColor = vbWhite
                    cmd40.Caption = ""
                    cmd40.BackColor = vbWhite
        Set RS_PRODUCTO = New Recordset
        Dim DIM_SQLITEM
        DIM_SQLITEM = "SELECT * FROM Inventario08 "
        RS_PRODUCTO.Open DIM_SQLITEM, PUB_CONEXION_EASY, adOpenStatic, adLockReadOnly
        With RS_PRODUCTO
        If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
         Exit Function
        Else
        With RS_PRODUCTO
            .MoveFirst
            Do While Not .EOF
            
            
            If .Fields("ID") = 1 Then
                    'cmd1.Picture = LoadPicture(App.Path & .Fields("IMAGEN"))
                    cmd1.BackColor = vbWhite
                    cmd1.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 2 Then
                    cmd2.BackColor = vbWhite
                    cmd2.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 3 Then
                    cmd3.BackColor = vbWhite
                    cmd3.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 4 Then
                    cmd4.BackColor = vbWhite
                    cmd4.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 5 Then
                    cmd5.BackColor = vbWhite
                    cmd5.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 6 Then
                    cmd6.BackColor = vbWhite
                    cmd6.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 7 Then
                    cmd7.BackColor = vbWhite
                    cmd7.Caption = .Fields("nombre")
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Fields("ID") = 8 Then
                    cmd8.BackColor = vbWhite
                    cmd8.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 9 Then
                    cmd9.BackColor = vbWhite
                    cmd9.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 10 Then
                    cmd10.BackColor = vbWhite
                    cmd10.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 11 Then
                    cmd11.BackColor = vbWhite
                    cmd11.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 12 Then
                    cmd12.BackColor = vbWhite
                    cmd12.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 13 Then
                    cmd13.BackColor = vbWhite
                    cmd13.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 14 Then
                    cmd14.BackColor = vbWhite
                    cmd14.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 15 Then
                    cmd15.BackColor = vbWhite
                    cmd15.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 16 Then
                    cmd16.BackColor = vbWhite
                    cmd16.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 17 Then
                    cmd17.BackColor = vbWhite
                    cmd17.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 18 Then
                    cmd18.BackColor = vbWhite
                    cmd18.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 19 Then
                    cmd19.BackColor = vbWhite
                    cmd19.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 20 Then
                    cmd20.BackColor = vbWhite
                    cmd20.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 21 Then
                    cmd21.BackColor = vbWhite
                    cmd21.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 22 Then
                    cmd22.BackColor = vbWhite
                    cmd22.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 23 Then
                    cmd23.BackColor = vbWhite
                    cmd23.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 24 Then
                    cmd24.BackColor = vbWhite
                    cmd24.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 25 Then
                    cmd25.BackColor = vbWhite
                    cmd25.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 26 Then
                    cmd26.BackColor = vbWhite
                    cmd26.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 27 Then
                    cmd27.BackColor = vbWhite
                    cmd27.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 28 Then
                    cmd28.BackColor = vbWhite
                    cmd28.Caption = .Fields("nombre")
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            If .Fields("ID") = 29 Then
                    cmd29.BackColor = vbWhite
                    cmd29.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 30 Then
                    cmd30.BackColor = vbWhite
                    cmd30.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 31 Then
                    cmd31.BackColor = vbWhite
                    cmd31.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 32 Then
                    cmd32.BackColor = vbWhite
                    cmd32.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 33 Then
                    cmd33.BackColor = vbWhite
                    cmd33.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 34 Then
                    cmd34.BackColor = vbWhite
                    cmd34.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 35 Then
                    cmd35.BackColor = vbWhite
                    cmd35.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 36 Then
                    cmd36.BackColor = vbWhite
                    cmd36.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 37 Then
                    cmd37.BackColor = vbWhite
                    cmd37.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 38 Then
                    cmd38.BackColor = vbWhite
                    cmd38.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 39 Then
                    cmd39.BackColor = vbWhite
                    cmd39.Caption = .Fields("nombre")
            End If
            If .Fields("ID") = 40 Then
                    cmd40.BackColor = vbWhite
                    cmd40.Caption = .Fields("nombre")
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                .MoveNext
            Loop
            

        End With
        End If
        End With
        Set RS_PRODUCTO = Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Sub VScroll1_Change()
  
End Sub
Public Function addcliente()
If DIM_CLIENTE = "" Then
Text5.Text = "Cliente"
Text3.Text = "0"
Else
Text5.Text = DIM_CLIENTE
Text3.Text = DIM_RTNCIENTE
End If
End Function
Public Function LlenarInventarioTotal()
'On Error Resume Next

    With ListView1
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add , , "Codigo", 2000
        .ColumnHeaders.Add , , "Producto", 2500
        .ColumnHeaders.Add , , "Cantidad", 1000
        .ColumnHeaders.Add , , "Total", 1000
    End With
    
    
Dim RS_TOTAL As ADODB.Recordset
Set RS_TOTAL = New Recordset
'like '" & DIM_NODOC & "'"
RS_TOTAL.Open "Select codigo,nombre,egresos,total from GR_INVENTARIO_SALDO ", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
'RS_TOTAL.Open "Select SUM(TOTAL),SUM(ISV),SUM(DESCUENTO) from INVSalida1 where NDVentas like '" & DIM_NODOC & "'", PUB_CONEXION_EASY, adOpenKeyset, adLockOptimistic
With RS_TOTAL
        If .RecordCount <> 0 Then
        If RS_TOTAL.BOF = True And RS_TOTAL.EOF = True Then
        ListView1.ListItems.Clear
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF

                Set Items = ListView1.ListItems.Add(, , .Fields(0) & "")
                Items.SubItems(1) = .Fields(1) & ""
                Items.SubItems(2) = .Fields(2) & ""
                Items.SubItems(3) = .Fields(3) & ""
                .MoveNext
            Loop
        End If
         End If
    End With
Set RS_TOTAL = Nothing
 
End Function
Public Function LlenarDatosBoton()

            If RS_PRODUCTO.EOF = True And RS_PRODUCTO.BOF = True Then
             Exit Function
            Else
            txt_Descuento.Text = RS_PRODUCTO.Fields("NOMBRE")
            DIM_SUBTOTAL = RS_PRODUCTO.Fields("PUNITARIO")
                    If DIM_SEGUIR = True Then
                    PUB_VALOR_C = RS_PRODUCTO.Fields("VALOR")
                    End If
            
            DIM_PUNITARIO = RS_PRODUCTO.Fields("PUNITARIO")
            DIM_CODIGO = RS_PRODUCTO.Fields("nombre")
            DIM_CODIGO_1 = RS_PRODUCTO.Fields("codigo")
            Text4.Text = RS_PRODUCTO.Fields("ISVV")
            End If
        Set RS_PRODUCTO = Nothing
        TXT_CANTIDAD.SetFocus
            If TXT_CANTIDAD.Text = "" Then
                Label4.Caption = "1"
                TXT_CANTIDAD.Text = "Cantidad = 1"
                Text2.Text = "Valor = " + Format(PUB_VALOR_C, "###,###,##0.00")
                LBL_VALOR = "Lps." & Format(PUB_VALOR_C, "###,###,##0.00")
                Text1.Text = Format(PUB_VALOR_C, "###,###,##0")
                DIM_1 = PUB_VALOR_C
            Else
            Label4.Caption = Val(Label4) + "1"
            TXT_CANTIDAD.Text = "Cantidad = " + Label4.Caption
            DIM_1 = PUB_VALOR_C * Label4
            LBL_VALOR = "Lps." & Format(DIM_1, "###,###,##0.00")
            Text1.Text = ""
            Text1.Text = Format(PUB_VALOR_C, "###,###,##0")
            Text2.Text = "Valor = " + Format(DIM_1, "###,###,##0")
            CALCULAR_IMPUESTO
            End If
End Function


