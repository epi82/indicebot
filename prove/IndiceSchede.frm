VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form IndiceSchede 
   Caption         =   "Crea Lista Schede Botaniche"
   ClientHeight    =   9900
   ClientLeft      =   240
   ClientTop       =   915
   ClientWidth     =   14790
   Icon            =   "IndiceSchede.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   14790
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Lista Schede"
      TabPicture(0)   =   "IndiceSchede.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "mshSchedaSelezionata"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "mshSchede"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "mshListaSchedaSelezionata"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxTestoScheda"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "mshListaSchede"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ListaSchede"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Ordine per Famiglie"
      TabPicture(1)   =   "IndiceSchede.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "LabelFamiglia"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "NumFamiglie"
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(5)=   "LabelNumeroFamiglie"
      Tab(1).Control(6)=   "InfoFamiglie"
      Tab(1).Control(7)=   "mshListaFamiglie"
      Tab(1).Control(8)=   "ListaFamiglie"
      Tab(1).Control(9)=   "CreaHtmlFamiglie1"
      Tab(1).Control(10)=   "CreaHtmlFamiglie2"
      Tab(1).Control(11)=   "CreaHtmlFamiglie3"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Ordine per Generi"
      TabPicture(2)   =   "IndiceSchede.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "LabelGenere"
      Tab(2).Control(2)=   "LabelNumeroSpecie"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "InfoGeneri"
      Tab(2).Control(5)=   "mshListaGeneri"
      Tab(2).Control(6)=   "ListaGeneri"
      Tab(2).Control(7)=   "CreaCodiceGeneri"
      Tab(2).Control(8)=   "CreaHtmlGeneri3"
      Tab(2).Control(9)=   "CreaHtmlGeneri2"
      Tab(2).Control(10)=   "CreaHtmlGeneri1"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Ordine per Specie"
      TabPicture(3)   =   "IndiceSchede.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3"
      Tab(3).Control(1)=   "LabelSpecie"
      Tab(3).Control(2)=   "InfoSpecie"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "LabelNumSpecie"
      Tab(3).Control(5)=   "mshListaSpecie"
      Tab(3).Control(6)=   "ListaSpecie"
      Tab(3).Control(7)=   "CreaCodiceSpecie"
      Tab(3).Control(8)=   "CreaHtmlSpecie1"
      Tab(3).Control(9)=   "CreaHtmlSpecie2"
      Tab(3).Control(10)=   "CreaHtmlSpecie3"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Ordine Nome Comune"
      TabPicture(4)   =   "IndiceSchede.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "LabelNomeComune"
      Tab(4).Control(2)=   "Label11"
      Tab(4).Control(3)=   "LabelNumeroNomi"
      Tab(4).Control(4)=   "InfoNomi"
      Tab(4).Control(5)=   "mshListaNomiComuni"
      Tab(4).Control(6)=   "CreaHtmlNomi3"
      Tab(4).Control(7)=   "CreaHtmlNomi2"
      Tab(4).Control(8)=   "CreaHtmlNomi1"
      Tab(4).Control(9)=   "ListaNomiComuni"
      Tab(4).Control(10)=   "CreaCodiceNomeComune"
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "Testo HTML"
      TabPicture(5)   =   "IndiceSchede.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "CodiceHTML"
      Tab(5).ControlCount=   1
      Begin VB.ListBox ListaSchede 
         Height          =   8055
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton CreaCodiceNomeComune 
         Caption         =   "Crea Codice HTML"
         Height          =   975
         Left            =   -73920
         TabIndex        =   33
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox ListaNomiComuni 
         Height          =   8055
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton CreaHtmlNomi1 
         Caption         =   "Crea File HTML   Fase 1 Testata"
         Height          =   975
         Left            =   -74040
         TabIndex        =   31
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlNomi2 
         Caption         =   "Crea Corpo HTML Lista  corrente"
         Enabled         =   0   'False
         Height          =   3255
         Left            =   -74040
         TabIndex        =   30
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlNomi3 
         Caption         =   "Chiudi File HTML"
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74040
         TabIndex        =   29
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlSpecie3 
         Caption         =   "Chiudi File HTML"
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74040
         TabIndex        =   27
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlSpecie2 
         Caption         =   "Crea Corpo HTML Lista  corrente"
         Enabled         =   0   'False
         Height          =   3255
         Left            =   -74040
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlSpecie1 
         Caption         =   "Crea File HTML   Fase 1 Testata"
         Height          =   975
         Left            =   -74040
         TabIndex        =   25
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlGeneri1 
         Caption         =   "    CREA TESTATA     file HTML"
         Height          =   975
         Left            =   -74040
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlGeneri2 
         Caption         =   "      CREA CORPO       file HTML"
         Enabled         =   0   'False
         Height          =   3255
         Left            =   -74040
         TabIndex        =   23
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlGeneri3 
         Caption         =   "CHIUDI File HTML"
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74040
         TabIndex        =   22
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton CreaHtmlFamiglie3 
         Caption         =   "Fase 3 - Chiudi File HTML"
         Enabled         =   0   'False
         Height          =   975
         Left            =   -72240
         TabIndex        =   21
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton CreaHtmlFamiglie2 
         Caption         =   "Fase 2 - Crea Corpo HTML della Famiglia selezionata"
         Enabled         =   0   'False
         Height          =   3495
         Left            =   -72240
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton CreaHtmlFamiglie1 
         Caption         =   "Fase 1 - Crea Testata del File HTML"
         Height          =   975
         Left            =   -72240
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton CreaCodiceSpecie 
         Caption         =   "Crea Codice HTML"
         Height          =   975
         Left            =   -73920
         TabIndex        =   18
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CreaCodiceGeneri 
         Caption         =   "Crea Codice HTML"
         Height          =   975
         Left            =   -73920
         TabIndex        =   16
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox ListaSpecie 
         Height          =   8250
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.ListBox ListaGeneri 
         Height          =   8250
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.ListBox ListaFamiglie 
         Height          =   8055
         Left            =   -74760
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaSchede 
         Bindings        =   "IndiceSchede.frx":03B2
         Height          =   615
         Left            =   9120
         TabIndex        =   3
         Top             =   8040
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1085
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaFamiglie 
         Bindings        =   "IndiceSchede.frx":03CF
         Height          =   7095
         Left            =   -70440
         TabIndex        =   4
         Top             =   1080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   12515
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaGeneri 
         Bindings        =   "IndiceSchede.frx":03EE
         Height          =   6975
         Left            =   -72120
         TabIndex        =   6
         Top             =   1320
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12303
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaSpecie 
         Bindings        =   "IndiceSchede.frx":040B
         Height          =   6855
         Left            =   -72120
         TabIndex        =   7
         Top             =   1320
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12091
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin RichTextLib.RichTextBox CodiceHTML 
         Height          =   8535
         Left            =   -71040
         TabIndex        =   28
         Top             =   480
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   15055
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"IndiceSchede.frx":0428
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaNomiComuni 
         Bindings        =   "IndiceSchede.frx":04AA
         Height          =   6975
         Left            =   -72120
         TabIndex        =   34
         Top             =   1320
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12303
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin RichTextLib.RichTextBox TxTestoScheda 
         Height          =   3975
         Left            =   3840
         TabIndex        =   56
         Top             =   720
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7011
         _Version        =   393217
         TextRTF         =   $"IndiceSchede.frx":04CB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshListaSchedaSelezionata 
         Bindings        =   "IndiceSchede.frx":0556
         Height          =   615
         Left            =   8880
         TabIndex        =   58
         Top             =   7440
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1085
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSchede 
         Bindings        =   "IndiceSchede.frx":057E
         Height          =   2655
         Left            =   3840
         TabIndex        =   59
         Top             =   5040
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4683
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSchedaSelezionata 
         Bindings        =   "IndiceSchede.frx":0596
         Height          =   615
         Left            =   3840
         TabIndex        =   60
         Top             =   8160
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1085
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label16 
         Caption         =   "Scheda selezionata"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   63
         Top             =   7920
         Width           =   4935
      End
      Begin VB.Label Label15 
         Caption         =   "Dati disponibili nel Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   62
         Top             =   4800
         Width           =   4935
      End
      Begin VB.Label Label14 
         Caption         =   "Lista Schede"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label13 
         Caption         =   "Testo Invision per Schede ""Botaniche"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   57
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label InfoFamiglie 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -72120
         TabIndex        =   54
         Top             =   8400
         Width           =   11415
      End
      Begin VB.Label LabelNumeroFamiglie 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -62040
         TabIndex        =   53
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Numero Specie selezionate"
         Height          =   375
         Left            =   -64200
         TabIndex        =   52
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label InfoNomi 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -72120
         TabIndex        =   51
         Top             =   8400
         Width           =   11415
      End
      Begin VB.Label LabelNumeroNomi 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -64680
         TabIndex        =   50
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Numero Specie selezionate"
         Height          =   375
         Left            =   -66840
         TabIndex        =   49
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label LabelNumSpecie 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -64920
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Numero Specie selezionate"
         Height          =   375
         Left            =   -67080
         TabIndex        =   47
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label InfoSpecie 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -72120
         TabIndex        =   46
         Top             =   8280
         Width           =   11415
      End
      Begin VB.Label InfoGeneri 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -72120
         TabIndex        =   45
         Top             =   8400
         Width           =   11415
      End
      Begin VB.Label Label10 
         Caption         =   "Numero Specie selezionate"
         Height          =   375
         Left            =   -67200
         TabIndex        =   44
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label LabelNumeroSpecie 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -65040
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
      Begin VB.Label NumFamiglie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73440
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Num.Famiglie"
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LabelNomeComune 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67560
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Lista Schede dei Nomi Comuni che iniziano per"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   35
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label LabelSpecie 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67920
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Lista Schede delle Specie che iniziano per"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   14
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label LabelGenere 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67920
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Lista Schede dei Generi che iniziano per"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   12
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label LabelFamiglia 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67320
         TabIndex        =   11
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Lista delle Schede della Famiglia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   10
         Top             =   720
         Width           =   3135
      End
   End
   Begin MSAdodcLib.Adodc AdoListaSchede 
      Height          =   330
      Left            =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaSchede"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox ListaIni 
      Height          =   255
      Left            =   11880
      TabIndex        =   1
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   9360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoListaGeneri 
      Height          =   330
      Left            =   0
      Top             =   9600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaGeneri"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoListaFamiglie 
      Height          =   330
      Left            =   2520
      Top             =   9240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaFamiglie"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoListaSpecie 
      Height          =   330
      Left            =   2520
      Top             =   9600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaSpecie"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoListaNomiComuni 
      Height          =   330
      Left            =   5160
      Top             =   9600
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaNomiComuni"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoListaSchedaSelezionata 
      Height          =   330
      Left            =   4320
      Top             =   9240
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaSchede"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AsoListaSchedaSelezionata 
      Height          =   330
      Left            =   7560
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoListaSchedaSelezionata"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoSchede 
      Height          =   330
      Left            =   10320
      Top             =   9600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoSchede"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoSchedaSelezionata 
      Height          =   330
      Left            =   9480
      Top             =   9120
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoSchedaSelezionata"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LabelData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12120
      TabIndex        =   42
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Data aggiornamento"
      Height          =   255
      Left            =   10560
      TabIndex        =   41
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label LabelNum_Schede 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9360
      TabIndex        =   40
      Top             =   9360
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Num.Schede totali"
      Height          =   255
      Left            =   7920
      TabIndex        =   39
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   9360
      Width           =   2655
   End
End
Attribute VB_Name = "IndiceSchede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DbDati As String
Dim Riga As String
Dim sorgente As String
Dim Num_Schede As Double
Dim Num_Fam As Double
Dim Fam1, Fam2, Fam3, Fam4, Fam5, Fam6 As Integer
Dim genere_cor As String
Dim specie_cor As String
Dim specie1_cor As String
Dim nome_cor As String
Dim nome1_cor As String
Dim genere1_cor As String
Dim famiglia_cor As String
Dim famiglia1_cor As String
Dim nomecomune_cor As String
Dim i As Integer
Dim FamigliaOld As String
Dim FamigliaNew As String
Dim GenereOld As String
Dim GenereNew As String
Dim SpecieOld As String
Dim NomeComuneNew As String
Dim NomeComuneOld As String
Dim SpecieNew As String
Dim Testo As String
Dim ImageCom As String
Dim TxTipo As String
Dim TxTopic As String
Dim TxNome As String
Dim TxFamiglia As String
Dim TxCom As String
Dim TxNomeComune As String
Dim TxGenere As String
Dim TxSpecie As String
Dim TxAutore As String
Dim TxAutoreScheda As String
Dim TxNew As String
Dim TxNN As String
Dim Riga1 As String
Dim Riga2 As String
Dim Riga3 As String
Dim Riga4 As String
Dim Riga5 As String
Dim Riga6 As String
Dim Riga7 As String
Dim Riga8 As String
Dim FileOri As String
Dim FileOri2 As String
Dim FileTmp As String
Dim CodiceTabellaIni As String
Dim CodiceFine As String
Dim DataAgg As String
Dim SchedaSelezionata  As String


Private Sub CreaCodiceGeneri_Click()
Testo = ""
For i = 1 To mshListaGeneri.Rows - 1
mshListaGeneri.Row = i
' prima riga del codice HTML
Riga1 = "<tr height=" & Chr(34) & "18" & Chr(34) & ">" & Chr(13)
' Riga 2 = Tipo di Scheda
mshListaGeneri.Col = 0
TxTipo = mshListaGeneri.Text
Riga2 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "3%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">" & TxTipo & "</div></td>" & Chr(13)
' Riga 3 Link Topic e Nome specie
' Topic
mshListaGeneri.Col = 9
TxTopic = Chr(34) & mshListaGeneri.Text & Chr(34)
' Genere
mshListaGeneri.Col = 2 ' Genere
TxGenere = mshListaGeneri.Text
' Specie
mshListaGeneri.Col = 3 ' Specie
TxSpecie = mshListaGeneri.Text
' Autore
mshListaGeneri.Col = 4 ' Autore
TxAutore = mshListaGeneri.Text
' NN
mshListaGeneri.Col = 5 ' NN
TxNN = mshListaGeneri.Text
TxNome = TxGenere & " " & TxSpecie
If TxNN <> "" Then
TxAutore = TxAutore & " <b>" & TxNN & "</b>"
End If
Riga3 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "40%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><a class=" & Chr(34) & "link2" & Chr(34) & "href=" & TxTopic & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxNome & "</a>  " & TxAutore & "<br></td>" & Chr(13)
' Riga 4 Campo Commestibile, Velenoso, Officinale
mshListaGeneri.Col = 6
If mshListaGeneri.Text = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaGeneri.Text = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaGeneri.Text = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaGeneri.Text = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaGeneri.Text = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
'ImageCom = "images/bullet_VO.gif"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>" & Chr(13)
End If
' Riga 5 Famiglia
mshListaGeneri.Col = 7
TxFamiglia = mshListaGeneri.Text
Riga5 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "20%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & ">" & TxFamiglia & "<br></td>" & Chr(13)
' Riga 6 NomeComune
mshListaGeneri.Col = 8
TxNomeComune = mshListaGeneri.Text
Riga6 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & ">" & TxNomeComune & "</td>" & Chr(13)
Riga7 = "</tr>" & Chr(13)
Testo = Testo & Riga1 & Riga2 & Riga3 & Riga4 & Riga5 & Riga6 & Riga7
CodiceHTML.Text = Testo
Next i
SSTab1.Tab = 5
End Sub


Private Sub CreaCodiceNomeComune_Click()
Testo = ""
For i = 1 To mshListaNomiComuni.Rows - 1
mshListaNomiComuni.Row = i
' prima riga del codice HTML
Riga1 = "<tr height=" & Chr(34) & "18" & Chr(34) & ">" & Chr(13)
' Riga 2 = Tipo di Scheda
mshListaNomiComuni.Col = 0
TxTipo = mshListaNomiComuni.Text
Riga2 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "3%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">" & TxTipo & "</div></td>" & Chr(13)
' Riga 3 Link Topic e Nome specie
' Topic
mshListaNomiComuni.Col = 9
TxTopic = Chr(34) & mshListaNomiComuni.Text & Chr(34)
mshListaNomiComuni.Col = 8
TxNomeComune = mshListaNomiComuni.Text
' Genere
mshListaNomiComuni.Col = 2 ' Genere
TxGenere = mshListaNomiComuni.Text
' Specie
mshListaNomiComuni.Col = 3 ' Specie
TxSpecie = mshListaNomiComuni.Text
' Autore
mshListaNomiComuni.Col = 4 ' Autore
TxAutore = mshListaNomiComuni.Text
' NN
mshListaNomiComuni.Col = 5 ' NN
TxNN = mshListaNomiComuni.Text
TxNome = TxGenere & " " & TxSpecie
If TxNN <> "" Then
TxAutore = TxAutore & " <b>" & TxNN & "</b>"
End If
Riga3 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "35%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><a class=" & Chr(34) & "link" & Chr(34) & "href=" & TxTopic & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxNomeComune & "</a>  " & "<br></td>" & Chr(13)
' Riga 4 Campo Commestibile, Velenoso, Officinale
mshListaNomiComuni.Col = 6
If mshListaNomiComuni.Text = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaNomiComuni.Text = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaNomiComuni.Text = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaNomiComuni.Text = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaNomiComuni.Text = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
'ImageCom = "images/bullet_VO.gif"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>" & Chr(13)
End If
' Riga 5 Famiglia
mshListaNomiComuni.Col = 7
TxFamiglia = mshListaNomiComuni.Text
Riga5 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "40%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><b>" & TxNome & "</b> " & TxAutore & "<br></td>" & Chr(13)
' Riga 6 NomeComune
mshListaNomiComuni.Col = 8
TxNomeComune = mshListaNomiComuni.Text
Riga6 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & ">" & TxFamiglia & "</td>" & Chr(13)
Riga7 = "</tr>" & Chr(13)
Testo = Testo & Riga1 & Riga2 & Riga3 & Riga4 & Riga5 & Riga6 & Riga7
CodiceHTML.Text = Testo
Next i
SSTab1.Tab = 5
End Sub

Private Sub CreaCodiceSpecie_Click()
Testo = ""
For i = 1 To mshListaSpecie.Rows - 1
mshListaSpecie.Row = i
' prima riga del codice HTML
Riga1 = "<tr height=" & Chr(34) & "18" & Chr(34) & ">" & Chr(13)
' Riga 2 = Tipo di Scheda
mshListaSpecie.Col = 0
TxTipo = mshListaSpecie.Text
Riga2 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "3%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">" & TxTipo & "</div></td>" & Chr(13)
' Riga 3 Link Topic e Nome specie
' Topic
mshListaSpecie.Col = 9
TxTopic = Chr(34) & mshListaSpecie.Text & Chr(34)
' Genere
mshListaSpecie.Col = 3 ' Genere
TxGenere = mshListaSpecie.Text
' Specie
mshListaSpecie.Col = 2 ' Specie
TxSpecie = mshListaSpecie.Text
' Autore
mshListaSpecie.Col = 4 ' Autore
TxAutore = mshListaSpecie.Text
' NN
mshListaSpecie.Col = 5 ' NN
TxNN = mshListaSpecie.Text
TxNome = TxSpecie & ", " & TxGenere
If TxNN <> "" Then
TxNN = " <b>" & TxNN & "</b>"
Else
TxNN = " "
End If
Riga3 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "40%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><a class=" & Chr(34) & "link2" & Chr(34) & "href=" & TxTopic & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxNome & "</a>  " & TxNN & "<br></td>" & Chr(13)
' Riga 4 Campo Commestibile, Velenoso, Officinale
mshListaSpecie.Col = 6
If mshListaSpecie.Text = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaSpecie.Text = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaSpecie.Text = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaSpecie.Text = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
ElseIf mshListaSpecie.Text = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>" & Chr(13)
'ImageCom = "images/bullet_VO.gif"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "5%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>" & Chr(13)
End If
' Riga 5 Famiglia
mshListaSpecie.Col = 7
TxFamiglia = mshListaSpecie.Text
Riga5 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " width=" & Chr(34) & "20%" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & ">" & TxFamiglia & "<br></td>" & Chr(13)
' Riga 6 NomeComune
mshListaSpecie.Col = 8
TxNomeComune = mshListaSpecie.Text
Riga6 = "<td bgcolor=" & Chr(34) & "#f8f8ff" & Chr(34) & " height=" & Chr(34) & "18" & Chr(34) & ">" & TxNomeComune & "</td>" & Chr(13)
Riga7 = "</tr>" & Chr(13)
Testo = Testo & Riga1 & Riga2 & Riga3 & Riga4 & Riga5 & Riga6 & Riga7
CodiceHTML.Text = Testo
Next i
SSTab1.Tab = 5
End Sub

Private Sub CreaHtmlFamiglie2_Click()
For i = 0 To ListaFamiglie.ListCount - 1
ListaFamiglie.ListIndex = i
LabelFamiglia.Caption = ListaFamiglie.List(i)
famiglia_cor = ListaFamiglie.List(i)
famiglia1_cor = "'" & ListaFamiglie.List(i) & "'"
AdoListaFamiglie.RecordSource = "SELECT Tipo,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,AutoreScheda,Topic,New From Lista_Schede where Famiglia = " & famiglia1_cor & " Order by Genere,Specie,nn"
AdoListaFamiglie.Refresh
LabelNumeroFamiglie.Caption = mshListaFamiglie.Rows - 1
Dim Msg3, Style3, Titolo3, Risposta3
Msg3 = "Confermare la prosecuzione della procedura con i generi della Famiglia  " & Chr(13) & Chr(13) & famiglia1_cor & " ?" & Chr(13) & Chr(13) & "Premere SI e attendere la comparsa del messaggio successivo."  ' Definisce il messaggio.
Style3 = vbYesNoCancel + vbCritical + vbDefaultButton2   ' Definisce i pulsanti.
Titolo3 = "Crea Corpo File HTML Generi"   ' Definisce il titolo.
      ' Visualizza il messaggio.
Risposta3 = MsgBox(Msg3, Style3, Titolo3)
If Risposta3 = vbYes Then   ' L'utente sceglie il pulsante Sì.
' Prima fase: parte iniziale della Tabella Famiglie
' Prima fase: parte iniziale
Open FileTmp For Append As #1
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "1" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "96" & Chr(34) & "><a href=" & Chr(34) & "#testata" & Chr(34) & "><img src=" & Chr(34) & "images/su2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "19" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "770" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "#00008b" & Chr(34) & "><a name=" & Chr(34) & LabelFamiglia.Caption & Chr(34) & "><b>" & LabelFamiglia.Caption & "</b></a></font></div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & "><div align=" & Chr(34) & "right" & Chr(34) & "></div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table class=" & Chr(34) & "testo2" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">Scheda</td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Genere e specie</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Com.</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Famiglia</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Nome comune</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Autore</div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table class=" & Chr(34) & "Testo0" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Close #1


Dim y As Integer
' Seconda fase: Corpo della Tabella
For y = 1 To mshListaFamiglie.Rows - 1
' Seconda fase: Corpo della Tabella
mshListaFamiglie.Row = y
mshListaFamiglie.Col = 0 ' Tipo
TxTipo = mshListaFamiglie.Text
mshListaFamiglie.Col = 1 ' Genere
TxGenere = mshListaFamiglie.Text
mshListaFamiglie.Col = 2 ' Specie
TxSpecie = mshListaFamiglie.Text
mshListaFamiglie.Col = 3 ' Autore
TxAutore = mshListaFamiglie.Text
mshListaFamiglie.Col = 4 ' nn
TxNN = mshListaFamiglie.Text
mshListaFamiglie.Col = 5 ' Com
TxCom = mshListaFamiglie.Text
mshListaFamiglie.Col = 6 ' Famiglia
TxFamiglia = mshListaFamiglie.Text
mshListaFamiglie.Col = 7 ' NomeComune
TxNomeComune = mshListaFamiglie.Text
mshListaFamiglie.Col = 8 ' AutoreScheda
TxAutoreScheda = mshListaFamiglie.Text
mshListaFamiglie.Col = 9 ' Topic
TxTopic = mshListaFamiglie.Text
mshListaFamiglie.Col = 10 ' New
TxNew = mshListaFamiglie.Text

' Riga1
Riga1 = "<tr height=" & Chr(34) & "21" & Chr(34) & ">"
' Riga2 = Tipo Scheda
If TxTipo = "+" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda1.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxTipo = "-" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda4.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
End If
' Riga3 = Genere e specie
    If TxNew <> "" Then ' Riga3 = Genere e Specie (mette NEW se il campo New è attivo)
    '<td bgcolor="#DFE6EF" width="336" height="21"><a class="link2"href="http://forum.funghiitaliani.it/index.php?showtopic=32640" target="_blank">Abutilon theophrasti</a>  Medicus<br></td>
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxGenere & " " & TxSpecie & "</a> " & TxAutore & "  <img src=" & Chr(34) & "images/new.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "14" & Chr(34) & " width=" & Chr(34) & "30" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></td>"
    Else
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxGenere & " " & TxSpecie & "</a> " & TxAutore & "</td>"
    End If
' Riga4 = Com
If TxCom = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>"
End If
' Riga5 = Famiglia
'<td bgcolor="#DFE6EF" width="116" height="21">Malvaceae<br></td>
Riga5 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxFamiglia & "</td>"
' Riga6 = Nome Comune
'<td bgcolor="#DFE6EF" width="250" height="21">Cencio molle<br></td>
Riga6 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxNomeComune & "</td>"
' Riga8 = AutoreScheda
'<td bgcolor="#DFE6EF" height="21"><div align="left">Mirna e Attilio</div></td>
Riga7 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "left" & Chr(34) & ">" & TxAutoreScheda & "</div></td>"
Riga8 = "</tr>"

Open FileTmp For Append As #2
    Print #2, Riga1
    Print #2, Riga2
    Print #2, Riga3
    Print #2, Riga4
    Print #2, Riga5
    Print #2, Riga6
    Print #2, Riga7
    Print #2, Riga8
    'Print #2, Riga9
Close #2
Next y

Open FileTmp For Append As #3
    Print #3, "</table>"
    Print #3, "</div>"
    Print #3, "</td>"
    Print #3, "</tr>"
    Print #3, "</table>"
    Print #3, "<br>"
Close #3

ElseIf Risposta3 = vbNo Then

Else   ' L'utente sceglie il pulsante No.
   Exit For   ' Esegue un'azione.
End If
Next i
CreaHtmlFamiglie2.Enabled = False
CreaHtmlFamiglie3.Enabled = True
InfoFamiglie.Caption = "ATTENZIONE: la procedura di creazione del CORPO è terminata. Selezionare il pulsante <CHIUDI File HTML> per completare il file <schede_bot_famiglie.html>"
End Sub

Private Sub CreaHtmlFamiglie1_Click()
' creazione file html - copia parte iniziale
FileOri = CurDir & "/Codicetestata-famiglie.txt"
FileOri2 = CurDir & "/Codicetestata-famiglie2.txt"
FileTmp = CurDir & "/schede_bot_famiglie.html"
FileCopy FileOri, FileTmp
' creazione file html - Inserimento numero Schede e data aggiornamento
Open FileTmp For Append As #1
Print #1, "<font color=" & Chr(34) & "red" & Chr(34) & "> n. " & Num_Schede & " Schede - aggiornato al " & DataAgg & "</font><br>"
    Open FileOri2 For Input As #2
        Do While Not EOF(2)   ' Ripete fino alla fine del file.
        Line Input #2, Riga   ' Assegna la riga a una variabile.
        Print #1, Riga    ' Scrive nella finestra Immediata.
        Loop
    Close #2
Print #1, "<table class=" & Chr(34) & "Table_Corpo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td>"
Print #1, "<div class=" & Chr(34) & "Testo3" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Close #1
CreaHtmlFamiglie1.Enabled = False
CreaHtmlFamiglie2.Enabled = True
InfoFamiglie.Caption = "ATTENZIONE: per proseguire con la creazione del File HTML premere il pulsante <CREA CORPO file HTML>, compare un Messaggio che prevede la conferma per ogni lettera presente nella Lista."
End Sub

Private Sub CreaHtmlFamiglie3_Click()
Open FileTmp For Append As #4
    Print #4, "</div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "<br>"
    Print #4, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "32" & Chr(34) & ">"
    Print #4, "<tr height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<td height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
    Print #4, "<b><font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & ">A.M.I.N.T. - Associazione Micologica Italiana Naturalistica Telematica</font></b></div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "</center>"
    Print #4, "</body>"
    Print #4, "</html>"
Close #4
End Sub

Private Sub CreaHtmlGeneri1_Click()
FileOri = CurDir & "/Codicetestata-Generi.txt"
FileTmp = CurDir & "/schede_bot_generi.html"
FileCopy FileOri, FileTmp
Open FileTmp For Append As #1
Print #1, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "88" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<td height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & "><b>Indice delle Schede botaniche in ordine di Genere</b></font><br>"
Print #1, "<table width=" & Chr(34) & "40" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "40" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & "></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & ">(a cura di Giuliano Salvai - realizzazione Web di Gianni Dose)</font><br>"
Print #1, "<hr width=" & Chr(34) & "400" & Chr(34) & ">"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & "><b>" & Num_Schede & " Schede - aggiornamento " & LabelData.Caption & "</b></font></div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "<center>"
Print #1, "<table class=" & Chr(34) & "Table_Corpo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td>"
Print #1, "<div class=" & Chr(34) & "Testo3" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Close #1
CreaHtmlGeneri1.Enabled = False
CreaHtmlGeneri2.Enabled = True
InfoGeneri.Caption = "ATTENZIONE: per proseguire con la creazione del File HTML premere il pulsante <CREA CORPO file HTML>, compare un Messaggio che prevede la conferma per ogni lettera presente nella Lista."
End Sub

Private Sub CreaHtmlGeneri2_Click()
' Fase 2 - crea corpo del file HTML
For i = 0 To ListaGeneri.ListCount - 1
ListaGeneri.ListIndex = i
LabelGenere.Caption = ListaGeneri.List(i)
genere_cor = ListaGeneri.List(i)
genere1_cor = "'" & ListaGeneri.List(i) & "'"
AdoListaGeneri.RecordSource = "SELECT Tipo,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,AutoreScheda,Topic,New From Lista_Schede where Mid(Genere,1,1) = " & genere1_cor & " Order by Genere,Specie,nn"
AdoListaGeneri.Refresh
LabelNumeroSpecie.Caption = mshListaGeneri.Rows - 1
Dim Msg3, Style3, Titolo3, Risposta3
Msg3 = "Confermare la prosecuzione della procedura con i generi che cominciano per  " & Chr(13) & Chr(13) & genere1_cor & " ?" & Chr(13) & Chr(13) & "Premere SI e attendere la comparsa del messaggio successivo."  ' Definisce il messaggio.
Style3 = vbYesNoCancel + vbCritical + vbDefaultButton2   ' Definisce i pulsanti.
Titolo3 = "Crea Corpo File HTML Generi"   ' Definisce il titolo.
      ' Visualizza il messaggio.
Risposta3 = MsgBox(Msg3, Style3, Titolo3)
If Risposta3 = vbYes Then   ' L'utente sceglie il pulsante Sì.
' Prima fase: parte iniziale
Open FileTmp For Append As #1
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "1" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "96" & Chr(34) & "><a href=" & Chr(34) & "#testata" & Chr(34) & "><img src=" & Chr(34) & "images/su2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "19" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "770" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "5" & Chr(34) & " color=" & Chr(34) & "#00008b" & Chr(34) & "><a name=" & Chr(34) & LabelGenere.Caption & Chr(34) & "><b>" & LabelGenere.Caption & "</b></a></font></div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & "><div align=" & Chr(34) & "right" & Chr(34) & "></div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table class=" & Chr(34) & "testo2" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">Scheda</td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Genere e specie</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Com.</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Famiglia</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Nome comune</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Autore</div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table class=" & Chr(34) & "Testo0" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Close #1

Dim y As Integer
' Seconda fase: Corpo della Tabella
For y = 1 To mshListaGeneri.Rows - 1
' Seconda fase: Corpo della Tabella
mshListaGeneri.Row = y
mshListaGeneri.Col = 0 ' Tipo
TxTipo = mshListaGeneri.Text
mshListaGeneri.Col = 1 ' Genere
TxGenere = mshListaGeneri.Text
mshListaGeneri.Col = 2 ' Specie
TxSpecie = mshListaGeneri.Text
mshListaGeneri.Col = 3 ' Autore
TxAutore = mshListaGeneri.Text
mshListaGeneri.Col = 4 ' nn
TxNN = mshListaGeneri.Text
mshListaGeneri.Col = 5 ' Com
TxCom = mshListaGeneri.Text
mshListaGeneri.Col = 6 ' Famiglia
TxFamiglia = mshListaGeneri.Text
mshListaGeneri.Col = 7 ' NomeComune
TxNomeComune = mshListaGeneri.Text
mshListaGeneri.Col = 8 ' AutoreScheda
TxAutoreScheda = mshListaGeneri.Text
mshListaGeneri.Col = 9 ' Topic
TxTopic = mshListaGeneri.Text
mshListaGeneri.Col = 10 ' New
TxNew = mshListaGeneri.Text

' Riga1
Riga1 = "<tr height=" & Chr(34) & "21" & Chr(34) & ">"
' Riga2 = Tipo Scheda
If TxTipo = "+" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda1.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxTipo = "-" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda4.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
End If
' Riga3 = Genere e specie
    If TxNew <> "" Then ' Riga3 = Genere e Specie (mette NEW se il campo New è attivo)
    '<td bgcolor="#DFE6EF" width="336" height="21"><a class="link2"href="http://forum.funghiitaliani.it/index.php?showtopic=32640" target="_blank">Abutilon theophrasti</a>  Medicus<br></td>
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxGenere & " " & TxSpecie & "</a> " & TxAutore & "  <img src=" & Chr(34) & "images/new.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "14" & Chr(34) & " width=" & Chr(34) & "30" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></td>"
    Else
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxGenere & " " & TxSpecie & "</a> " & TxAutore & "</td>"
    End If
' Riga4 = Com
If TxCom = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>"
End If
' Riga5 = Famiglia
'<td bgcolor="#DFE6EF" width="116" height="21">Malvaceae<br></td>
Riga5 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxFamiglia & "</td>"
' Riga6 = Nome Comune
'<td bgcolor="#DFE6EF" width="250" height="21">Cencio molle<br></td>
Riga6 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxNomeComune & "</td>"
' Riga8 = AutoreScheda
'<td bgcolor="#DFE6EF" height="21"><div align="left">Mirna e Attilio</div></td>
Riga7 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "left" & Chr(34) & ">" & TxAutoreScheda & "</div></td>"
Riga8 = "</tr>"

Open FileTmp For Append As #2
    Print #2, Riga1
    Print #2, Riga2
    Print #2, Riga3
    Print #2, Riga4
    Print #2, Riga5
    Print #2, Riga6
    Print #2, Riga7
    Print #2, Riga8
    'Print #2, Riga9
Close #2
Next y

Open FileTmp For Append As #3
    Print #3, "</table>"
    Print #3, "</div>"
    Print #3, "</td>"
    Print #3, "</tr>"
    Print #3, "</table>"
    Print #3, "<br>"
Close #3

ElseIf Risposta3 = vbNo Then

Else   ' L'utente sceglie il pulsante No.
   Exit For   ' Esegue un'azione.
End If
Next i
CreaHtmlGeneri2.Enabled = False
CreaHtmlGeneri3.Enabled = True
InfoGeneri.Caption = "ATTENZIONE: la procedura di creazione del CORPO è terminata. Selezionare il pulsante <CHIUDI File HTML> per completare il file <schede_bot_generi.html>"
End Sub

Private Sub CreaHtmlGeneri3_Click()
Open FileTmp For Append As #4
    Print #4, "</div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "<br>"
    Print #4, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "32" & Chr(34) & ">"
    Print #4, "<tr height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<td height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
    Print #4, "<b><font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & ">A.M.I.N.T. - Associazione Micologica Italiana Naturalistica Telematica</font></b></div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "</center>"
    Print #4, "</body>"
    Print #4, "</html>"
Close #4
End Sub

Private Sub CreaHtmlNomi1_Click()
FileOri = CurDir & "/Codicetestata-Nomi.txt"
FileTmp = CurDir & "/schede_bot_nomi.html"
FileCopy FileOri, FileTmp
Open FileTmp For Append As #1
Print #1, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "88" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<td height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & "><b>Indice delle Schede botaniche in ordine di Nome Comune</b></font><br>"
Print #1, "<table width=" & Chr(34) & "40" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "40" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & "></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & ">(a cura di Giuliano Salvai - realizzazione Web di Gianni Dose)</font><br>"
Print #1, "<hr width=" & Chr(34) & "400" & Chr(34) & ">"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & "><b>" & Num_Schede & " Schede - aggiornamento " & LabelData.Caption & "</b></font></div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "<center>"
Print #1, "<table class=" & Chr(34) & "Table_Corpo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td>"
Print #1, "<div class=" & Chr(34) & "Testo3" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Close #1
CreaHtmlNomi1.Enabled = False
CreaHtmlNomi2.Enabled = True
InfoNomi.Caption = "ATTENZIONE: per proseguire con la creazione del File HTML premere il pulsante <CREA CORPO file HTML>, compare un Messaggio che prevede la conferma per ogni lettera presente nella Lista."
End Sub

Private Sub CreaHtmlNomi2_Click()
' Prima fase: parte iniziale della Tabella Nomi Comuni
' Fase 2 - crea corpo del file HTML
For i = 0 To ListaNomiComuni.ListCount - 1
ListaNomiComuni.ListIndex = i
LabelNomeComune.Caption = ListaNomiComuni.List(i)
nome_cor = ListaNomiComuni.List(i)
nome1_cor = "'" & ListaNomiComuni.List(i) & "'"
AdoListaNomiComuni.RecordSource = "SELECT Tipo,NomeComune,Genere,Specie,Autore,nn,Com,Famiglia,AutoreScheda,Topic,New From Lista_Schede where Mid(NomeComune,1,1) = " & nome1_cor & " Order by NomeComune"
AdoListaNomiComuni.Refresh
LabelNumeroNomi.Caption = mshListaNomiComuni.Rows - 1
Dim Msg3, Style3, Titolo3, Risposta3
Msg3 = "Confermare la prosecuzione della procedura con i Nomi Comuni che cominciano per  " & Chr(13) & Chr(13) & nome1_cor & " ?" & Chr(13) & Chr(13) & "Premere SI e attendere la comparsa del messaggio successivo."  ' Definisce il messaggio.
Style3 = vbYesNoCancel + vbCritical + vbDefaultButton2   ' Definisce i pulsanti.
Titolo3 = "Crea Corpo File HTML NomiComuni"   ' Definisce il titolo.
      ' Visualizza il messaggio.
Risposta3 = MsgBox(Msg3, Style3, Titolo3)
If Risposta3 = vbYes Then   ' L'utente sceglie il pulsante Sì.
' Prima fase: parte iniziale
Open FileTmp For Append As #1
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "1" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "96" & Chr(34) & "><a href=" & Chr(34) & "#testata" & Chr(34) & "><img src=" & Chr(34) & "images/su2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "19" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "770" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "5" & Chr(34) & " color=" & Chr(34) & "#00008b" & Chr(34) & "><a name=" & Chr(34) & LabelNomeComune.Caption & Chr(34) & "><b>" & LabelNomeComune.Caption & "</b></a></font></div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & "><div align=" & Chr(34) & "right" & Chr(34) & "></div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table class=" & Chr(34) & "testo2" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">Scheda</td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Nome comune</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Com.</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Famiglia</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Genere e specie</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Autore</div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table class=" & Chr(34) & "Testo0" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Close #1

Dim y As Integer
' Seconda fase: Corpo della Tabella
For y = 1 To mshListaNomiComuni.Rows - 1
mshListaNomiComuni.Row = y
mshListaNomiComuni.Col = 0 ' Tipo
TxTipo = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 1 ' NomeComune
TxNomeComune = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 2 ' NomeComune
TxGenere = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 3 ' Specie
TxSpecie = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 4 ' Autore
TxAutore = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 5 ' nn
TxNN = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 6 ' Com
TxCom = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 7 ' Famiglia
TxFamiglia = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 8 ' AutoreScheda
TxAutoreScheda = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 9 ' Topic
TxTopic = mshListaNomiComuni.Text
mshListaNomiComuni.Col = 10 ' New
TxNew = mshListaNomiComuni.Text

' Riga1
Riga1 = "<tr height=" & Chr(34) & "21" & Chr(34) & ">"
' Riga2 = Tipo Scheda
If TxTipo = "+" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda1.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxTipo = "-" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda4.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
End If
' Riga3 = NomeComune
    If TxNew <> "" Then ' Riga3 = Genere e Specie (mette NEW se il campo New è attivo)
    '<td bgcolor="#DFE6EF" width="336" height="21"><a class="link2"href="http://forum.funghiitaliani.it/index.php?showtopic=32640" target="_blank">Abutilon theophrasti</a>  Medicus<br></td>
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxNomeComune & "</a><img src=" & Chr(34) & "images/new.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "14" & Chr(34) & " width=" & Chr(34) & "30" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></td>"
    Else
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxNomeComune & "</a></td>"
    End If
' Riga4 = Com
If TxCom = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>"
End If
' Riga5 = Famiglia
'<td bgcolor="#DFE6EF" width="116" height="21">Malvaceae<br></td>
Riga5 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxFamiglia & "</td>"
' Riga6 = Nome Comune
'<td bgcolor="#DFE6EF" width="250" height="21">Cencio molle<br></td>
Riga6 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><b>" & TxGenere & " " & TxSpecie & "</b> " & TxAutore & "</td>"
' Riga8 = AutoreScheda
'<td bgcolor="#DFE6EF" height="21"><div align="left">Mirna e Attilio</div></td>
Riga7 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "left" & Chr(34) & ">" & TxAutoreScheda & "</div></td>"
Riga8 = "</tr>"

Open FileTmp For Append As #2
    Print #2, Riga1
    Print #2, Riga2
    Print #2, Riga3
    Print #2, Riga4
    Print #2, Riga5
    Print #2, Riga6
    Print #2, Riga7
    Print #2, Riga8
    'Print #2, Riga9
Close #2
Next y

Open FileTmp For Append As #3
    Print #3, "</table>"
    Print #3, "</div>"
    Print #3, "</td>"
    Print #3, "</tr>"
    Print #3, "</table>"
    Print #3, "<br>"
Close #3

ElseIf Risposta3 = vbNo Then

Else   ' L'utente sceglie il pulsante No.
   Exit For   ' Esegue un'azione.
End If
Next i
CreaHtmlNomi2.Enabled = False
CreaHtmlNomi3.Enabled = True
InfoNomi.Caption = "ATTENZIONE: la procedura di creazione del CORPO è terminata. Selezionare il pulsante <CHIUDI File HTML> per completare il file <schede_bot_generi.html>"
End Sub

Private Sub CreaHtmlNomi3_Click()
Open FileTmp For Append As #4
    Print #4, "</div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "<br>"
    Print #4, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "32" & Chr(34) & ">"
    Print #4, "<tr height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<td height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
    Print #4, "<b><font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & ">A.M.I.N.T. - Associazione Micologica Italiana Naturalistica Telematica</font></b></div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "</center>"
    Print #4, "</body>"
    Print #4, "</html>"
Close #4
End Sub

Private Sub CreaHtmlSpecie1_Click()
FileOri = CurDir & "/Codicetestata-Specie.txt"
FileTmp = CurDir & "/schede_bot_specie.html"
FileCopy FileOri, FileTmp
Open FileTmp For Append As #1
Print #1, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "88" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<td height=" & Chr(34) & "75" & Chr(34) & ">"
Print #1, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & "><b>Indice delle Schede botaniche in ordine di Specie</b></font><br>"
Print #1, "<table width=" & Chr(34) & "40" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "10" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "40" & Chr(34) & " height=" & Chr(34) & "10" & Chr(34) & "></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & ">(a cura di Giuliano Salvai - realizzazione Web di Gianni Dose)</font><br>"
Print #1, "<hr width=" & Chr(34) & "400" & Chr(34) & ">"
Print #1, "<font color=" & Chr(34) & "white" & Chr(34) & "><b>" & Num_Schede & " Schede - aggiornamento " & LabelData.Caption & "</b></font></div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "<center>"
Print #1, "<table class=" & Chr(34) & "Table_Corpo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td>"
Print #1, "<div class=" & Chr(34) & "Testo3" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
Close #1
InfoSpecie.Caption = "ATTENZIONE: per proseguire con la creazione del File HTML premere il pulsante <CREA CORPO file HTML>, compare un Messaggio che prevede la conferma per ogni lettera presente nella Lista."
CreaHtmlSpecie1.Enabled = False
CreaHtmlSpecie2.Enabled = True
End Sub

Private Sub CreaHtmlSpecie2_Click()
' Fase 2 - crea corpo del file HTML
For i = 0 To ListaSpecie.ListCount - 1
ListaSpecie.ListIndex = i
LabelSpecie.Caption = ListaSpecie.List(i)
specie_cor = ListaSpecie.List(i)
specie1_cor = "'" & ListaSpecie.List(i) & "'"
AdoListaSpecie.RecordSource = "SELECT Tipo,Specie,Genere,Autore,nn,Com,Famiglia,NomeComune,AutoreScheda,Topic,New From Lista_Schede where Mid(Specie,1,1) = " & specie1_cor & " Order by Specie,Genere,nn"
AdoListaSpecie.Refresh
LabelNumSpecie.Caption = mshListaSpecie.Rows - 1
Dim Msg3, Style3, Titolo3, Risposta3
Msg3 = "Confermare la prosecuzione della procedura con le Specie che cominciano per  " & Chr(13) & Chr(13) & specie1_cor & " ?" & Chr(13) & Chr(13) & "Premere SI e attendere la comparsa del messaggio successivo."  ' Definisce il messaggio.
Style3 = vbYesNoCancel + vbCritical + vbDefaultButton2   ' Definisce i pulsanti.
Titolo3 = "Crea Corpo File HTML Specie"   ' Definisce il titolo.
      ' Visualizza il messaggio.
Risposta3 = MsgBox(Msg3, Style3, Titolo3)
If Risposta3 = vbYes Then   ' L'utente sceglie il pulsante Sì.
' Prima fase: parte iniziale
Open FileTmp For Append As #1
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "1" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "96" & Chr(34) & "><a href=" & Chr(34) & "#testata" & Chr(34) & "><img src=" & Chr(34) & "images/su2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "19" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></a></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & " width=" & Chr(34) & "770" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<font size=" & Chr(34) & "5" & Chr(34) & " color=" & Chr(34) & "#00008b" & Chr(34) & "><a name=" & Chr(34) & LabelSpecie.Caption & Chr(34) & "><b>" & LabelSpecie.Caption & "</b></a></font></div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#dfe6ef" & Chr(34) & "><div align=" & Chr(34) & "right" & Chr(34) & "></div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table class=" & Chr(34) & "testo2" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Print #1, "<tr height=" & Chr(34) & "26" & Chr(34) & ">"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & ">Scheda</td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Genere e specie</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Com.</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Famiglia</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Nome comune</div></td>"
Print #1, "<td bgcolor=" & Chr(34) & "#567DB4" & Chr(34) & " height=" & Chr(34) & "26" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & ">Autore</div></td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"
Print #1, "</td>"
Print #1, "</tr>"
Print #1, "</table>"
Print #1, "<table width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & ">"
Print #1, "<tr>"
Print #1, "<td width=" & Chr(34) & "970" & Chr(34) & ">"
Print #1, "<div align=" & Chr(34) & "center" & Chr(34) & ">"
Print #1, "<table class=" & Chr(34) & "Testo0" & Chr(34) & " width=" & Chr(34) & "970" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "2" & Chr(34) & " cellpadding=" & Chr(34) & "4" & Chr(34) & ">"
Close #1

Dim y As Integer
' Seconda fase: Corpo della Tabella
For y = 1 To mshListaSpecie.Rows - 1
' Seconda fase: Corpo della Tabella
mshListaSpecie.Row = y
mshListaSpecie.Col = 0 ' Tipo
TxTipo = mshListaSpecie.Text
mshListaSpecie.Col = 1 ' Specie
TxSpecie = mshListaSpecie.Text
mshListaSpecie.Col = 2 ' Genere
TxGenere = mshListaSpecie.Text
mshListaSpecie.Col = 3 ' Autore
TxAutore = mshListaSpecie.Text
mshListaSpecie.Col = 4 ' nn
TxNN = mshListaSpecie.Text
mshListaSpecie.Col = 5 ' Com
TxCom = mshListaSpecie.Text
mshListaSpecie.Col = 6 ' Famiglia
TxFamiglia = mshListaSpecie.Text
mshListaSpecie.Col = 7 ' NomeComune
TxNomeComune = mshListaSpecie.Text
mshListaSpecie.Col = 8 ' AutoreScheda
TxAutoreScheda = mshListaSpecie.Text
mshListaSpecie.Col = 9 ' Topic
TxTopic = mshListaSpecie.Text
mshListaSpecie.Col = 10 ' New
TxNew = mshListaSpecie.Text

' Riga1
Riga1 = "<tr height=" & Chr(34) & "21" & Chr(34) & ">"
' Riga2 = Tipo Scheda
If TxTipo = "+" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda1.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxTipo = "-" Then
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda2.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga2 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "48" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/scheda4.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & " width=" & Chr(34) & "25" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
End If
' Riga3 = Genere e specie
    If TxNew <> "" Then ' Riga3 = Genere e Specie (mette NEW se il campo New è attivo)
    '<td bgcolor="#DFE6EF" width="336" height="21"><a class="link2"href="http://forum.funghiitaliani.it/index.php?showtopic=32640" target="_blank">Abutilon theophrasti</a>  Medicus<br></td>
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxSpecie & ", " & TxGenere & "</a> " & TxAutore & "  <img src=" & Chr(34) & "images/new.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "14" & Chr(34) & " width=" & Chr(34) & "30" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></td>"
    Else
    Riga3 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "336" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><a class=" & Chr(34) & "link4" & Chr(34) & " href=" & Chr(34) & "http://www.funghiitaliani.it/index.php?show" & TxTopic & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">" & TxSpecie & ", " & TxGenere & "</a> " & TxAutore & "</td>"
    End If
' Riga4 = Com
If TxCom = "C" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_verde.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "CO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_CO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "V" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_rosso.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "VO" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_VO.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "31" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
ElseIf TxCom = "O" Then
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "><img src=" & Chr(34) & "images/bullet_off.gif" & Chr(34) & " alt=" & Chr(34) & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & "></div></td>"
Else
Riga4 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "46" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "center" & Chr(34) & "></div></td>"
End If
' Riga5 = Famiglia
'<td bgcolor="#DFE6EF" width="116" height="21">Malvaceae<br></td>
Riga5 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "116" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxFamiglia & "</td>"
' Riga6 = Nome Comune
'<td bgcolor="#DFE6EF" width="250" height="21">Cencio molle<br></td>
Riga6 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " width=" & Chr(34) & "250" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & ">" & TxNomeComune & "</td>"
' Riga8 = AutoreScheda
'<td bgcolor="#DFE6EF" height="21"><div align="left">Mirna e Attilio</div></td>
Riga7 = "<td bgcolor=" & Chr(34) & "#DFE6EF" & Chr(34) & " height=" & Chr(34) & "21" & Chr(34) & "><div align=" & Chr(34) & "left" & Chr(34) & ">" & TxAutoreScheda & "</div></td>"
Riga8 = "</tr>"

Open FileTmp For Append As #2
    Print #2, Riga1
    Print #2, Riga2
    Print #2, Riga3
    Print #2, Riga4
    Print #2, Riga5
    Print #2, Riga6
    Print #2, Riga7
    Print #2, Riga8
    'Print #2, Riga9
Close #2
Next y

Open FileTmp For Append As #3
    Print #3, "</table>"
    Print #3, "</div>"
    Print #3, "</td>"
    Print #3, "</tr>"
    Print #3, "</table>"
    Print #3, "<br>"
Close #3

ElseIf Risposta3 = vbNo Then

Else   ' L'utente sceglie il pulsante No.
   Exit For   ' Esegue un'azione.
End If
Next i
CreaHtmlSpecie2.Enabled = False
CreaHtmlSpecie3.Enabled = True
InfoSpecie.Caption = "ATTENZIONE: la procedura di creazione del CORPO è terminata. Selezionare il pulsante <CHIUDI File HTML> per completare il file <schede_bot_generi.html>"
End Sub

Private Sub CreaHtmlSpecie3_Click()
Open FileTmp For Append As #4
    Print #4, "</div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "<br>"
    Print #4, "<table class=" & Chr(34) & "Table_Titolo" & Chr(34) & " width=" & Chr(34) & "980" & Chr(34) & " height=" & Chr(34) & "32" & Chr(34) & ">"
    Print #4, "<tr height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<td height=" & Chr(34) & "28" & Chr(34) & ">"
    Print #4, "<div class=" & Chr(34) & "Testo0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">"
    Print #4, "<b><font size=" & Chr(34) & "3" & Chr(34) & " color=" & Chr(34) & "white" & Chr(34) & ">A.M.I.N.T. - Associazione Micologica Italiana Naturalistica Telematica</font></b></div>"
    Print #4, "</td>"
    Print #4, "</tr>"
    Print #4, "</table>"
    Print #4, "</center>"
    Print #4, "</body>"
    Print #4, "</html>"
Close #4

End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
'CentraSchermo
IndiceSchede.Left = (Screen.Width - IndiceSchede.Width) \ 2
IndiceSchede.Top = (Screen.Height - IndiceSchede.Height - 400) \ 2
DataAgg = Date
LabelData.Caption = DataAgg
' apre il file Schede_Forum.ini e predispone le impostazioni iniziali
'Open "Indice_Schede.ini" For Input As #1
'    Do While Not EOF(1)
'        Line Input #1, riga
'        ListaIni.AddItem riga
'    Loop
'Close #1

DbDati = "Lista_Schede_Flora.mdb"   ' Database ArchivioSchede.mdb

sorgente = DbDati
genere_cor = "'A'"
specie_cor = "'a'"
famiglia_cor = "'Acanthaceae'"
nomecomune_cor = "'A'"

'
    AdoListaSchede.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoListaSchedaSelezionata.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoListaFamiglie.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoListaGeneri.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoListaSpecie.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoListaNomiComuni.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoSchede.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    AdoSchedaSelezionata.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & sorgente & ";Persist Security Info=False"
    
' dati della tabella SchedeFlora
    AdoSchede.RecordSource = "SELECT Nome,Autore,Famiglia,NomeComune From ElencoSchede Order by Nome"
    AdoSchede.Refresh
    For i = 1 To AdoSchede.Recordset.RecordCount
    mshSchede.Col = 0
    mshSchede.Row = i
    ListaSchede.AddItem mshSchede.Text
    Next i
    mshSchede.ColWidth(0) = "2500"       'Nome
    mshSchede.ColWidth(1) = "1500"       'Autori
    mshSchede.ColWidth(2) = "1500"       'Famiglia
    mshSchede.ColWidth(3) = "4000"       'Nome comune
    'mshSchede.ColWidth(4) = "1500"       '

    
    AdoListaSchede.RecordSource = "SELECT * From Lista_Schede Order by Famiglia, Genere,Specie,nn"
    'AdoListaSchede.RecordSource = "SELECT * From Lista_Schede Order by Genere,Specie,nn"
    AdoListaSchede.Refresh
    AdoListaSchede.Recordset.MoveLast
    Num_Schede = AdoListaSchede.Recordset.RecordCount       'calcolo del numero dei records della tabella
    AdoListaSchede.Recordset.MoveFirst
    LabelNum_Schede.Caption = Num_Schede
    ' Crea Lista Famiglie
    FamigliaOld = ""
    For i = 1 To AdoListaSchede.Recordset.RecordCount
    'Dim Nome, gen, spe As String
    'mshListaSchede.Col = 2
    'mshListaSchede.Row = i
    'gen = mshListaSchede.Text
    'mshListaSchede.Col = 3
    'mshListaSchede.Row = i
    'spe = mshListaSchede.Text
    ' ListaSchede.AddItem gen & " " & spe
    mshListaSchede.Row = i
    mshListaSchede.Col = 6
    FamigliaNew = mshListaSchede.Text
    If FamigliaNew <> FamigliaOld Then
        ListaFamiglie.AddItem FamigliaNew
        FamigliaOld = FamigliaNew
    End If
    Next i
    
    ' Crea Lista Generi
    AdoListaSchede.RecordSource = "SELECT Genere,Specie,nn From Lista_Schede Order by Genere,Specie,nn"
    AdoListaSchede.Refresh
    GenereOld = ""
    For i = 1 To AdoListaSchede.Recordset.RecordCount
    mshListaSchede.Row = i
    mshListaSchede.Col = 0
    GenereNew = Mid(mshListaSchede.Text, 1, 1)
    If GenereNew <> GenereOld Then
        ListaGeneri.AddItem GenereNew
        GenereOld = GenereNew
    End If
    Next i
    
    ' Crea Lista Specie
    AdoListaSchede.RecordSource = "SELECT Specie,Genere,nn From Lista_Schede Order by Specie,Genere,nn"
    AdoListaSchede.Refresh
    SpecieOld = ""
    For i = 1 To AdoListaSchede.Recordset.RecordCount
    mshListaSchede.Row = i
    mshListaSchede.Col = 0
    SpecieNew = Mid(mshListaSchede.Text, 1, 1)
    If SpecieNew <> SpecieOld Then
        ListaSpecie.AddItem SpecieNew
        SpecieOld = SpecieNew
    End If
    Next i
    
    ' Crea Lista Nomi Comuni
    AdoListaSchede.RecordSource = "SELECT NomeComune,Genere,specie,nn From Lista_Schede Order by NomeComune,Genere,specie,nn"
    AdoListaSchede.Refresh
    NomeComuneOld = ""
    For i = 1 To AdoListaSchede.Recordset.RecordCount
    mshListaSchede.Row = i
    mshListaSchede.Col = 0
    NomeComuneNew = Mid(mshListaSchede.Text, 1, 1)
    If NomeComuneNew <> NomeComuneOld Then
        ListaNomiComuni.AddItem NomeComuneNew
        NomeComuneOld = NomeComuneNew
    End If
    Next i
    
    genere_cor = "'" & ListaGeneri.List(0) & "'"
    specie_cor = "'" & ListaSpecie.List(0) & "'"
    famiglia_cor = "'" & ListaFamiglie.List(0) & "'"
    nomecomune_cor = "'" & ListaNomiComuni.List(0) & "'"

    AdoListaFamiglie.RecordSource = "SELECT Tipo,AutoreScheda,Famiglia,Genere,Specie,Autore,nn,Com,NomeComune,Topic From Lista_Schede where Famiglia = " & famiglia_cor & " Order by Genere,Specie,nn"
    AdoListaFamiglie.Refresh
    
    AdoListaGeneri.RecordSource = "SELECT Tipo,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,AutoreScheda,Topic,New From Lista_Schede where Mid(Genere,1,1) = " & genere_cor & " Order by Genere,Specie,nn"
    AdoListaGeneri.Refresh
    mshListaGeneri.ColWidth(0) = "300"       'Tipo
    mshListaGeneri.ColWidth(1) = "1500"       'Genere
    mshListaGeneri.ColWidth(2) = "1500"       'Specie
    mshListaGeneri.ColWidth(3) = "1500"       'Autore
    mshListaGeneri.ColWidth(4) = "300"       'nn
    mshListaGeneri.ColWidth(5) = "300"       'Com
    mshListaGeneri.ColWidth(6) = "1200"       'Famiglia
    mshListaGeneri.ColWidth(7) = "1500"       'NomeComune
    mshListaGeneri.ColWidth(8) = "1000"       'AutoreScheda
    mshListaGeneri.ColWidth(9) = "1000"       'Topic
    mshListaGeneri.ColWidth(10) = "300"       'New
    LabelNumeroSpecie.Caption = mshListaGeneri.Rows - 1
    
    AdoListaSpecie.RecordSource = "SELECT Tipo,Specie,Genere,Autore,nn,Com,Famiglia,NomeComune,AutoreScheda,Topic,New From Lista_Schede where Mid(Specie,1,1) = " & specie_cor & " Order by Specie,Genere,nn"
    AdoListaSpecie.Refresh
    mshListaSpecie.ColWidth(0) = "300"       'Tipo
    mshListaSpecie.ColWidth(1) = "1500"       'Genere
    mshListaSpecie.ColWidth(2) = "1500"       'Specie
    mshListaSpecie.ColWidth(3) = "1500"       'Autore
    mshListaSpecie.ColWidth(4) = "300"       'nn
    mshListaSpecie.ColWidth(5) = "300"       'Com
    mshListaSpecie.ColWidth(6) = "1200"       'Famiglia
    mshListaSpecie.ColWidth(7) = "1500"       'NomeComune
    mshListaSpecie.ColWidth(8) = "1000"       'AutoreScheda
    mshListaSpecie.ColWidth(9) = "1000"       'Topic
    mshListaSpecie.ColWidth(10) = "300"       'New
    LabelNumSpecie.Caption = mshListaSpecie.Rows - 1
    
    AdoListaNomiComuni.RecordSource = "SELECT Tipo,AutoreScheda,NomeComune,Genere,Specie,Autore,nn,Com,Famiglia,Topic From Lista_Schede where Mid(NomeComune,1,1) = " & nomecomune_cor & " Order by NomeComune,Genere,Specie,nn"
    AdoListaNomiComuni.Refresh
    mshListaNomiComuni.ColWidth(0) = "300"       'Tipo
    mshListaNomiComuni.ColWidth(1) = "1500"       'NomeComune
    mshListaNomiComuni.ColWidth(2) = "1500"       'Genere
    mshListaNomiComuni.ColWidth(3) = "1500"       'Specie
    mshListaNomiComuni.ColWidth(4) = "1500"       'Autore
    mshListaNomiComuni.ColWidth(5) = "300"       'nn
    mshListaNomiComuni.ColWidth(6) = "300"       'Com
    mshListaNomiComuni.ColWidth(7) = "1200"       'Famiglia
    mshListaNomiComuni.ColWidth(8) = "1000"       'AutoreScheda
    mshListaNomiComuni.ColWidth(9) = "1000"       'Topic
    mshListaNomiComuni.ColWidth(10) = "300"       'New
    LabelNumeroNomi.Caption = mshListaSpecie.Rows - 1
    
    mshListaSchede.ColWidth(0) = "200"       'ID
    mshListaSchede.ColWidth(1) = "500"       'Autori
    mshListaSchede.ColWidth(2) = "1500"       '
    mshListaSchede.ColWidth(3) = "1500"       '
    mshListaSchede.ColWidth(4) = "1500"       '
    mshListaSchede.ColWidth(5) = "1500"       '
    mshListaSchede.ColWidth(6) = "1500"       '
    
    LabelFamiglia.Caption = ListaFamiglie.List(0)
    LabelGenere.Caption = ListaGeneri.List(0)
    LabelSpecie.Caption = ListaSpecie.List(0)
    LabelNomeComune.Caption = ListaNomiComuni.List(0)
    Num_Fam = ListaFamiglie.ListCount
    NumFamiglie.Caption = Num_Fam
    Fam1 = Int(Num_Fam / 6) + Num_Fam Mod 6
    'Label4.Caption = Num_Fam Mod 6
    'Label4.Caption = Fam1
    Fam2 = Fam1 * 2
    Fam3 = Fam1 * 3
    Fam4 = Fam1 * 4
    Fam5 = Fam1 * 5
    Fam6 = Fam1 * 6
    Label4.Caption = Fam1 & " " & Fam2 & " " & Fam3 & " " & Fam4 & " " & Fam5 & " " & Fam6
    InfoGeneri.Caption = "ATTENZIONE: per iniziare la procedura di creazione del File HTML premere il pulsante <CREA TESTATA file HTML>"
End Sub

Private Sub ListaFamiglie_Click()
LabelFamiglia.Caption = ListaFamiglie.List(ListaFamiglie.ListIndex)
famiglia_cor = "'" & ListaFamiglie.List(ListaFamiglie.ListIndex) & "'"
AdoListaFamiglie.RecordSource = "SELECT Tipo,AutoreScheda,Famiglia,Genere,Specie,Autore,nn,Com,NomeComune,Topic From Lista_Schede where Famiglia = " & famiglia_cor & " Order by Genere,Specie,nn"
AdoListaFamiglie.Refresh
End Sub

Private Sub ListaGeneri_Click()
LabelGenere.Caption = ListaGeneri.List(ListaGeneri.ListIndex)
genere_cor = "'" & ListaGeneri.List(ListaGeneri.ListIndex) & "'"
AdoListaGeneri.RecordSource = "SELECT Tipo,AutoreScheda,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,Topic From Lista_Schede where Mid(Genere,1,1) = " & genere_cor & " Order by Genere,Specie,nn"
AdoListaGeneri.Refresh
End Sub

Private Sub ListaNomiComuni_Click()
LabelNomeComune.Caption = ListaNomiComuni.List(ListaNomiComuni.ListIndex)
nomecomune_cor = "'" & ListaNomiComuni.List(ListaNomiComuni.ListIndex) & "'"
AdoListaNomiComuni.RecordSource = "SELECT Tipo,AutoreScheda,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,Topic From Lista_Schede where Mid(NomeComune,1,1) = " & nomecomune_cor & " Order by NomeComune,Genere,Specie,nn"
AdoListaNomiComuni.Refresh
End Sub

Private Sub ListaSchede_Click()
SchedaSelezionata = "'" & ListaSchede.List(ListaSchede.ListIndex) & "'"
    AdoSchedaSelezionata.RecordSource = "SELECT Nome,Autore,Famiglia,NomeComune From ElencoSchede where Nome = " & SchedaSelezionata & " Order by Nome"
    AdoSchedaSelezionata.Refresh
    mshSchedaSelezionata.ColWidth(0) = "2500"       'Nome
    mshSchedaSelezionata.ColWidth(1) = "1500"       'Autori
    mshSchedaSelezionata.ColWidth(2) = "1500"       'Famiglia
    mshSchedaSelezionata.ColWidth(3) = "4000"       'Nome comune
    Dim NomeSel, AutoreSel, FamigliaSel, NomeComuneSel As String
    NomeSel = AdoSchedaSelezionata.Recordset!Nome
    AutoreSel = AdoSchedaSelezionata.Recordset!Autore
    FamigliaSel = AdoSchedaSelezionata.Recordset!Famiglia
    NomeComuneSel = AdoSchedaSelezionata.Recordset!NomeComune
    TxTestoScheda.Text = "[color=#800080][size=4][b][i]" & NomeSel & "[/i][/b] " & AutoreSel & "[/size][/color]"
    TxTestoScheda.Text = TxTestoScheda.Text & Chr(13) & Chr(13) & "[b]Famiglia: " & FamigliaSel & "[/i][/b]"
    TxTestoScheda.Text = TxTestoScheda.Text & Chr(13) & Chr(13) & "[b]" & NomeComuneSel & "[/b]"
End Sub

Private Sub ListaSpecie_Click()
LabelSpecie.Caption = ListaSpecie.List(ListaSpecie.ListIndex)
specie_cor = "'" & ListaSpecie.List(ListaSpecie.ListIndex) & "'"
AdoListaSpecie.RecordSource = "SELECT Tipo,AutoreScheda,Genere,Specie,Autore,nn,Com,Famiglia,NomeComune,Topic From Lista_Schede where Mid(Specie,1,1) = " & specie_cor & " Order by Specie,Genere,nn"
AdoListaSpecie.Refresh
End Sub

