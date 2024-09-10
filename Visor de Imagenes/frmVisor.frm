VERSION 5.00
Begin VB.Form frmVisor 
   BackColor       =   &H000080FF&
   Caption         =   "Visor de Imagenes"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImagen 
      Caption         =   "&Ver Imagen"
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   5400
      Width           =   3135
   End
   Begin VB.DirListBox dirCarpetas 
      Height          =   2340
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.FileListBox filArchivos 
      Height          =   4185
      Left            =   4800
      Pattern         =   "*.jpg; *.bmp;*.wmf; *.ico; *.gif"
      TabIndex        =   7
      Top             =   720
      Width           =   3135
   End
   Begin VB.DriveListBox drvUnidad 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4335
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblArchivo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del archivo:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label lblSeleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblDir 
      Caption         =   "Directorio:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblUnidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblRuta 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblDireccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImagen_Click()
Rem Al darle click en la imagen
Rem Guarda el archivo seleccionado en la variable SelectedFile
SelectedFile = filArchivos.Path & "\" & filArchivos.FileName

Rem Lo muestra en el contenedo de Imagen con la propiedad Picture
Rem y el metodo LoadPicture()
Image1.Picture = LoadPicture(SelectedFile)
End Sub

Private Sub dirCarpetas_Change()
Rem Al cambiar de directorio:
Rem 1 - Muestra la ruta sleccionada en la etiqueta
lblRuta.Caption = dirCarpetas.Path

Rem 2 - Actualiza ambas rutas la de carpetas con la de archivos
filArchivos.Path = dirCarpetas.Path
End Sub

Private Sub drvUnidad_Change()
Rem Al cambiar de unidad de disco
'Áctualiza los directorios o carpetas
dirCarpetas.Path = drvUnidad.Drive
End Sub

Private Sub filArchivos_Click()
Rem Al dar Click en un archivo
Rem Escribe el nombre del mismo en la caja de texto
txtArchivo.Text = filArchivos.FileName
End Sub

