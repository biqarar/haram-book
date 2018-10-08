VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form show_book 
   BackColor       =   &H00404000&
   Caption         =   "‰„«Ì‘ ò«·«Â«Ì „ÊÃÊœ«’·«Õ «ÿ·«⁄« "
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "show_book.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   465
      Left            =   11280
      TabIndex        =   13
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Frame Motor 
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      Begin MSAdodcLib.Adodc P_kl 
         Height          =   375
         Left            =   600
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from P_kl"
         Caption         =   "P_kl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc P_forosh 
         Height          =   375
         Left            =   600
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from P_forosh"
         Caption         =   "P_forosh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc P_kl2 
         Height          =   375
         Left            =   600
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Developer\haram\Ketab_Paziresh\Data_Book_P.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from P_kl"
         Caption         =   "P_kl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "show_book.frx":08CA
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   4210688
      ForeColor       =   65535
      HeadLines       =   1
      RowHeight       =   34
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ò«·«Â«Ì „ÊÃÊœ"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ID_Kl"
         Caption         =   "òœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "xname"
         Caption         =   "‰«„ ò«·«"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "qeymat"
         Caption         =   "ﬁÌ„ "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "takhfif"
         Caption         =   " Œ›Ì›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "tedad"
         Caption         =   " ⁄œ«œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "factor"
         Caption         =   "‘„«—Â ›«ò Ê—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "xdate"
         Caption         =   " «—ÌŒ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "user"
         Caption         =   "ò«—»—"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "tozih"
         Caption         =   " Ê÷ÌÕ« "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Õ–› ò«·«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   19
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "À»  ›—Ê‘ «Ì‰ ò«·«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "«’·«Õ «ÿ·«⁄« "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ê÷⁄Ì  ›«ò Ê—"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "›⁄«·"
      DataField       =   "vazeyat"
      DataSource      =   "P_kl2"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   13380
      TabIndex        =   14
      Top             =   5640
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "ID_Kl"
      DataSource      =   "P_kl2"
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " ⁄œ«œ »Â ›—Ê‘ —”ÌœÂ œ—  «—ÌŒ ›Êﬁ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "„»·€ œ—Ì«› Ì «„—Ê“"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   10
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "«ÿ·«⁄«  œ— œ” —” ‰Ì” "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "«ÿ·«⁄«  œ— œ” —” ‰Ì” "
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1065
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "«ÿ·«⁄«  œ— œ” —” ‰Ì” "
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1065
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "„»·€ ò· œ—Ì«› Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "«ÿ·«⁄«  œ— œ” —” ‰Ì” "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "„ÊÃÊœÌ ›⁄·Ì"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã” ÃÊ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   13200
      TabIndex        =   1
      Top             =   360
      Width           =   525
   End
End
Attribute VB_Name = "show_book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Change()
Mojodi_feli (P_kl.Recordset.Fields("id_kl"))

End Sub

Function Mojodi_feli(id_kl_)
On Error Resume Next

P_forosh.Refresh
P_forosh.RecordSource = "select * from p_forosh where id_kl like ('" & id_kl_ & "')" ' and xdate like ('" & Text8.Text & "')"
P_forosh.Refresh
Tedad_ = 0
For I = 1 To P_forosh.Recordset.RecordCount
Tedad_ = Tedad_ + Val(P_forosh.Recordset.Fields("tedad"))
P_forosh.Recordset.MoveNext
Next I
P_kl.Refresh
P_kl.RecordSource = "select * from p_kl where id_kl = " & id_kl_
P_kl.Refresh

Label11.Caption = Val(P_kl.Recordset.Fields("tedad")) - Tedad_
If Val(Label11.Caption) = 0 Then
P_kl.Recordset.Fields("vazeyat") = 0
P_kl.Recordset.Update
End If

Label8.Caption = Val(P_kl.Recordset.Fields("qeymat")) * Tedad_


P_forosh.Refresh
P_forosh.RecordSource = "select * from p_forosh where id_kl like ('" & id_kl_ & "') and xdate like ('%" & Text8.Text & "%')"
P_forosh.Refresh
Tedad_ = 0
For I = 1 To P_forosh.Recordset.RecordCount
Tedad_ = Tedad_ + Val(P_forosh.Recordset.Fields("tedad"))
P_forosh.Recordset.MoveNext
Next I
P_kl.Refresh
P_kl.RecordSource = "select * from p_kl where id_kl = " & id_kl_
P_kl.Refresh

Label5.Caption = Tedad_

Label6.Caption = Val(P_kl.Recordset.Fields("qeymat")) * Tedad_

End Function

Private Sub Label10_Change()
If Label10.Caption = "1" Then
Label10.BackColor = &HC000&
Label10.Caption = "›⁄«·"
ElseIf Label10.Caption = "0" Then
Label10.BackColor = &H808080
Label10.Caption = "€Ì— ›⁄«·"
End If

End Sub

Private Sub Label14_Click()
Add_kl.Show

Add_kl.P_kl.Refresh
Add_kl.P_kl.RecordSource = "select * from p_kl where id_kl = " & Me.P_kl2.Recordset.Fields("id_kl")
Add_kl.P_kl.Refresh
Add_kl.Text5.Text = "update"

End Sub

Private Sub Label15_Click()
Forosh.Show
For I = 0 To Forosh.Combo1.ListCount - 1
a = Split(Forosh.Combo1.List(I), " _ ")
If Me.P_kl2.Recordset.Fields("id_kl") = Val(a(0)) Then
Forosh.Combo1.Text = Forosh.Combo1.List(I)
End If
Next I


End Sub

Private Sub Label16_Click()
On Error Resume Next
If MsgBox("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ò«·« —« Õ–› ò‰Ìœ", vbQuestion + vbYesNo, "Õ–› ò«·«") = vbYes Then


P_forosh.Refresh
P_forosh.RecordSource = "select * from p_forosh where id_kl like ('" & Me.P_kl2.Recordset.Fields("id_kl") & "')"
P_forosh.Refresh
If P_forosh.Recordset.EOF = False Or P_forosh.Recordset.BOF = False Then
MsgBox "„‘Œ’«  «Ì‰ ò«·« œ— Õ«· «” ›«œÂ «” " & Chr$(10) & "«„ò«‰ Õ–› «Ì‰ ò«·« ÊÃÊœ ‰œ«—œ", vbCritical, "Œÿ«"
Exit Sub
End If

P_kl2.Recordset.Delete
End If

'P_kl2.RecordSource = "delete from p_kl where id_kl = " & Me.P_kl2.Recordset.Fields("id_kl")


End Sub

Private Sub Label2_Change()
Mojodi_feli (Label2.Caption)

End Sub

Private Sub Text1_Change()
On Error Resume Next

t = Text1.Text
P_kl2.Refresh
P_kl2.RecordSource = "select * from P_kl where user like ('%" & t & "%') or xdate like ('%" & t & "%') or factor like ('%" & t & "%') or xname like ('%" & t & "%') or tozih like ('%" & t & "%') or qeymat like('%" & t & "%') or tedad like ('%" & t & "%') "
P_kl2.Refresh

End Sub
