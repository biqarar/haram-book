VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Add_kl 
   BackColor       =   &H00404000&
   Caption         =   "«÷«›Â ò—œ‰ ò«·«"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_kl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Text            =   "add_new"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "user"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "xdate"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "ID_Kl"
      DataSource      =   "P_kl"
      Enabled         =   0   'False
      Height          =   465
      Left            =   3480
      TabIndex        =   16
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "qeymat"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "tedad"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "factor"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "xname"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "tozih"
      DataSource      =   "P_kl"
      Height          =   465
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   5295
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
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   0
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
         Left            =   720
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
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "»«“ê‘ "
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ÕÊÌ· êÌ—‰œÂ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2400
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Œ—Ìœ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5610
      TabIndex        =   18
      Top             =   2760
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "òœ ò«·«"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5760
      TabIndex        =   17
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ò«·«Ì ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "–ŒÌ—Â"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ﬁÌ„  "
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5805
      TabIndex        =   15
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5820
      TabIndex        =   14
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ›«ò Ê—"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2400
      TabIndex        =   13
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Ê÷ÌÕ« "
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5685
      TabIndex        =   12
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«·«"
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   5760
      TabIndex        =   11
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "Add_kl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tarikh As String

Function TarikhShamsi(Optional date1 As String, Optional SmallDate1 As Boolean) As String
On Error Resume Next

      '====================================================
      Dim d, p, w, mon, mm, Ym, u, v, rp, X, I, Ys, Ms, Dm, P1, D1, Ds, DateShamsi
      d = Array(20, 19, 20, 20, 21, 21, 22, 22, 22, 22, 21, 21)
      p = Array(11, 12, 10, 12, 11, 11, 10, 10, 10, 9, 10, 10)
      w = Array("Ìò‘‰»Â", "œÊ‘‰»Â", "”Â ‘‰»Â", "çÂ«—‘‰»Â", "Å‰Ã‘‰»Â", "Ã„⁄Â", "‘‰»Â")
      
      If SmallDate1 = True Then
            mon = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
      Else
            mon = Array("›—Ê—œÌ‰", "«—œÌ»Â‘ ", "Œ—œ«œ", " Ì—", "„—œ«œ", "‘Â—ÌÊ—", "„Â—", "¬»«‰", "¬–—", "œÌ", "»Â„‰", "«”›‰œ")
      End If
      
      If date1 = "" Then date1 = Date
      
      Dm = Day(date1) '»œ”  ¬Ê—œ‰ —Ê“
      mm = Month(date1) '»œ”  ¬Ê—œ‰ „«Â
      Ym = Year(date1) '»œ”  ¬Ê—œ‰ ”«·
      u = 0
      rp = 0
      If (Ym Mod 4) = 0 Then u = 1 ' ‘ŒÌ’ ò»Ì”Â »Êœ‰
      If ((Ym Mod 100) = 0 And (Ym Mod 400) <> 0) Then u = 0 ' ‘ŒÌ’ ò»Ì”Â ‰»Êœ‰
      Ys = Ym - 622 ' »œÌ· ”«· „Ì·«œÌ »Â ‘„”Ì
      X = Ys - 22
      X = X Mod 33
      If ((X Mod 4) = 0 And X <> 32) Then rp = 1
      I = Not (rp - 2) + Not (u - 2) * 2
      X = 0
      If (I = 0 And mm = 3) Then X = 1
      If I = 0 Then I = 3
      Ms = (9 + mm) Mod 13
      If Ms < 10 Then Ms = Ms + 1
      D1 = d(mm - 1)
      If (I = 1 And mm > 2) Then D1 = D1 - 1
      If (I = 2 And mm < 3) Then D1 = D1 - 1
      P1 = p(mm - 1)
      If (I = 1 And mm > 2) Then P1 = P1 + 1
      If (I = 2 And mm < 4) Then P1 = P1 + 1
      If (Dm > 0 And Dm <= D1) Then
             Ds = P1 + Dm + X - 1
          X = 1
      Else
          Ds = Dm - D1
          Ms = Ms + 1
          If Ms = 13 Then Ms = 1
          X = 2
      End If
      If ((mm = 3 And X = 2) Or mm > 3) Then Ys = Ys + 1
      If SmallDate1 = True Then
'     ??? ??? ?? ???? ???? ???????? ???????? ?? ??? ?? ?? ???? ????? ?? ?????
'            TarikhShamsi = Trim(Str(Ys)) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
            TarikhShamsi = Mid(Trim(Str(Ys)), 3, 2) + "/" + Trim(mon(Ms - 1)) + "/" + Trim(Str(Ds))
           If Val(Ys) < 10 Then Ys = "0" & Val(Ys)
           If Val(Ms) < 10 Then Ms = "0" & Val(Ms)
           If Val(Ds) < 10 Then Ds = "0" & Val(Ds)
            
            Tarikh = Val(Ys) & "/" & (Ms) & "/" & Val(Ds)
            ' Tarikh.Caption = Ys & Ms & Ds
      Else
            TarikhShamsi = w(Weekday(Date) - 1) + " " + Str(Ds) + " " + mon(Ms - 1) + " " + Str(Ys)
           If Val(Ys) < 10 Then Ys = "0" & Val(Ys)
           If Val(Ms) < 10 Then Ms = "0" & Val(Ms)
           If Val(Ds) < 10 Then Ds = "0" & Val(Ds)
            
            Tarikh = Val(Ys) & "/" & (Ms) & "/" & Val(Ds)
            'Tarikh.Caption = Ys & Ms & Ds
            Tarikh = Ys & Ms & Ds
      End If

End Function
Private Sub Form_Load()
On Error Resume Next

P_forosh.Refresh
P_kl.Refresh
TarikhShamsi
Call Label8_Click

Text8.Text = Tarikh

End Sub

Private Sub Label12_Click()
Unload Me

End Sub

Private Sub Label7_Click()
On Error Resume Next
If Text2.Text = "" Or Text4.Text = "" Or Text6.Text = "" Or Len(Text8.Text) <> 8 Then Exit Sub
If Text5.Text = "add_new" Then

P_kl2.Refresh
P_kl2.RecordSource = "select * from p_kl where xname like ('" & Text2.Text & "') and qeymat like ('" & Text6.Text & "')  and xdate like ('" & Text8.Text & "') and tedad like ('" & Text4.Text & "') and factor like ('" & Text3.Text & "')"
P_kl2.Refresh
If P_kl2.Recordset.EOF = False Or P_kl2.Recordset.BOF = False Then
MsgBox "«Ì‰ „‘Œ’«  ﬁ»·« À»  ‘œÂ «” ", vbInformation, "Â‘œ«—"
Exit Sub
End If

P_kl.Recordset.Fields("vazeyat") = 1

P_kl.Recordset.Update
MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation, ".::."
ElseIf Text5.Text = "update" Then

P_kl.Recordset.Update
MsgBox " €ÌÌ—«  «⁄„«· ‘œ", vbInformation, ".::."

End If

End Sub

Private Sub Label8_Click()
On Error Resume Next
Text5.Text = "add_new"

P_kl.Refresh
P_kl.Recordset.AddNew

End Sub

