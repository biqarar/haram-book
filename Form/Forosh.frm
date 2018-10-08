VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forosh 
   BackColor       =   &H00404000&
   Caption         =   "À»  ›—Ê‘ ò«·«"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Forosh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
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
      Height          =   255
      Left            =   0
      TabIndex        =   12
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataSource      =   "l"
         Height          =   465
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   465
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "À» "
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
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".»Â ›—Ê‘ —”Ìœ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   4920
         TabIndex        =   8
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "œ—  «—ÌŒ"
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   6360
      Y1              =   3600
      Y2              =   3600
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
      TabIndex        =   16
      Top             =   3720
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
      TabIndex        =   15
      Top             =   3720
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
      TabIndex        =   14
      Top             =   4320
      Width           =   3375
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
      TabIndex        =   13
      Top             =   4320
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
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
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
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " ⁄œ«œ »Â ›—Ê‘ —”ÌœÂ «„—Ê“"
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
      TabIndex        =   0
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "Forosh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tedad_ As Double
Dim Form_load_ As Integer
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
P_forosh.RecordSource = "select * from p_forosh where id_kl like ('" & id_kl_ & "') and xdate like ('" & Text8.Text & "')"
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

Private Sub Combo1_Click()
On Error Resume Next

a = Split(Combo1.Text, " _ ")
If Form_load_ = 0 Then Mojodi_feli (a(0))

End Sub


Private Sub Form_Load()
TarikhShamsi
Text8.Text = Tarikh
Refresh_combo1
Call Combo1_Click

End Sub
Function Refresh_combo1()
On Error Resume Next

Form_load_ = 1

P_kl.Refresh
P_kl.RecordSource = "select * from p_kl where vazeyat like ('1')"
P_kl.Refresh
Combo1.Clear

For I = 1 To P_kl.Recordset.RecordCount
Combo1.AddItem (P_kl.Recordset.Fields("id_kl") & " _ " & P_kl.Recordset.Fields("xname") & "    ﬁÌ„     " & P_kl.Recordset.Fields("qeymat") & "   «—ÌŒ Œ—Ìœ " & P_kl.Recordset.Fields("xdate"))
Combo1.Text = Combo1.List(0)
P_kl.Recordset.MoveNext
Next I
Form_load_ = 0
End Function
Private Sub Label7_Click()
On Error Resume Next

If Text8.Text = "" Or Text1.Text = "" Or Val(Text1.Text) <= 0 Or Len(Text8.Text) <> 8 Or Combo1.Text = "" Then Exit Sub
If Val(Label11.Caption) < Val(Text1.Text) Then
MsgBox "„ﬁœ«— Ê«—œ ‘œÂ »Ì‘ — «“ „ÊÃÊœÌ ›⁄·Ì «” ", vbCritical, "Œÿ«"
Exit Sub
End If

a = Split(Combo1.Text, " _ ")
P_forosh.Refresh
P_forosh.RecordSource = "select * from p_forosh where id_kl like ('" & a(0) & "') and tedad like ('" & Text1.Text & "') and xdate like ('" & Text8.Text & "')"
P_forosh.Refresh

If P_forosh.Recordset.EOF = False Or P_forosh.Recordset.BOF = False Then
If MsgBox("ç‰œ „Ê—œ „‘«»Â Â„Ì‰ «ÿ·«⁄«  Ì«›  ‘œÂ «”  ¬Ì« „ÿ„∆‰ Â” Ìœ òÂ «ÿ·«⁄«   ò—«—Ì ‰Ì” ", vbExclamation + vbYesNo, "Â‘œ«—") = vbYes Then
GoTo 1
Else
Exit Sub
End If
End If
1

P_forosh.Refresh
P_forosh.Recordset.AddNew
P_forosh.Recordset.Fields("id_kl") = a(0)
P_forosh.Recordset.Fields("tedad") = Text1.Text
P_forosh.Recordset.Fields("xdate") = Text8.Text
P_forosh.Recordset.Update
P_forosh.Refresh
Mojodi_feli (a(0))
Refresh_combo1
MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation, ".::."

End Sub

