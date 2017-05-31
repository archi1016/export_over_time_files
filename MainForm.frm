VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   524
   ScaleMode       =   3  '像素
   ScaleWidth      =   776
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox DateDay 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5220
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   2520
      Width           =   1080
   End
   Begin VB.ComboBox DateMonth 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3960
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   2520
      Width           =   1080
   End
   Begin VB.ComboBox DateYear 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2400
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   2520
      Width           =   1380
   End
   Begin VB.CommandButton DoExport 
      Caption         =   "匯出 (&Q)"
      Height          =   540
      Left            =   9540
      TabIndex        =   11
      Top             =   7020
      Width           =   1800
   End
   Begin VB.CommandButton BrowserExportDir 
      Caption         =   "瀏覽 (&E)..."
      Height          =   540
      Left            =   9540
      TabIndex        =   9
      Top             =   3240
      Width           =   1800
   End
   Begin VB.TextBox ExportDirectory 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2400
      TabIndex        =   8
      Top             =   3240
      Width           =   6840
   End
   Begin VB.CommandButton BrowserSourceDir 
      Caption         =   "瀏覽 (&W)..."
      Height          =   540
      Left            =   9540
      TabIndex        =   2
      Top             =   1800
      Width           =   1800
   End
   Begin VB.TextBox SourceDirectory 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   6840
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  '平面
      BackColor       =   &H8000000D&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  '像素
      ScaleWidth      =   776
      TabIndex        =   10
      Top             =   0
      Width           =   11640
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "&3. 匯出資料夾:"
      Height          =   360
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   3360
      Width           =   1785
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "&2. 大於日期:"
      Height          =   360
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "&1. 來源資料夾:"
      Height          =   360
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   1920
      Width           =   1785
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = App.ProductName
    
    Call InitControl
End Sub

Private Sub InitControl()
    Dim I As Long
    Dim Y As Long
    
    Y = Year(Date)
    For I = 2014 To Y
        DateYear.AddItem CStr(I)
    Next
    For I = 1 To 12
        DateMonth.AddItem ConvLongToString(I, 2)
    Next
    For I = 1 To 31
        DateDay.AddItem ConvLongToString(I, 2)
    Next
End Sub

Private Sub StartListOverTimeFiles()
    Dim ST As SYSTEMTIME
    Dim FT As FILETIME
    Dim SF As String
    Dim EF As String
    Dim nQty As Long
    
    With ST
        .wYear = CInt(DateYear.List(DateYear.ListIndex))
        .wMonth = CInt(DateMonth.List(DateMonth.ListIndex))
        .wDay = CInt(DateDay.List(DateDay.ListIndex))
        .wMinute = 5
    End With
    SystemTimeToFileTime ST, FT
    
    SF = SourceDirectory.Text
    EF = ExportDirectory.Text
    
    DoExport.Enabled = False
    DoEvents
    
    nQty = 0
    Call ListOverTimeFiles(SF, "", FT, EF, nQty)
    
    DoExport.Enabled = True
    DoEvents
End Sub

Private Sub ListOverTimeFiles(SrcFP As String, ByVal SubFP As String, FT As FILETIME, ExpFP As String, nQty As Long)
    Dim FN As String
    Dim hFind As Long
    Dim WFD As WIN32_FIND_DATA
    Dim nAttr As Long
    Dim I As Long
    Dim EF As String

    FN = SrcFP + SubFP + "\*"
    hFind = FindFirstFileW(StrPtr(FN), WFD)
    If INVALID_HANDLE_VALUE <> hFind Then
        I = 1
        Do
            FN = String$(MAX_PATH, vbNullChar)
            CopyMemory StrPtr(FN), VarPtr(WFD.cFileName(0)), MAX_PATH * 2
            FN = StrCutNull(FN)
            
            nAttr = WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
            If 0 = nAttr Then
                If -1 = CompareFileTime(FT, WFD.ftLastWriteTime) Then
                    EF = ExpFP + SubFP + "\" + FN
                    Call CheckAndCreateFolder(ReturnParentDirectory(EF))
                    CopyFile SrcFP + SubFP + "\" + FN, EF
                    
                    nQty = nQty + 1
                End If
            Else
                If "." <> FN Then
                    If ".." <> FN Then
                        Call ListOverTimeFiles(SrcFP, SubFP + "\" + FN, FT, ExpFP, nQty)
                    End If
                End If
            End If
            
            I = I + 1
            If 0 = (I And &H3F) Then DoEvents
            
        Loop Until (0 = FindNextFileW(hFind, WFD))
        FindClose hFind
    End If
End Sub

Private Sub BrowserExportDir_Click()
    Dim FP As String
    
    FP = GetSelectedFolder(Me.hWnd, "請選取匯出資料夾")
    If "" <> FP Then
        If "\" <> Right$(FP, 1) Then
            ExportDirectory.Text = FP
        Else
            Call MsgError("不支援根目錄！")
        End If
    End If
End Sub

Private Sub BrowserSourceDir_Click()
    Dim FP As String
    
    FP = GetSelectedFolder(Me.hWnd, "請選取來源資料夾")
    If "" <> FP Then
        If "\" <> Right$(FP, 1) Then
            SourceDirectory.Text = FP
        Else
            Call MsgError("不支援根目錄！")
        End If
    End If
End Sub

Private Sub DoExport_Click()
    If "" <> SourceDirectory.Text Then
        If -1 <> DateYear.ListIndex Then
            If -1 <> DateMonth.ListIndex Then
                If -1 <> DateDay.ListIndex Then
                    If "" <> ExportDirectory.Text Then
                        Call StartListOverTimeFiles
                    Else
                        Call MsgError("未指定匯出資料夾！")
                    End If
                Else
                    Call MsgError("未指定大於日期！")
                End If
            Else
                Call MsgError("未指定大於日期！")
            End If
        Else
            Call MsgError("未指定大於日期！")
        End If
    Else
        Call MsgError("未指定來源資料夾！")
    End If
End Sub
