VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "意欣外贸管理系统"
   ClientHeight    =   6630
   ClientLeft      =   2595
   ClientTop       =   2880
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11820
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出系统"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   5760
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "lblTimer"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   5880
      Width           =   720
   End
   Begin VB.Menu mnuBusiness 
      Caption         =   " 日常业务 |"
      WindowList      =   -1  'True
      Begin VB.Menu mnuQuotation 
         Caption         =   "录入报价单"
      End
      Begin VB.Menu mnuEditQuotation 
         Caption         =   "编辑报价单"
      End
   End
   Begin VB.Menu mnuDocuments 
      Caption         =   "单据管理 |"
   End
   Begin VB.Menu mnuBasicInfo 
      Caption         =   "基础信息 |"
      Begin VB.Menu mnuAddCustomer 
         Caption         =   "添加客户"
      End
      Begin VB.Menu mnuModCustomer 
         Caption         =   "维护客户"
      End
      Begin VB.Menu seprator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddSupplier 
         Caption         =   "添加供应商"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "维护供应商"
      End
      Begin VB.Menu seprator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddContact 
         Caption         =   "添加联系人"
      End
      Begin VB.Menu mnuModContact 
         Caption         =   "维护联系人"
      End
   End
   Begin VB.Menu mnuBrowsePrint 
      Caption         =   "浏览打印"
   End
   Begin VB.Menu mnuSysConfig 
      Caption         =   "系统设置"
      Begin VB.Menu mnuUser 
         Caption         =   "用户管理"
      End
      Begin VB.Menu mnuUserConfig 
         Caption         =   "使用设置"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    frmManageUser.Show
    
End Sub

Private Sub mnuAddCustomer_Click()
    frmAddCustomer.Show
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuAddUser_Click()

End Sub

Private Sub mnuPrintOut_Click()

End Sub

Private Sub Timer1_Timer()
    lblTimer.Caption = Date
End Sub
