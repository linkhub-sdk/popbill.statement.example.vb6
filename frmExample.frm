VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 전자명세서 SDK 예제"
   ClientHeight    =   10860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   11910
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnDetachStatement 
      Caption         =   "전자명세서 첨부해제"
      Height          =   375
      Left            =   3000
      TabIndex        =   68
      Top             =   9960
      Width           =   2115
   End
   Begin VB.Frame Frame7 
      Caption         =   " 전자명세서 관련 기능 "
      Height          =   7380
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   11475
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스"
         Height          =   2655
         Left            =   5040
         TabIndex        =   61
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "삭제"
            Height          =   495
            Left            =   1560
            Style           =   1  '그래픽
            TabIndex        =   64
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton btnCancelISsue_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   480
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   63
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   405
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   62
            Top             =   480
            Width           =   1020
         End
         Begin VB.Line Line4 
            X1              =   1080
            X2              =   1965
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   240
            Top             =   360
            Width           =   2040
         End
         Begin VB.Line Line5 
            X1              =   840
            X2              =   840
            Y1              =   1680
            Y2              =   600
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " 문서 정보 "
         Height          =   2535
         Left            =   240
         TabIndex        =   47
         Top             =   4200
         Width           =   2010
         Begin VB.CommandButton btnSearch 
            Caption         =   "문서 목록조회"
            Height          =   375
            Left            =   210
            TabIndex        =   66
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "문서 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   51
            Top             =   270
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "문서 정보(대량)"
            Height          =   390
            Left            =   210
            TabIndex        =   50
            Top             =   705
            Width           =   1590
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "문서 이력"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   1140
            Width           =   1590
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "문서 상세 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   48
            Top             =   1590
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " 부가 서비스"
         Height          =   3015
         Left            =   2520
         TabIndex        =   43
         Top             =   4200
         Width           =   2580
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "전자명세서 첨부"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   2070
            Width           =   2115
         End
         Begin VB.CommandButton btnFAXSEnd 
            Caption         =   "선팩스 전송"
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   1640
            Width           =   2115
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   46
            Top             =   1200
            Width           =   2115
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   45
            Top             =   735
            Width           =   2115
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   44
            Top             =   300
            Width           =   2115
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   1290
         Left            =   8880
         TabIndex        =   40
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "매출 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   42
            Top             =   705
            Width           =   1500
         End
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   41
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " 문서 정보 "
         Height          =   2565
         Left            =   5400
         TabIndex        =   34
         Top             =   4200
         Width           =   3210
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "이메일(공급받는자) 링크 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   39
            Top             =   1590
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "다량 인쇄 팝업 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   38
            Top             =   1140
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "문서 내용 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintURL 
            Caption         =   "수신자 인쇄 팝업 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   35
            Top             =   2040
            Width           =   2745
         End
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   2730
         TabIndex        =   32
         Top             =   1350
         Width           =   2025
      End
      Begin VB.TextBox txtFormCode 
         Height          =   345
         Left            =   2730
         TabIndex        =   31
         Top             =   945
         Width           =   2025
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9480
         Top             =   6600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame10 
         Caption         =   " 첨부파일 "
         Height          =   1335
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   4560
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "파일 삭제"
            Height          =   390
            Left            =   3120
            TabIndex        =   21
            Top             =   840
            Width           =   1245
         End
         Begin VB.TextBox txtFileID 
            Height          =   330
            Left            =   240
            TabIndex        =   20
            Text            =   "파일아이디"
            Top             =   840
            Width           =   2820
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "첨부 목록"
            Height          =   390
            Left            =   1800
            TabIndex        =   19
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "파일 첨부"
            Height          =   390
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "임시저장- 발행 프로세스"
         Height          =   2700
         Left            =   7800
         TabIndex        =   12
         Top             =   480
         Width           =   3510
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행"
            Height          =   525
            Left            =   322
            Style           =   1  '그래픽
            TabIndex        =   30
            Top             =   1365
            Width           =   1020
         End
         Begin VB.CommandButton btnCancel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   480
            Left            =   285
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   2055
            Width           =   1095
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "등록"
            Height          =   375
            Left            =   1305
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "수정"
            Height          =   375
            Left            =   2265
            Style           =   1  '그래픽
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "삭제"
            Height          =   495
            Left            =   2355
            Style           =   1  '그래픽
            TabIndex        =   13
            Top             =   2040
            Width           =   855
         End
         Begin VB.Line Line2 
            X1              =   840
            X2              =   2850
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임시저장"
            Height          =   180
            Left            =   465
            TabIndex        =   16
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            FillColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   300
            Top             =   345
            Width           =   3000
         End
         Begin VB.Line Line3 
            X1              =   2835
            X2              =   2835
            Y1              =   2025
            Y2              =   825
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   840
            Y1              =   2040
            Y2              =   720
         End
      End
      Begin VB.ComboBox cboItemCode 
         Height          =   300
         ItemData        =   "frmExample.frx":0000
         Left            =   2730
         List            =   "frmExample.frx":0016
         TabIndex        =   11
         Text            =   "거래명세서"
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "관리번호 사용여부 확인"
         Height          =   375
         Left            =   2565
         TabIndex        =   10
         Top             =   1830
         Width           =   2190
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "명세서 종류 : "
         Height          =   180
         Left            =   1530
         TabIndex        =   33
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "양식코드( FormCode ) : "
         Height          =   180
         Left            =   615
         TabIndex        =   22
         Top             =   1050
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "관리번호( MgtKey) : "
         Height          =   180
         Left            =   915
         TabIndex        =   9
         Top             =   1455
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2610
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   11475
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련 "
         Height          =   1815
         Left            =   9240
         TabIndex        =   58
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   1770
         Left            =   6840
         TabIndex        =   27
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   "팝빌 로그인 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   1770
         Left            =   4680
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   2250
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   2625
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   360
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   2265
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   255
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   1770
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   53
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   255
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'링크아이디
Private Const LinkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private statementService As New PBDocService

Private Function selectedItemCode() As Integer
    selectedItemCode = 121
 
    If cboItemCode.Text = "거래명세서" Then selectedItemCode = 121
    If cboItemCode.Text = "청구서" Then selectedItemCode = 122
    If cboItemCode.Text = "견적서" Then selectedItemCode = 123
    If cboItemCode.Text = "발주서" Then selectedItemCode = 124
    If cboItemCode.Text = "입금표" Then selectedItemCode = 125
    If cboItemCode.Text = "영수증" Then selectedItemCode = 126
    
End Function
Private Sub btn_GetURL_PBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "PBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub



Private Sub btnAttachFile_Click()
    Dim FilePath As String
    CommonDialog1.FileName = ""
    
    CommonDialog1.ShowOpen
    
    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
    
    Dim Response As PBResponse
  
    
    Set Response = statementService.AttachFile(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, FilePath, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub
Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    SubItemCode = 121           '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubMgtKey = "20151223-01"   '첨부할 전자명세서 관리번호
      
    Set Response = statementService.AttachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCancel_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnCancelISsue_2_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "발행 취소 메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
   
    
    Set Response = statementService.CheckMgtKeyInUse(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub

Private Sub btnDelete_2_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnDelete_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnDeleteFile_Click()
    Dim Response As PBResponse
   
    
    Set Response = statementService.DeleteFile(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtFileID.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    SubItemCode = 121           '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubMgtKey = "20151223-01"   '첨부해제할 전자명세서 관리번호
      
    Set Response = statementService.DetachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnFAXSEnd_Click()

    Dim Statement As New PBStatement
    
    Statement.sendNum = "07075103710"            '팩스전송 발신번호
    Statement.receiveNum = "070111222"           '팩스전송 수신번호
       
    Statement.writeDate = "20151007"             '필수, 기재상 작성일자
    Statement.purposeType = "영수"               '필수, {영수, 청구}
    Statement.taxType = "과세"                   '필수, {과세, 영세, 면세}
    Statement.formCode = txtFormCode.Text
    
    Statement.itemCode = selectedItemCode
    
    Statement.mgtKey = txtMgtKey.Text            '팩스전송 파일명
    
    Statement.senderCorpNum = txtCorpNum.Text
    Statement.senderTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Statement.senderCorpName = "공급자 상호"
    Statement.senderCEOName = "공급자"" 대표자 성명"
    Statement.senderAddr = "공급자 주소"
    Statement.senderBizClass = "공급자 업종"
    Statement.senderBizType = "공급자 업태,업태2"
    Statement.senderContactName = "공급자 담당자명"
    Statement.senderEmail = "test@test.com"
    Statement.senderTEL = "070-7070-0707"
    Statement.senderHP = "010-000-2222"
    
    Statement.receiverCorpNum = "8888888888"
    Statement.receiverCorpName = "공급받는자 상호"
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    Statement.receiverAddr = "공급받는자 주소"
    Statement.receiverBizClass = "공급받는자 업종"
    Statement.receiverBizType = "공급받는자 업태"
    Statement.receiverContactName = "공급받는자 담당자명"
    Statement.receiverEmail = "test@receiver.com"
    
    Statement.supplyCostTotal = "100000"         '필수 공급가액 합계
    Statement.taxTotal = "10000"                 '필수 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Statement.serialNum = "123"
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    Statement.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Statement.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    Statement.faxsendYN = False          '발행시 Fax발송시 설정.
    Statement.smssendYN = False '발행시 문자발송기능 사용시 활용
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    Dim i
    Dim newDetail As PBDocDetail
    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
        newDetail.itemName = "품명" + CStr(i)
        newDetail.spec = "규격"
        newDetail.unit = "단위"
        newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "비고"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '추가속성, [참고] 전자명세서 기본양식 추가속성 테이블 참조 http://blog.linkhub.co.kr/2514/
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Dim ReceiptNum As String
    
    ReceiptNum = statementService.FAXSend(txtCorpNum.Text, Statement, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
End Sub

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
    
End Sub

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = statementService.GetChargeInfo(txtCorpNum.Text, selectedItemCode)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (요금) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = statementService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetDetailInfo_Click()
    Dim docDetailInfo As PBStatement
    
    Set docDetailInfo = statementService.GetDetailInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
     
    If docDetailInfo Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "writeDate : " + docDetailInfo.writeDate + vbCrLf
    
    tmp = tmp + "taxType : " + docDetailInfo.taxType + vbCrLf
    
    tmp = tmp + "senderCorpNum : " + docDetailInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderTaxRegID : " + docDetailInfo.senderTaxRegID + vbCrLf
    tmp = tmp + "senderCorpName : " + docDetailInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCEOName : " + docDetailInfo.senderCEOName + vbCrLf
    tmp = tmp + "senderAddr : " + docDetailInfo.senderAddr + vbCrLf
    tmp = tmp + "senderBizClass : " + docDetailInfo.senderBizClass + vbCrLf
    tmp = tmp + "senderBizType : " + docDetailInfo.senderBizType + vbCrLf
    tmp = tmp + "senderContactName : " + docDetailInfo.senderContactName + vbCrLf
    tmp = tmp + "senderDeptName : " + docDetailInfo.senderDeptName + vbCrLf
    tmp = tmp + "senderTEL : " + docDetailInfo.senderTEL + vbCrLf
    tmp = tmp + "senderHP : " + docDetailInfo.senderHP + vbCrLf
    tmp = tmp + "senderEmail : " + docDetailInfo.senderEmail + vbCrLf
    

    tmp = tmp + "receiverCorpNum : " + docDetailInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverTaxRegID : " + docDetailInfo.receiverTaxRegID + vbCrLf
    tmp = tmp + "receiverCorpName : " + docDetailInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCEOName : " + docDetailInfo.receiverCEOName + vbCrLf
    tmp = tmp + "receiverAddr : " + docDetailInfo.receiverAddr + vbCrLf
    tmp = tmp + "receiverBizClass : " + docDetailInfo.receiverBizClass + vbCrLf
    tmp = tmp + "receiverBizType : " + docDetailInfo.receiverBizType + vbCrLf
    tmp = tmp + "receiverContactName : " + docDetailInfo.receiverContactName + vbCrLf
    tmp = tmp + "receiverDeptName : " + docDetailInfo.receiverDeptName + vbCrLf
    tmp = tmp + "receiverTEL : " + docDetailInfo.receiverTEL + vbCrLf
    tmp = tmp + "receiverHP : " + docDetailInfo.receiverHP + vbCrLf
    tmp = tmp + "receiverEmail : " + docDetailInfo.receiverEmail + vbCrLf

    '''  상세내역 생략 '''
    tmp = tmp + "Properties" + vbCrLf
    
    Dim key
    
    For Each key In docDetailInfo.propertyBag.keys
        tmp = tmp + vbTab + key + " : " + docDetailInfo.propertyBag.Item(key) + vbCrLf
    Next
    
    tmp = tmp + "detailList" + vbCrLf
    
    Dim detail As PBDocDetail
     
    For Each detail In docDetailInfo.detailList
        tmp = tmp + vbTab + CStr(detail.serialNum) + " : " + detail.itemName + " | " + detail.supplyCost + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub


Private Sub btnGetEPrintURL_Click()
    Dim url As String
  
    
    url = statementService.GetEPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    
    Set resultList = statementService.GetFiles(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "serialNum | attachedfile | displayName |  RegDT" + vbCrLf
    
    Dim file As PBAttachFile
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetInfo_Click()
    Dim docInfo As PBDocInfo
  
    
    Set docInfo = statementService.GetInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
     
    If docInfo Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "itemKey : " + docInfo.itemKey + vbCrLf
    tmp = tmp + "stateCode : " + CStr(docInfo.stateCode) + vbCrLf
    tmp = tmp + "taxType : " + docInfo.taxType + vbCrLf
    tmp = tmp + "purposeType : " + docInfo.purposeType + vbCrLf
    
    tmp = tmp + "writeDate : " + docInfo.writeDate + vbCrLf
    
    tmp = tmp + "senderCorpName : " + docInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCorpNum : " + docInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderPrintYN : " + CStr(docInfo.senderPrintYN) + vbCrLf
    
    tmp = tmp + "receiverCorpName : " + docInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCorpNum : " + docInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverPrintYN : " + CStr(docInfo.receiverPrintYN) + vbCrLf
    
    tmp = tmp + "supplyCostTotal : " + docInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal : " + docInfo.taxTotal + vbCrLf
    
    tmp = tmp + "issueDT : " + docInfo.issueDT + vbCrLf
    tmp = tmp + "stateDT : " + docInfo.stateDT + vbCrLf
    tmp = tmp + "openYN : " + CStr(docInfo.openYN) + vbCrLf
    tmp = tmp + "openDT : " + docInfo.openDT + vbCrLf
    
    
    tmp = tmp + "stateMemo : " + docInfo.stateMemo + vbCrLf
    
    tmp = tmp + "regDT : " + docInfo.regDT + vbCrLf
    
    
    MsgBox tmp
    
    
End Sub

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    
    '관리번호 배열, 최대 1000건
    KeyList.Add "20160112-01"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    Set resultList = statementService.GetInfos(txtCorpNum.Text, selectedItemCode, KeyList, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT | SenderPrintYN | ReceiverPrintYN " + vbCrLf
    
    Dim info As PBDocInfo
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | "
        tmp = tmp + info.writeDate + " | " + info.regDT + " | " + CStr(info.senderPrintYN) + " | " + CStr(info.receiverPrintYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    
    Set resultList = statementService.GetLogs(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "DocLogType | Log | ProcType | ProcCorpName | ProcMemo | RegDT | IP" + vbCrLf
    
    Dim log As PBDocLog
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = statementService.GetMailURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    KeyList.Add "123123"
    KeyList.Add "123123"
    KeyList.Add "123"
    KeyList.Add "123123123"
    
    url = statementService.GetMassPrintURL(txtCorpNum.Text, selectedItemCode, KeyList, txtUserID.Text)
     
    If url = "" Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = statementService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetPopUpURL_Click()
    Dim url As String
  
    
    url = statementService.GetPopUpURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
    
End Sub

Private Sub btnGetPrintURL_Click()
    Dim url As String
  
    
    url = statementService.GetPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetURL_WRITE_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "WRITE")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnIssue_Click()
    Dim Response As PBResponse
  
    
    Set Response = statementService.Issue(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "발행메모", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    
End Sub


Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.LinkID = LinkID '링크 아이디
    joinData.CorpNum = "1231212312" '사업자번호, "-" 제외.
    joinData.ceoname = "대표자성명"
    joinData.corpName = "회원상호"
    joinData.addr = "주소"
    joinData.ZipCode = "500-100"
    joinData.bizType = "업태"
    joinData.bizClass = "업종"
    joinData.id = "userid"      '6자 이상 20자 미만.
    joinData.pwd = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = statementService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub


Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = statementService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = statementService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.id = "testkorea_20151007"      '담당자 아이디
    joinData.pwd = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = statementService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnRegister_Click()
    Dim Statement As New PBStatement
    
    Statement.writeDate = "20151012"             '필수, 기재상 작성일자
    Statement.purposeType = "영수"               '필수, {영수, 청구}
    Statement.taxType = "과세"                   '필수, {과세, 영세, 면세}
    Statement.formCode = txtFormCode.Text
    
    Statement.itemCode = selectedItemCode
    
    Statement.mgtKey = txtMgtKey.Text
    
    Statement.senderCorpNum = txtCorpNum.Text
    Statement.senderTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Statement.senderCorpName = "공급자 상호"
    Statement.senderCEOName = "공급자"" 대표자 성명"
    Statement.senderAddr = "공급자 주소"
    Statement.senderBizClass = "공급자 업종"
    Statement.senderBizType = "공급자 업태,업태2"
    Statement.senderContactName = "공급자 담당자명"
    Statement.senderEmail = "test@test.com"
    Statement.senderTEL = "070-7070-0707"
    Statement.senderHP = "010-000-2222"
    
    Statement.receiverCorpNum = "8888888888"
    Statement.receiverCorpName = "공급받는자 상호"
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    Statement.receiverAddr = "공급받는자 주소"
    Statement.receiverBizClass = "공급받는자 업종"
    Statement.receiverBizType = "공급받는자 업태"
    Statement.receiverContactName = "공급받는자 담당자명"
    Statement.receiverEmail = "test@receiver.com"
    
    Statement.supplyCostTotal = "100000"         '필수 공급가액 합계
    Statement.taxTotal = "10000"                 '필수 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Statement.serialNum = "123"
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    Statement.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Statement.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    Statement.faxsendYN = False          '발행시 Fax발송시 설정.
    Statement.smssendYN = True '발행시 문자발송기능 사용시 활용
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    Dim i
    Dim newDetail As PBDocDetail
    For i = 1 To 120
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
        newDetail.itemName = "품명" + CStr(i)
        newDetail.spec = "규격"
        newDetail.unit = "단위"
        newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "비고"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '추가속성
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Balance", "100000"
    
    
    Dim Response As PBResponse
    
    Set Response = statementService.Register(txtCorpNum.Text, Statement, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
    

End Sub
Private Sub btnRegistIssue_Click()

    Dim Statement As New PBStatement
    
    Statement.memo = "즉시발행 메모"
    Statement.writeDate = "20151012"             '필수, 기재상 작성일자
    Statement.purposeType = "영수"               '필수, {영수, 청구}
    Statement.taxType = "과세"                   '필수, {과세, 영세, 면세}
    Statement.formCode = txtFormCode.Text
    
    Statement.itemCode = selectedItemCode
    
    Statement.mgtKey = txtMgtKey.Text
    
    Statement.senderCorpNum = txtCorpNum.Text
    Statement.senderTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Statement.senderCorpName = "공급자 상호"
    Statement.senderCEOName = "공급자"" 대표자 성명"
    Statement.senderAddr = "공급자 주소"
    Statement.senderBizClass = "공급자 업종"
    Statement.senderBizType = "공급자 업태,업태2"
    Statement.senderContactName = "공급자 담당자명"
    Statement.senderEmail = "test@test.com"
    Statement.senderTEL = "070-7070-0707"
    Statement.senderHP = "010-000-2222"
    
    Statement.receiverCorpNum = "8888888888"
    Statement.receiverCorpName = "공급받는자 상호"
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    Statement.receiverAddr = "공급받는자 주소"
    Statement.receiverBizClass = "공급받는자 업종"
    Statement.receiverBizType = "공급받는자 업태"
    Statement.receiverContactName = "공급받는자 담당자명"
    Statement.receiverEmail = "test@receiver.com"
    
    Statement.supplyCostTotal = "100000"         '필수 공급가액 합계
    Statement.taxTotal = "10000"                 '필수 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Statement.serialNum = "123"
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    Statement.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Statement.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    Statement.faxsendYN = False          '발행시 Fax발송시 설정.
    Statement.smssendYN = True '발행시 문자발송기능 사용시 활용
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    Dim i
    Dim newDetail As PBDocDetail
    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20140410"   '거래일자  yyyyMMdd
        newDetail.itemName = "품명" + CStr(i)
        newDetail.spec = "규격"
        newDetail.unit = "단위"
        newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "비고"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '추가속성, [참고] 전자명세서 기본양식 추가속성 테이블 참조 http://blog.linkhub.co.kr/2514/
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    
    Dim Response As PBResponse
    
    Set Response = statementService.RegistIssue(txtCorpNum.Text, Statement, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSearch_Click()
    Dim docSearchList As PBDocSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim State As New Collection
    Dim itemCode As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    
    DType = "I"             '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    SDate = "20151001"      '[필수] 시작일자, yyyyMMdd
    EDate = "20160112"      '[필수] 종료일자, yyyyMMdd
    
    State.Add "100"         '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성
    State.Add "2**"         '2,3번째 와일드카드 가능
    State.Add "3**"
    
    itemCode.Add "121"      '문서종류코드 배열, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    itemCode.Add "122"
    itemCode.Add "123"
    itemCode.Add "124"
    itemCode.Add "125"
    itemCode.Add "126"
    
    Page = 1                '페이지 번호
    PerPage = 15            '페이지 목록개수, 최대 1000건
    Order = "D"             '정렬방향, D-내림차순(기본값), A-오름차순
    
    Set docSearchList = statementService.Search(txtCorpNum.Text, DType, SDate, EDate, State, itemCode, Page, PerPage, Order)
     
    If docSearchList Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code : " + CStr(docSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(docSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(docSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(docSearchList.pageNum) + vbCrLf
    tmp = tmp + "perCount : " + CStr(docSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + docSearchList.message + vbCrLf + vbCrLf
    
    
    tmp = tmp + "ItemCode | ItemKey | StateCode | TaxType | WriteDate | SenderCorpName | SenderCorpNum | SenderPrintYN | ReceiverCorpName | ReceiverCorpNum | ReceiverPrintYN " + _
            " | SupplyCostTotal | TaxTotal | RegDT" + vbCrLf
            
    Dim info As PBDocInfo
    
    For Each info In docSearchList.list
        tmp = tmp + CStr(info.itemCode) + " | "
        tmp = tmp + info.itemKey + " | "
        tmp = tmp + CStr(info.stateCode) + " | "
        tmp = tmp + info.taxType + " | "
        tmp = tmp + info.writeDate + " | "
        tmp = tmp + info.senderCorpName + " | "
        tmp = tmp + info.senderCorpNum + " | "
        tmp = tmp + CStr(info.senderPrintYN) + " | "
        
        tmp = tmp + info.receiverCorpName + " | "
        tmp = tmp + info.receiverCorpNum + " | "
        tmp = tmp + CStr(info.receiverPrintYN) + " | "
        
        tmp = tmp + info.supplyCostTotal + " | "
        tmp = tmp + info.taxTotal + " | "
        tmp = tmp + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
       
End Sub

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
  
    
    Set Response = statementService.SendEmail(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "test@test.com", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
  
    
    Set Response = statementService.SendFax(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "07075106766", "111-2222-4444", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
 
    
    Set Response = statementService.SendSMS(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, "07075106766", "111-2222-4444", "문자 내용 최대 90Byte", txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = statementService.GetUnitCost(txtCorpNum.Text, selectedItemCode)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "발행단가 : " + CStr(unitCost)
End Sub

Private Sub btnUpdate_Click()

   
    
    Dim Statement As New PBStatement
    
    Statement.writeDate = "20140319"             '필수, 기재상 작성일자
    Statement.purposeType = "영수"               '필수, {영수, 청구}
    Statement.taxType = "과세"                   '필수, {과세, 영세, 면세}
    Statement.mgtKey = txtMgtKey.Text
    
    Statement.senderCorpNum = "1231212312"
    Statement.senderTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    Statement.senderCorpName = "공급자 상호"
    Statement.senderCEOName = "공급자"" 대표자 성명"
    Statement.senderAddr = "공급자 주소"
    Statement.senderBizClass = "공급자 업종"
    Statement.senderBizType = "공급자 업태,업태2"
    Statement.senderContactName = "공급자 담당자명"
    Statement.senderEmail = "test@test.com"
    Statement.senderTEL = "070-7070-0707"
    Statement.senderHP = "010-000-2222"
    
    Statement.receiverCorpNum = "8888888888"
    Statement.receiverCorpName = "공급받는자 상호"
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    Statement.receiverAddr = "공급받는자 주소"
    Statement.receiverBizClass = "공급받는자 업종"
    Statement.receiverBizType = "공급받는자 업태"
    Statement.receiverContactName = "공급받는자 담당자명"
    Statement.receiverEmail = "test@receiver.com"
    
    Statement.supplyCostTotal = "100000"         '필수 공급가액 합계
    Statement.taxTotal = "10000"                 '필수 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    Statement.serialNum = "123"
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    Statement.businessLicenseYN = False '사업자등록증 이미지 첨부시 설정.
    Statement.bankBookYN = False         '통장사본 이미지 첨부시 설정.
    Statement.faxsendYN = False          '발행시 Fax발송시 설정.
    Statement.smssendYN = True '발행시 문자발송기능 사용시 활용
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    
    Dim newDetail As New PBDocDetail
    
    newDetail.serialNum = 1
    newDetail.purchaseDT = "20140410"
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.qty = "1" '수량
    newDetail.unitCost = "100000"
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    
    Statement.detailList.Add newDetail
    
    Set newDetail = New PBDocDetail
    newDetail.serialNum = 2
    newDetail.itemName = "품명2_수정됨"
    
    Statement.detailList.Add newDetail
    
    
    
    Dim Response As PBResponse
    
    Set Response = statementService.Update(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, Statement, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub


Private Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & " "
    Next l
    
    'Remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
End Function


Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
    Set Response = statementService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.ceoname = "대표자"         '대표자명
    CorpInfo.corpName = "상호"          '상호명
    CorpInfo.addr = "서울특별시"        '주소
    CorpInfo.bizType = "업태"           '업태
    CorpInfo.bizClass = "업종"          '업종
    
    Set Response = statementService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub Form_Load()
    statementService.Initialize LinkID, SecretKey
    
    '연동환경설정값, True-테스트용 False-상업용
    statementService.IsTest = True
End Sub

