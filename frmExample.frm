VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 전자명세서 SDK 예제"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   17565
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnDetachStatement 
      Caption         =   "전자명세서 첨부해제"
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Top             =   8040
      Width           =   2235
   End
   Begin VB.Frame Frame7 
      Caption         =   " 전자명세서 관련 기능 "
      Height          =   7380
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   13995
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스"
         Height          =   2655
         Left            =   5040
         TabIndex        =   58
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "삭제"
            Height          =   495
            Left            =   1560
            Style           =   1  '그래픽
            TabIndex        =   61
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton btnCancelISsue_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   480
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   60
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "즉시발행"
            Height          =   405
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   59
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
         TabIndex        =   46
         Top             =   4200
         Width           =   2010
         Begin VB.CommandButton btnSearch 
            Caption         =   "문서 목록조회"
            Height          =   375
            Left            =   210
            TabIndex        =   63
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "문서 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   50
            Top             =   270
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "문서 정보(대량)"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   705
            Width           =   1590
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "문서 이력"
            Height          =   390
            Left            =   210
            TabIndex        =   48
            Top             =   1140
            Width           =   1590
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "문서 상세 정보"
            Height          =   390
            Left            =   210
            TabIndex        =   47
            Top             =   1590
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " 부가 서비스"
         Height          =   2295
         Left            =   2520
         TabIndex        =   42
         Top             =   4200
         Width           =   4980
         Begin VB.CommandButton btnUpdateemailconfig 
            Caption         =   "알림메일 전송목록 수정"
            Height          =   375
            Left            =   2520
            TabIndex        =   75
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton btnListemailconfig 
            Caption         =   "알림메일 전송목록 조회"
            Height          =   375
            Left            =   2520
            TabIndex        =   74
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "전자명세서 첨부"
            Height          =   375
            Left            =   2520
            TabIndex        =   64
            Top             =   300
            Width           =   2235
         End
         Begin VB.CommandButton btnFAXSEnd 
            Caption         =   "선팩스 전송"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   1640
            Width           =   2115
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "팩스 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   45
            Top             =   1200
            Width           =   2115
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "문자 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   44
            Top             =   735
            Width           =   2115
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "이메일 전송"
            Height          =   390
            Left            =   225
            TabIndex        =   43
            Top             =   300
            Width           =   2115
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " 기타 URL "
         Height          =   1290
         Left            =   11400
         TabIndex        =   39
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton btnGetURL_SBOX 
            Caption         =   "매출 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   41
            Top             =   705
            Width           =   1500
         End
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "임시 문서함"
            Height          =   390
            Left            =   210
            TabIndex        =   40
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " 문서 정보 "
         Height          =   2565
         Left            =   7920
         TabIndex        =   33
         Top             =   4200
         Width           =   3210
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "이메일(공급받는자) 링크 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   38
            Top             =   1590
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "다량 인쇄 팝업 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   37
            Top             =   1140
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "문서 내용 보기 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintURL 
            Caption         =   "수신자 인쇄 팝업 URL"
            Height          =   390
            Left            =   195
            TabIndex        =   34
            Top             =   2040
            Width           =   2745
         End
      End
      Begin VB.TextBox txtMgtKey 
         Height          =   330
         Left            =   2730
         TabIndex        =   31
         Top             =   1350
         Width           =   2025
      End
      Begin VB.TextBox txtFormCode 
         Height          =   345
         Left            =   2730
         TabIndex        =   30
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
            TabIndex        =   29
            Top             =   1365
            Width           =   1020
         End
         Begin VB.CommandButton btnCancel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   480
            Left            =   285
            Style           =   1  '그래픽
            TabIndex        =   28
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
         TabIndex        =   32
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
      Height          =   2250
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   16755
      Begin VB.Frame Frame16 
         Caption         =   "파트너과금 포인트"
         Height          =   1695
         Left            =   14040
         TabIndex        =   69
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "연동과금 포인트"
         Height          =   1695
         Left            =   11760
         TabIndex        =   68
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   360
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련 "
         Height          =   1695
         Left            =   9600
         TabIndex        =   55
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   1770
         Left            =   6840
         TabIndex        =   26
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton btnGetSealURL 
            Caption         =   "인감 및 첨부문서 등록 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   "팝빌 로그인 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   1770
         Left            =   4800
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   1770
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   2625
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   1770
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   51
            Top             =   735
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   1215
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
'=========================================================================
'
' 팝빌 전자명세서 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 :
' - 업데이트 일자 : 2017-08-30
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const LinkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'전자명세서 서비스 객체 생성
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

'=========================================================================
' 전자명세서에 첨부파일을 등록합니다.
' - 첨부파일 등록은 전자명세서가 [임시저장] 상태인 경우에만 가능합니다.
' - 첨부파일은 최대 5개까지 등록할 수 있습니다.
'=========================================================================

Private Sub btnAttachFile_Click()
    Dim FilePath As String
    Dim Response As PBResponse
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    FilePath = CommonDialog1.FileName
    
    If FilePath = "" Then Exit Sub
    
    Set Response = statementService.AttachFile(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, FilePath)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 전자명세서에 다른 전자명세서 1건을 첨부합니다.
'=========================================================================

Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubItemCode = 121
    
    '첨부할 전자명세서 관리번호
    SubMgtKey = "20151223-01"
      
    Set Response = statementService.AttachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자명세서를 [발행취소] 처리합니다.
'=========================================================================

Private Sub btnCancel_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "전자명세서 발행취소 메모"
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자명세서를 [발행취소] 처리합니다.
'=========================================================================

Private Sub btnCancelISsue_2_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "발행 취소 메모"
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서 관리번호 중복여부를 확인합니다.
' - 관리번호는 1~24자리로 숫자, 영문 '-', '_' 조합으로 구성할 수 있습니다.
'=========================================================================

Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckMgtKeyInUse(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 1건의 전자명세서를 삭제합니다.
' - 전자명세서를 삭제하면 사용된 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소]
'=========================================================================

Private Sub btnDelete_2_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 1건의 전자명세서를 삭제합니다.
' - 전자명세서를 삭제하면 사용된 문서관리번호(mgtKey)를 재사용할 수 있습니다.
' - 삭제가능한 문서 상태 : [임시저장], [발행취소]
'=========================================================================

Private Sub btnDelete_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서에 첨부된 파일을 삭제합니다.
' - 파일을 식별하는 파일아이디는 첨부파일 목록(GetFileList API) 의 응답항목
'   중 파일아이디(AttachedFile) 값을 통해 확인할 수 있습니다.
'=========================================================================

Private Sub btnDeleteFile_Click()
    Dim Response As PBResponse
   
    Set Response = statementService.DeleteFile(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtFileID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서에 첨부된 다른 전자명세서를 첨부해제합니다.
'=========================================================================

Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    SubItemCode = 121
    
    '첨부해제할 전자명세서 관리번호
    SubMgtKey = "20151223-01"
      
    Set Response = statementService.DetachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌에 등록하지 않고 전자명세서를 팩스전송합니다.
' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [팩스] > [전송내역]
'   메뉴에서 전송결과를 확인할 수 있습니다.
'=========================================================================

Private Sub btnFAXSEnd_Click()
    Dim Statement As New PBStatement
    Dim ReceiptNum As String
    Dim newDetail As PBDocDetail
    Dim i
    
    '팩스 발신번호
    Statement.sendNum = "07043042991"
    
    '팩스 수신번호
    Statement.receiveNum = "070000111"
       
    '[필수] 기재상 작성일자, 날짜형식(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서관리번호, 숫자, 영문, '-', '_' 조합 (최대24자리)으로 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               공급자 정보
    '=========================================================================
    
    '공급자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '공급자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '공급자 상호
    Statement.senderCorpName = "공급자 상호"
    
    '공급자 대표자 성명
    Statement.senderCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Statement.senderAddr = "공급자 주소"
    
    '공급자 종목
    Statement.senderBizClass = "공급자 종목"
    
    '공급자 업태
    Statement.senderBizType = "공급자 업태,업태2"
    
    '공급자 담당자성명
    Statement.senderContactName = "공급자 담당자명"
    
    '공급자 이메일
    Statement.senderEmail = "test@test.com"
    
    '공급자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '공급자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        공급받는자 정보
    '=========================================================================
    
    '공급받는자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '공급받는자 상호
    Statement.receiverCorpName = "공급받는자 상호"
    
    '공급받는자 대표자 성명
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Statement.receiverAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Statement.receiverBizClass = "공급받는자 종목 "
    
    '공급받는자 업태
    Statement.receiverBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Statement.receiverContactName = "공급받는자 담당자명"
    
    '공급받는자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
        
    '기재 상 일련번호 항목
    Statement.serialNum = "123"
    
    '기재 상 비고 항목
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Statement.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Statement.bankBookYN = False
    
    '발행시 알림문자 발송여부
    Statement.smssendYN = True
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    
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
    
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    ReceiptNum = statementService.FAXSend(txtCorpNum.Text, Statement)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
End Sub



'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 연동회원의 전자명세서 API 서비스 과금정보를 확인합니다.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = statementService.GetChargeInfo(txtCorpNum.Text, selectedItemCode)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub



'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = statementService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서 1건의 상세정보를 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[전자명세서 API 연동매뉴얼] > 4.1.
'   전자명세서 구성" 을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetDetailInfo_Click()
    Dim docDetailInfo As PBStatement
    Dim tmp As String
    Dim key
    Dim detail As PBDocDetail
    
    Set docDetailInfo = statementService.GetDetailInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docDetailInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
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
    
    
    For Each key In docDetailInfo.propertyBag.keys
        tmp = tmp + vbTab + key + " : " + docDetailInfo.propertyBag.Item(key) + vbCrLf
    Next
    
    tmp = tmp + "detailList" + vbCrLf
     
    For Each detail In docDetailInfo.detailList
        tmp = tmp + vbTab + CStr(detail.serialNum) + " : " + detail.itemName + " | " + detail.supplyCost + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 전자명세서 인쇄(공급받는자) URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetEPrintURL_Click()
    Dim url As String
    
    url = statementService.GetEPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 전자명세서에 첨부된 파일의 목록을 확인합니다.
' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API)
'   호출시 이용할 수 있습니다.
'=========================================================================

Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim file As PBAttachFile
    
    Set resultList = statementService.GetFiles(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "serialNum | attachedfile | displayName |  RegDT" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 1건의 전자명세서 상태/요약 정보를 확인합니다.
' - 응답항목에 대한 자세한 정보는 "[전자명세서 API 연동매뉴얼] > 3.3.1.
'   GetInfo (상태 확인)"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetInfo_Click()
    Dim docInfo As PBDocInfo
    Dim tmp As String
    
    Set docInfo = statementService.GetInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' 다수건의 전자명세서 상태/요약 정보를 확인합니다.
' - 응답항목에 대한 자세한 정보는 "[전자명세서 API 연동매뉴얼] > 3.3.2.
'   GetInfos (상태 대량 확인)"을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBDocInfo
    
    '전자명세서 관리번호 배열, 최대 1000건
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    Set resultList = statementService.GetInfos(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "ItemKey | StateCode | TaxType | WriteDate | RegDT | SenderPrintYN | ReceiverPrintYN " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.itemKey + " | " + CStr(info.stateCode) + " | " + info.taxType + " | "
        tmp = tmp + info.writeDate + " | " + info.regDT + " | " + CStr(info.senderPrintYN) + " | " + CStr(info.receiverPrintYN) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서 상태 변경이력을 확인합니다.
' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
'   "[전자명세서 API 연동매뉴얼] > 3.3.4 GetLogs (상태 변경이력 확인)"
'   을 참조하시기 바랍니다.
'=========================================================================

Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim log As PBDocLog
    
    Set resultList = statementService.GetLogs(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType | Log | ProcType |  ProcMemo | RegDT | IP" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 공급받는자 메일링크 URL을 반환합니다.
' - 메일링크 URL은 유효시간이 존재하지 않습니다.
'=========================================================================

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = statementService.GetMailURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 다수건의 전자명세서 인쇄팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    '인쇄할 전자명세서 관리번호 배열, 최대 100건
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    url = statementService.GetMassPrintURL(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 파트너 포인트 충전 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = statementService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = statementService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 인감 및 첨부문서 등록 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetSealURL_Click()
    Dim url As String
    
    url = statementService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1건의 전자명세서 보기 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetPopUpURL_Click()
    Dim url As String
    
    url = statementService.GetPopUpURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 전자명세서 인쇄팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetPrintURL_Click()
    Dim url As String
  
    url = statementService.GetPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub



'=========================================================================
' 팝빌 > 전자명세서 > 매출문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 전자명세서 > 임시(연동)문서함 팝업 URL을 반환합니다.
' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1건의 [임시저장] 상태의 전자명세서를 발행처리합니다.
' - 발행시 포인트가 차감됩니다.
'=========================================================================

Private Sub btnIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '메모
    memo = "전자명세서 발행 메모"
    
    Set Response = statementService.Issue(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 연동회원 가입을 요청합니다.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1231212312"
    
    '대표자성명, 최대 30자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 70자
    joinData.corpName = "회원상호"
    
    '주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 40자
    joinData.bizType = "업태"
    
    '종목, 최대 40자
    joinData.bizClass = "종목"
    
    '아이디, 6자이상 20자 미만
    joinData.id = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.pwd = "pwd_must_be_long_enough"
    
    '담당자명, 최대 30자
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    '담당자 메일, 최대 70자
    joinData.ContactEmail = "test@test.com"
    
    Set Response = statementService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = statementService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서 관련 메일전송 항목에 대한 전송여부를 목록으로 반환합니다
'=========================================================================
Private Sub btnListemailconfig_Click()
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = statementService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "메일전송유형(EmailType) | 전송여부(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "SMT_ISSUE" Then
            tmp = tmp + "공급받는자에게 전자명세서가 발행 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_ACCEPT" Then
            tmp = tmp + "공급자에게 전자명세서가 승인 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_DENY" Then
            tmp = tmp + "공급자에게 전자명세서가 거부 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL" Then
            tmp = tmp + "공급받는자에게 전자명세서가 취소 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL_ISSUE" Then
            tmp = tmp + "공급받는자에게 전자명세서가 발행취소 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서 관련 메일전송 항목에 대한 전송여부를 수정합니다.
'=========================================================================
Private Sub btnUpdateemailconfig_Click()
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '메일 전송 유형
    emailType = "SMT_ISSUE"

    '전송 여부 (True = 전송, False = 미전송)
    sendYN = True
    
    Set Response = statementService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = statementService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 20자 미만
    joinData.id = "testkorea_20161011"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.pwd = "test@test.com"
    
    '담당자명, 최대 30자
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
    
    '담당자 팩스번호
    joinData.fax = "070-1234-1234"
    
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
        
    Set Response = statementService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

Private Sub btnRegister_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[필수] 기재상 작성일자, 날짜형식(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서관리번호, 숫자, 영문, '-', '_' 조합 (최대24자리)으로 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               공급자 정보
    '=========================================================================
    
    '공급자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '공급자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '공급자 상호
    Statement.senderCorpName = "공급자 상호"
    
    '공급자 대표자 성명
    Statement.senderCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Statement.senderAddr = "공급자 주소"
    
    '공급자 종목
    Statement.senderBizClass = "공급자 종목"
    
    '공급자 업태
    Statement.senderBizType = "공급자 업태,업태2"
    
    '공급자 담당자성명
    Statement.senderContactName = "공급자 담당자명"
    
    '공급자 이메일
    Statement.senderEmail = "test@test.com"
    
    '공급자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '공급자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        공급받는자 정보
    '=========================================================================
    
    '공급받는자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '공급받는자 상호
    Statement.receiverCorpName = "공급받는자 상호"
    
    '공급받는자 대표자 성명
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Statement.receiverAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Statement.receiverBizClass = "공급받는자 종목 "
    
    '공급받는자 업태
    Statement.receiverBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Statement.receiverContactName = "공급받는자 담당자명"
    
    '공급받는자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
        
    '기재 상 일련번호 항목
    Statement.serialNum = "123"
    
    '기재 상 비고 항목
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Statement.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Statement.bankBookYN = False
    
    '발행시 알림문자 발송여부
    Statement.smssendYN = True
  
    '상세항목 추가.
    Set Statement.detailList = New Collection

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
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.Register(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 1건의 전자명세서를 즉시발행 처리합니다.
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    Statement.memo = "즉시발행 메모"
    
    '[필수] 기재상 작성일자, 날짜형식(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서관리번호, 숫자, 영문, '-', '_' 조합 (최대24자리)으로 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               공급자 정보
    '=========================================================================
    
    '공급자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '공급자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '공급자 상호
    Statement.senderCorpName = "공급자 상호"
    
    '공급자 대표자 성명
    Statement.senderCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Statement.senderAddr = "공급자 주소"
    
    '공급자 종목
    Statement.senderBizClass = "공급자 종목"
    
    '공급자 업태
    Statement.senderBizType = "공급자 업태,업태2"
    
    '공급자 담당자성명
    Statement.senderContactName = "공급자 담당자명"
    
    '공급자 이메일
    Statement.senderEmail = "test@test.com"
    
    '공급자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '공급자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        공급받는자 정보
    '=========================================================================
    
    '공급받는자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '공급받는자 상호
    Statement.receiverCorpName = "공급받는자 상호"
    
    '공급받는자 대표자 성명
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Statement.receiverAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Statement.receiverBizClass = "공급받는자 종목 "
    
    '공급받는자 업태
    Statement.receiverBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Statement.receiverContactName = "공급받는자 담당자명"
    
    '공급받는자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
        
    '기재 상 일련번호 항목
    Statement.serialNum = "123"
    
    '기재 상 비고 항목
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Statement.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Statement.bankBookYN = False
    
    '발행시 알림문자 발송여부
    Statement.smssendYN = True
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    
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
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.RegistIssue(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 검색조건을 사용하여 전자명세서 목록을 조회합니다.
' - 응답항목에 대한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
'   3.3.3. Search (목록 조회)" 를 참조하시기 바랍니다.
'=========================================================================

Private Sub btnSearch_Click()
    Dim docSearchList As PBDocSearchList
    Dim DType As String
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim itemCode As New Collection
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim QString As String
    Dim tmp As String
    Dim info As PBDocInfo
    
    '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    DType = "W"
    
    '[필수] 시작일자, yyyyMMdd
    SDate = "20160901"
    
    '[필수] 종료일자, yyyyMMdd
    EDate = "20161031"
    
    '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성
    '2,3번째 와일드카드 가능
    state.Add "100"
    state.Add "2**"
    state.Add "3**"
    
    '문서종류코드 배열, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표,126-영수증
    itemCode.Add "121"
    itemCode.Add "122"
    itemCode.Add "123"
    itemCode.Add "124"
    itemCode.Add "125"
    itemCode.Add "126"
    
    '페이지 번호
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 15
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '거래처 정보, 거래처 상호 또는 거래처 사업자등록번호 기재, 미기재시 전체조회
    QString = ""
        
    Set docSearchList = statementService.Search(txtCorpNum.Text, DType, SDate, EDate, state, itemCode, Page, PerPage, Order, QString)
     
    If docSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code : " + CStr(docSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(docSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(docSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(docSearchList.pageNum) + vbCrLf
    tmp = tmp + "perCount : " + CStr(docSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + docSearchList.message + vbCrLf + vbCrLf
    
    
    tmp = tmp + "ItemCode | ItemKey | StateCode | TaxType | WriteDate | SenderCorpName | SenderCorpNum | SenderPrintYN | ReceiverCorpName | ReceiverCorpNum | ReceiverPrintYN " + _
            " | SupplyCostTotal | TaxTotal | RegDT" + vbCrLf
    
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

'=========================================================================
' 발행 안내메일을 재전송합니다.
'=========================================================================

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
    
    '수신메일주소
    receiverEmail = "test@test.com"
  
    Set Response = statementService.SendEmail(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서를 팩스전송합니다.
' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [팩스] > [전송내역]
'   메뉴에서 전송결과를 확인할 수 있습니다.
'=========================================================================

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    
    '발신번호
    senderNum = "070-4304-2991"
    
    '수신팩스번호
    receiverNum = "070-111-222"
    
    Set Response = statementService.SendFax(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 알림문자를 전송합니다. (단문/SMS- 한글 최대 45자)
' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [전송내역] 탭에서
'   전송결과를 확인할 수 있습니다.
'=========================================================================

Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    Dim Contents As String
    
    '발신번호
    senderNum = "070-4304-2991"
    
    '수신번호
    receiverNum = "010-111-222"
    
    '문자메시지 내용, 90Byte 초과된 내용은 삭제되어 전송됨
    Contents = "전자명세서를 발행하였습니다. 메일을 확인하여 주시기바랍니다"
    
    Set Response = statementService.SendSMS(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서 발행단가를 확인합니다.
'=========================================================================

Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = statementService.GetUnitCost(txtCorpNum.Text, selectedItemCode)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "발행단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 1건의 전자명세서를 수정합니다.
' - [임시저장] 상태의 전자명세서만 수정할 수 있습니다.
'=========================================================================

Private Sub btnUpdate_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[필수] 기재상 작성일자, 날짜형식(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서관리번호, 숫자, 영문, '-', '_' 조합 (최대24자리)으로 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               공급자 정보
    '=========================================================================
    
    '공급자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '공급자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '공급자 상호
    Statement.senderCorpName = "공급자 상호_수정"
    
    '공급자 대표자 성명
    Statement.senderCEOName = "공급자 대표자 성명"
    
    '공급자 주소
    Statement.senderAddr = "공급자 주소_수정"
    
    '공급자 종목
    Statement.senderBizClass = "공급자 종목_수정"
    
    '공급자 업태
    Statement.senderBizType = "공급자 업태,업태2"
    
    '공급자 담당자성명
    Statement.senderContactName = "공급자 담당자명"
    
    '공급자 이메일
    Statement.senderEmail = "test@test.com"
    
    '공급자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '공급자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        공급받는자 정보
    '=========================================================================
    
    '공급받는자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '공급받는자 상호
    Statement.receiverCorpName = "공급받는자 상호"
    
    '공급받는자 대표자 성명
    Statement.receiverCEOName = "공급받는자 대표자 성명"
    
    '공급받는자 주소
    Statement.receiverAddr = "공급받는자 주소"
    
    '공급받는자 종목
    Statement.receiverBizClass = "공급받는자 종목 "
    
    '공급받는자 업태
    Statement.receiverBizType = "공급받는자 업태"
    
    '공급받는자 담당자명
    Statement.receiverContactName = "공급받는자 담당자명"
    
    '공급받는자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
        
    '기재 상 일련번호 항목
    Statement.serialNum = "123"
    
    '기재 상 비고 항목
    Statement.remark1 = "비고1"
    Statement.remark2 = "비고2"
    Statement.remark3 = "비고3"
    
    '사업자등록증 이미지 첨부여부
    Statement.businessLicenseYN = False
    
    '통장사본 이미지 첨부여부
    Statement.bankBookYN = False
    
    '발행시 알림문자 발송여부
    Statement.smssendYN = True
  
    
    '상세항목 추가.
    Set Statement.detailList = New Collection
    
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
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.Update(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-4304-2991"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, Ture-회사조회, False-개인조
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
                
    Set Response = statementService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명
    CorpInfo.ceoname = "대표자"
    
    '상호
    CorpInfo.corpName = "상호"
    
    '주소
    CorpInfo.addr = "서울특별시"
    
    '업태
    CorpInfo.bizType = "업태"
    
    '종목
    CorpInfo.bizClass = "종목"
    
    Set Response = statementService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub



Private Sub Form_Load()
    '전자명세서 객체 초기화
    statementService.Initialize LinkID, SecretKey
    
    '연동환경설정값, True-개발용 False-상업용
    statementService.IsTest = True
End Sub

