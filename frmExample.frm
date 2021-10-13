VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 전자명세서 SDK 예제"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   17370
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   13800
      TabIndex        =   81
      Top             =   165
      Width           =   3255
   End
   Begin VB.CommandButton btnDetachStatement 
      Caption         =   "전자명세서 첨부해제"
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Caption         =   " 전자명세서 관련 기능 "
      Height          =   7020
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   16755
      Begin VB.Frame Frame9 
         Caption         =   "즉시발행 프로세스"
         Height          =   2655
         Left            =   5040
         TabIndex        =   58
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "삭제"
            Height          =   495
            Left            =   1560
            Style           =   1  '그래픽
            TabIndex        =   61
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton btnCancelIssue_sub 
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
         Caption         =   " 문서 정보"
         Height          =   2535
         Left            =   240
         TabIndex        =   46
         Top             =   4200
         Width           =   2010
         Begin VB.CommandButton btnSearch 
            Caption         =   "목록 조회"
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   1580
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "상태 확인"
            Height          =   390
            Left            =   240
            TabIndex        =   50
            Top             =   270
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "상태 대량 확인"
            Height          =   390
            Left            =   240
            TabIndex        =   49
            Top             =   705
            Width           =   1590
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "상태 변경이력"
            Height          =   390
            Left            =   240
            TabIndex        =   48
            Top             =   2000
            Width           =   1590
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "상세정보 확인"
            Height          =   390
            Left            =   240
            TabIndex        =   47
            Top             =   1150
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
            Width           =   2295
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
         Left            =   13920
         TabIndex        =   39
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "발행 문서함"
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
         Caption         =   " 보기/인쇄"
         Height          =   2565
         Left            =   7920
         TabIndex        =   33
         Top             =   4200
         Width           =   5250
         Begin VB.CommandButton btnGetViewURL 
            Caption         =   "전자명세서 보기 URL (메뉴, 버튼x)"
            Height          =   615
            Left            =   3120
            TabIndex        =   76
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "전자명세서 메일링크 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   38
            Top             =   2040
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "대량 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1570
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "전자명세서 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "전자명세서 보기 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintURL 
            Caption         =   "수신자 인쇄 팝업 URL"
            Height          =   390
            Left            =   210
            TabIndex        =   34
            Top             =   1150
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
         Top             =   6000
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
            Height          =   480
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "발행취소"
            Height          =   480
            Left            =   360
            Style           =   1  '그래픽
            TabIndex        =   28
            Top             =   2055
            Width           =   975
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
            Left            =   2280
            Style           =   1  '그래픽
            TabIndex        =   13
            Top             =   2040
            Width           =   975
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
         Caption         =   "문서번호 사용여부 확인"
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
         Caption         =   "문서번호( MgtKey) : "
         Height          =   180
         Left            =   915
         TabIndex        =   9
         Top             =   1455
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2730
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   16755
      Begin VB.Frame Frame16 
         Caption         =   "파트너과금 포인트"
         Height          =   2370
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
         Height          =   2370
         Left            =   11760
         TabIndex        =   68
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   1935
         End
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
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " 회사정보 관련 "
         Height          =   2370
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
         Height          =   2370
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
         Height          =   2370
         Left            =   4800
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   2370
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   2625
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "요금 단가 확인"
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   2370
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   51
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   1320
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   13200
      TabIndex        =   80
      Top             =   240
      Width           =   525
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
' - 업데이트 일자 : 2021-10-07
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

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
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/statement/vb/api#CheckIsMember
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
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/statement/vb/api#CheckID
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
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = ""
    
    Set info = statementService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 사이트와 동일한 전자명세서 1건의 상세 정보 페이지(사이트 상단, 좌측 메뉴 및 버튼 제외)의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetViewURL
'=========================================================================
Private Sub btnGetViewURL_Click()
    Dim url As String
  
    url = statementService.GetViewURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
    
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/statement/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = statementService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서 발행시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetUnitCost
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
' 팝빌 전자명세서 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = statementService.GetChargeInfo(txtCorpNum.Text, selectedItemCode)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (발행단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
        
    url = statementService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 인감 및 첨부문서 등록 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
'=========================================================================
Private Sub btnGetSealURL_Click()

    Dim url As String
    
    url = statementService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/statement/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "VB6STATE_01"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
        
    Set Response = statementService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/statement/vb/api#ListContact
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
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/statement/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = statementService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetCorpInfo
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
' 연동회원의 회사정보를 수정합니다
' - https://docs.popbill.com/statement/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = statementService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/statement/vb/api#GetBalance
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
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = statementService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = statementService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = statementService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/statement/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = statementService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 파트너가 전자명세서 관리 목적으로 할당하는 문서번호의 사용여부를 확인합니다.
' - 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
' - https://docs.popbill.com/statement/vb/api#CheckMgtKeyInUse
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
' 작성된 전자명세서 데이터를 팝빌에 저장과 동시에 발행하여, "발행완료" 상태로 처리합니다.
' - 팝빌 사이트 [전자명세서] > [환경설정] > [전자명세서 관리] 메뉴의 발행시 자동승인 옵션 설정을 통해 전자명세서를 "발행완료" 상태가 아닌 "승인대기" 상태로 발행 처리 할 수 있습니다.
' - https://docs.popbill.com/statement/vb/api#RegistIssue
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    Dim emailSubject As String
    
    Statement.memo = "즉시발행 메모"
    
    '[필수] 기재상 작성일자, 날자형식(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    

    '=========================================================================
    '                               발신자 정보
    '=========================================================================
    
    '발신자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '발신자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '발신자 상호
    Statement.senderCorpName = "발신자 상호"
    
    '발신자 상호명
    Statement.senderCEOName = "발신자 대표자 성명"
    
    '발신자 주소
    Statement.senderAddr = "발신자 주소"
    
    '발신자 종목
    Statement.senderBizClass = "발신자 종목"
    
    '발신자 업태
    Statement.senderBizType = "발신자 업태,업태2"
    
    '발신자 담당자성명
    Statement.senderContactName = "발신자 담당자명"
    
    '발신자 이메일
    Statement.senderEmail = "test@test.com"
    
    '발신자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '발신자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        수신자 정보
    '=========================================================================
    
    '수신자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '수신자 상호
    Statement.receiverCorpName = "수신자 상호"
    
    '수신자 대표자 성명
    Statement.receiverCEOName = "수신자 대표자 성명"
    
    '수신자 주소
    Statement.receiverAddr = "수신자 주소"
    
    '수신자 종목
    Statement.receiverBizClass = "수신자 종목 "
    
    '수신자 업태
    Statement.receiverBizType = "수신자 업태"
    
    '수신자 담당자명
    Statement.receiverContactName = "수신자 담당자명"
    
    '수신자 메일주소
    Statement.receiverEmail = "test@test.com"
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"
        
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
    
    '상세항목 추가. (배열 길이 제한 없음)
    '일련번호(serialNum)은 1부터 순차적으로 기재하시기 바랍니다
    Set Statement.detailList = New Collection
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210902"       '거래일자(yyyyMMdd)
        newDetail.itemName = "품명" + CStr(i)   '품목명
        newDetail.spec = "규격"                 '규격
        newDetail.unit = "단위"                 '단위
        newDetail.qty = "1" '수량               '소수점 2자리까지
        newDetail.unitCost = "100000"           '소수점 2자리까지
        newDetail.supplyCost = "100000"         '공급가액
        newDetail.tax = "10000"                 '세액
        newDetail.remark = "비고"               '비고
        newDetail.spare1 = "spare1"             '여분1
        newDetail.spare2 = "spare2"             '여분2
        newDetail.spare3 = "spare3"             '여분3
        newDetail.spare4 = "spare4"             '여분4
        newDetail.spare5 = "spare5"             '여분5
        
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '전잔액
    Statement.propertyBag.Add "Deposit", "10000"     '입금액
    Statement.propertyBag.Add "Balance", "100000"    '현잔액
    
    '안내메일 제목, 미기재시 기본양식으로 전송.
    emailSubject = ""
    
    Set Response = statementService.RegistIssue(txtCorpNum.Text, Statement, txtUserID.Text, emailSubject)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message + vbCrLf + "팝빌 승인번호 : " + Response.invoiceNum)
End Sub

'=========================================================================
' 발신자가 발행한 전자명세서를 발행취소합니다.
' - https://docs.popbill.com/statement/vb/api#Cancel
'=========================================================================
Private Sub btnCancelIssue_sub_Click()
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
' 삭제 가능한 상태의 전자명세서를 삭제합니다.
' - 삭제 가능한 상태: "임시저장", "취소", "승인거부", "발행취소"
' - 전자명세서를 삭제하면 사용된 문서번호(mgtKey)를 재사용할 수 있습니다.
' - https://docs.popbill.com/statement/vb/api#Delete
'=========================================================================
Private Sub btnDelete_sub_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 작성된 전자명세서 데이터를 팝빌에 저장합니다.
' - https://docs.popbill.com/statement/vb/api#Register
'=========================================================================
Private Sub btnRegister_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[필수] 기재상 작성일자, 날자형식(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               발신자 정보
    '=========================================================================
    
    '발신자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '발신자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '발신자 상호
    Statement.senderCorpName = "발신자 상호"
    
    '발신자 상호명
    Statement.senderCEOName = "발신자 대표자 성명"
    
    '발신자 주소
    Statement.senderAddr = "발신자 주소"
    
    '발신자 종목
    Statement.senderBizClass = "발신자 종목"
    
    '발신자 업태
    Statement.senderBizType = "발신자 업태,업태2"
    
    '발신자 담당자성명
    Statement.senderContactName = "발신자 담당자명"
    
    '발신자 이메일
    Statement.senderEmail = "test@test.com"
    
    '발신자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '발신자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        수신자 정보
    '=========================================================================
    
    '수신자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '수신자 상호
    Statement.receiverCorpName = "수신자 상호"
    
    '수신자 대표자 성명
    Statement.receiverCEOName = "수신자 대표자 성명"
    
    '수신자 주소
    Statement.receiverAddr = "수신자 주소"
    
    '수신자 종목
    Statement.receiverBizClass = "수신자 종목 "
    
    '수신자 업태
    Statement.receiverBizType = "수신자 업태"
    
    '수신자 담당자명
    Statement.receiverContactName = "수신자 담당자명"
    
    '수신자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"
        
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
  
    '상세항목 추가. (배열 길이 제한 없음)
    '일련번호(serialNum)은 1부터 순차적으로 기재하시기 바랍니다
    Set Statement.detailList = New Collection

    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210902"       '거래일자(yyyyMMdd)
        newDetail.itemName = "품명" + CStr(i)   '품목명
        newDetail.spec = "규격"                 '규격
        newDetail.unit = "단위"                 '단위
        newDetail.qty = "1" '수량               '소수점 2자리까지
        newDetail.unitCost = "100000"           '소수점 2자리까지
        newDetail.supplyCost = "100000"         '공급가액
        newDetail.tax = "10000"                 '세액
        newDetail.remark = "비고"               '비고
        newDetail.spare1 = "spare1"             '여분1
        newDetail.spare2 = "spare2"             '여분2
        newDetail.spare3 = "spare3"             '여분3
        newDetail.spare4 = "spare4"             '여분4
        newDetail.spare5 = "spare5"             '여분5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '전잔액
    Statement.propertyBag.Add "Deposit", "10000"     '입금액
    Statement.propertyBag.Add "Balance", "100000"    '현잔액
    
    
    Set Response = statementService.Register(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' "임시저장" 상태의 전자명세서를 수정합니다.건의 전자명세서를 [수정]합니다.' 1건의 전자명세서를 수정합니다.
' - https://docs.popbill.com/statement/vb/api#Update
'=========================================================================
Private Sub btnUpdate_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[필수] 기재상 작성일자, 날자형식(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               발신자 정보
    '=========================================================================
    
    '발신자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '발신자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '발신자 상호
    Statement.senderCorpName = "발신자 상호_수정"
    
    '발신자 상호명
    Statement.senderCEOName = "발신자 대표자 성명"
    
    '발신자 주소
    Statement.senderAddr = "발신자 주소_수정"
    
    '발신자 종목
    Statement.senderBizClass = "발신자 종목_수정"
    
    '발신자 업태
    Statement.senderBizType = "발신자 업태,업태2"
    
    '발신자 담당자성명
    Statement.senderContactName = "발신자 담당자명"
    
    '발신자 이메일
    Statement.senderEmail = "test@test.com"
    
    '발신자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '발신자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        수신자 정보
    '=========================================================================
    
    '수신자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '수신자 상호
    Statement.receiverCorpName = "수신자 상호"
    
    '수신자 대표자 성명
    Statement.receiverCEOName = "수신자 대표자 성명"
    
    '수신자 주소
    Statement.receiverAddr = "수신자 주소"
    
    '수신자 종목
    Statement.receiverBizClass = "수신자 종목 "
    
    '수신자 업태
    Statement.receiverBizType = "수신자 업태"
    
    '수신자 담당자명
    Statement.receiverContactName = "수신자 담당자명"
    
    '수신자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"
        
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
    
    '상세항목 추가. (배열 길이 제한 없음)
    '일련번호(serialNum)은 1부터 순차적으로 기재하시기 바랍니다
    Set Statement.detailList = New Collection
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210902"       '거래일자(yyyyMMdd)
        newDetail.itemName = "품명" + CStr(i)   '품목명
        newDetail.spec = "규격"                 '규격
        newDetail.unit = "단위"                 '단위
        newDetail.qty = "1" '수량               '소수점 2자리까지
        newDetail.unitCost = "100000"           '소수점 2자리까지
        newDetail.supplyCost = "100000"         '공급가액
        newDetail.tax = "10000"                 '세액
        newDetail.remark = "비고"               '비고
        newDetail.spare1 = "spare1"             '여분1
        newDetail.spare2 = "spare2"             '여분2
        newDetail.spare3 = "spare3"             '여분3
        newDetail.spare4 = "spare4"             '여분4
        newDetail.spare5 = "spare5"             '여분5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '전잔액
    Statement.propertyBag.Add "Deposit", "10000"     '입금액
    Statement.propertyBag.Add "Balance", "100000"    '현잔액
    
    Set Response = statementService.Update(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' "임시저장" 상태의 전자명세서를 발행하여, "발행완료" 상태로 처리합니다.
' - 팝빌 사이트 [전자명세서] > [환경설정] > [전자명세서 관리] 메뉴의 발행시 자동승인 옵션 설정을 통해 전자명세서를 "발행완료" 상태가 아닌 "승인대기" 상태로 발행 처리 할 수 있습니다.
' - 전자명세서 발행 함수 호출시 포인트가 과금되며, 수신자에게 발행 안내 메일이 발송됩니다.
' - https://docs.popbill.com/statement/vb/api#StmIssue
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
' 발신자가 발행한 전자명세서를 발행취소합니다.
' - https://docs.popbill.com/statement/vb/api#Cancel
'=========================================================================
Private Sub btnCancelIssue_Click()
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
' 삭제 가능한 상태의 전자명세서를 삭제합니다.
' - 삭제 가능한 상태: "임시저장", "취소", "승인거부", "발행취소"
' - 전자명세서를 삭제하면 사용된 문서번호(mgtKey)를 재사용할 수 있습니다.
' - https://docs.popbill.com/statement/vb/api#Delete
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
' "임시저장" 상태의 명세서에 1개의 파일을 첨부합니다. (최대 5개)
' - https://docs.popbill.com/statement/vb/api#AttachFile
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
' 전자명세서에 첨부된 파일목록을 확인합니다.
' - 응답항목 중 파일아이디(AttachedFile) 항목은 파일삭제(DeleteFile API) 호출시 이용할 수 있습니다.
' - https://docs.popbill.com/statement/vb/api#GetFiles
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
    
    tmp = "serialNum(일련번호) | attachedfile(파일아이디) | displayName(첨부파일명) |  RegDT(첨부일시)" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' "임시저장" 상태의 전자명세서에 첨부된 1개의 파일을 삭제합니다.
' - 파일을 식별하는 파일아이디는 첨부파일 목록(GetFiles API) 의 응답항목 중 파일아이디(AttachedFile) 값을 통해 확인할 수 있습니다.
' - https://docs.popbill.com/statement/vb/api#DeleteFile
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
' 전자명세서의 1건의 상태 및 요약정보 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetInfo
'=========================================================================
Private Sub btnGetInfo_Click()
    Dim docInfo As PBDocInfo
    Dim tmp As String
    
    Set docInfo = statementService.GetInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemCode (명세서 코드) :" + CStr(docInfo.itemCode) + vbCrLf
    tmp = tmp + "itemKey (팝빌번호) : " + docInfo.itemKey + vbCrLf
    tmp = tmp + "invoiceNum (팝빌 승인번호) : " + docInfo.invoiceNum + vbCrLf
    tmp = tmp + "mgtKey (문서번호) : " + docInfo.mgtKey + vbCrLf
    tmp = tmp + "taxType (세금형태) : " + docInfo.taxType + vbCrLf
    tmp = tmp + "writeDate (작성일자) : " + docInfo.writeDate + vbCrLf
    tmp = tmp + "regDT (임시저장일시) : " + docInfo.regDT + vbCrLf
    tmp = tmp + "senderCorpName (발신자 상호) : " + docInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCorpNum (발신자 사업자번호) : " + docInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderPrintYN (발신자 인쇄여부) :" + CStr(docInfo.senderPrintYN) + vbCrLf
    tmp = tmp + "receiverCorpName (수신자 상호) : " + docInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCorpNum (수신자 사업자번호) : " + docInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverPrintYN (수신자 인쇄여부) :" + CStr(docInfo.receiverPrintYN) + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + docInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) : " + docInfo.taxTotal + vbCrLf
    tmp = tmp + "purposeType (영수/청구) : " + docInfo.purposeType + vbCrLf
    tmp = tmp + "issueDT (발행일시) : " + docInfo.issueDT + vbCrLf
    tmp = tmp + "stateCode (상태코드) :" + CStr(docInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (상태 변경일시) : " + docInfo.stateDT + vbCrLf
    tmp = tmp + "stateMemo (상태메모) : " + docInfo.stateMemo + vbCrLf
    tmp = tmp + "openYN (개봉 여부) :" + CStr(docInfo.openYN) + vbCrLf
    tmp = tmp + "openDT (개봉 일시) : " + docInfo.openDT + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 다수건의 전자명세서 상태 및 요약정보 확인합니다. (1회 호출 시 최대 1,000건 확인 가능)
' - https://docs.popbill.com/statement/vb/api#GetInfos
'=========================================================================
Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBDocInfo
    
    '전자명세서 문서번호 배열 (최대 1000건)
    KeyList.Add "20150113-01"
    KeyList.Add "20150113-02"
    KeyList.Add "20150113-03"
    KeyList.Add "20150113-04"
    
    Set resultList = statementService.GetInfos(txtCorpNum.Text, selectedItemCode, KeyList)
            
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
        
    tmp = "itemCode(명세서 코드) | itemKey(팝빌번호) | invoiceNum(팝빌 승인번호) | mgtKey(문서번호) | taxType(세금형태) | " + vbCrLf
    tmp = tmp + "writeDate(작성일자) | regDT(임시저장일시) | senderCorpName(발신자 상호) | senderCorpNum(발신자 사업자번호) | " + vbCrLf
    tmp = tmp + "senderPrintYN(발신자 인쇄여부) | receiverCorpName(수신자 상호) | receiverCorpNum(수신자 사업자번호) | " + vbCrLf
    tmp = tmp + "receiverPrintYN(수신자 인쇄여부) | supplyCostTotal(공급가액 합계) | taxTotal(세액 합계) | purposeType(영수/청구) | " + vbCrLf
    tmp = tmp + "issueDT(발행일시) | stateCode(상태코드) | stateDT(상태 변경일시) | stateMemo(상태메모) | " + vbCrLf
    tmp = tmp + "openYN(개봉 여부) | openDT(개봉 일시)" + vbCrLf + vbCrLf
        
    For Each info In resultList
        tmp = tmp + CStr(info.itemCode) + " | " + info.itemKey + " | " + info.invoiceNum + " | " + info.mgtKey + " | " + info.taxType + " | " + info.writeDate + " | "
        tmp = tmp + info.regDT + " | " + info.senderCorpName + " | " + info.senderCorpNum + CStr(info.senderPrintYN) + " | " + info.receiverCorpName + " | "
        tmp = tmp + info.receiverCorpNum + CStr(info.receiverPrintYN) + " | " + info.supplyCostTotal + " | " + info.taxTotal + " | " + info.purposeType + " | "
        tmp = tmp + info.issueDT + CStr(info.stateCode) + " | " + info.stateDT + " | " + info.stateMemo + CStr(info.openYN) + " | " + info.openDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서 1건의 상세정보 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetDetailInfo
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
    
    tmp = tmp + "itemCode (문서종류 코드) : " + CStr(docDetailInfo.itemCode) + vbCrLf
    tmp = tmp + "mgtKey (문서번호) : " + docDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "invoiceNum (팝빌 승인번호) : " + docDetailInfo.invoiceNum + vbCrLf
    tmp = tmp + "formCode (맞춤양식 코드) : " + docDetailInfo.formCode + vbCrLf
    tmp = tmp + "writeDate (작성일자) : " + docDetailInfo.writeDate + vbCrLf
    tmp = tmp + "taxType (세금형태) : " + docDetailInfo.taxType + vbCrLf
    tmp = tmp + "purposeType (영수/청구) : " + docDetailInfo.purposeType + vbCrLf
    tmp = tmp + "serialNum (일련번호) : " + docDetailInfo.serialNum + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) : " + docDetailInfo.taxTotal + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + docDetailInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "totalAmount (합계금액) : " + docDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "remark1 (비고1) : " + docDetailInfo.remark1 + vbCrLf
    tmp = tmp + "remark2 (비고2) : " + docDetailInfo.remark2 + vbCrLf
    tmp = tmp + "remark3 (비고3) : " + docDetailInfo.remark3 + vbCrLf
        
    tmp = tmp + "senderCorpNum (발신자 사업자번호) : " + docDetailInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderTaxRegID (발신자 종사업장번호) : " + docDetailInfo.senderTaxRegID + vbCrLf
    tmp = tmp + "senderCorpName (발신자 상호) : " + docDetailInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCEOName (발신자 대표자성명) : " + docDetailInfo.senderCEOName + vbCrLf
    tmp = tmp + "senderAddr (발신자 주소) : " + docDetailInfo.senderAddr + vbCrLf
    tmp = tmp + "senderBizClass (발신자 종목) : " + docDetailInfo.senderBizClass + vbCrLf
    tmp = tmp + "senderBizType (발신자 업태) : " + docDetailInfo.senderBizType + vbCrLf
    tmp = tmp + "senderContactName (발신자 성명) : " + docDetailInfo.senderContactName + vbCrLf
    tmp = tmp + "senderDeptName (발신자 부서) : " + docDetailInfo.senderDeptName + vbCrLf
    tmp = tmp + "senderTEL (발신자 연락처) : " + docDetailInfo.senderTEL + vbCrLf
    tmp = tmp + "senderHP (발신자 휴대전화) : " + docDetailInfo.senderHP + vbCrLf
    tmp = tmp + "senderEmail (발신자 이메일) : " + docDetailInfo.senderEmail + vbCrLf
    tmp = tmp + "senderFAX (발신자 팩스) : " + docDetailInfo.senderFAX + vbCrLf

    tmp = tmp + "receiverCorpNum (수신자 사업자번호) : " + docDetailInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverTaxRegID (수신자 종사업장번호) : " + docDetailInfo.receiverTaxRegID + vbCrLf
    tmp = tmp + "receiverCorpName (수신자 상호) : " + docDetailInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCEOName (수신자 대표자성명) : " + docDetailInfo.receiverCEOName + vbCrLf
    tmp = tmp + "receiverAddr (수신자 주소) : " + docDetailInfo.receiverAddr + vbCrLf
    tmp = tmp + "receiverBizClass (수신자 종목) : " + docDetailInfo.receiverBizClass + vbCrLf
    tmp = tmp + "receiverBizType (수신자 업태) : " + docDetailInfo.receiverBizType + vbCrLf
    tmp = tmp + "receiverContactName (수신자 성명) : " + docDetailInfo.receiverContactName + vbCrLf
    tmp = tmp + "receiverDeptName (수신자 부서) : " + docDetailInfo.receiverDeptName + vbCrLf
    tmp = tmp + "receiverTEL (수신자 연락처) : " + docDetailInfo.receiverTEL + vbCrLf
    tmp = tmp + "receiverHP (수신자 휴대전화) : " + docDetailInfo.receiverHP + vbCrLf
    tmp = tmp + "receiverEmail (수신자 이메일) : " + docDetailInfo.receiverEmail + vbCrLf
    tmp = tmp + "receiverFAX (수신자 팩스) : " + docDetailInfo.receiverFAX + vbCrLf
        
    tmp = tmp + "detailList (상세항목)" + vbCrLf
    tmp = tmp + "serialNum(일련번호) | purchaseDT(거래일자) | itemName(품목명) | spec(규격) | qty(수량) |"
    tmp = tmp + "unitCost(단가) | supplyCost(공급가액) | tax(세액) | remark(비고) | spare1(여분1) "
    tmp = tmp + "spare2(여분2) | spare3(여분3) | spare4(여분4) | spare5(여분5) "
    For Each detail In docDetailInfo.detailList
        tmp = tmp + vbTab + CStr(detail.serialNum) + " : " + detail.purchaseDT + " | " + detail.itemName + " | "
        tmp = tmp + detail.spec + " | " + detail.qty + " | " + " | " + detail.unitCost + " | "
        tmp = tmp + detail.supplyCost + " | " + detail.tax + " | " + " | " + detail.remark + " | "
        tmp = tmp + detail.spare1 + " | " + detail.spare2 + " | " + " | " + detail.spare3 + " | "
        tmp = tmp + detail.spare4 + " | " + detail.spare5 + vbCrLf
    Next
    
    tmp = tmp + "Properties (추가속성)" + vbCrLf
    For Each key In docDetailInfo.propertyBag.keys
        tmp = tmp + vbTab + key + " : " + docDetailInfo.propertyBag.Item(key) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 검색조건에 해당하는 전자명세서를 조회합니다. (조회기간 단위 : 최대 6개월)
' - https://docs.popbill.com/statement/vb/api#Search
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
    
    '[필수] 일자유형, R-등록일시 W-작성일자 I-발행일시 중 택1
    DType = "W"
    
    '[필수] 시작일자, yyyyMMdd
    SDate = "20210901"
    
    '[필수] 종료일자, yyyyMMdd
    EDate = "20210910"
    
    '전송상태값 배열, 미기재시 전체상태조회, 문서상태값 3자리숫자 작성 2,3번째 와일드카드 가능
    '상태코드에 대한 자세한 사항은 "[전자명세서 API 연동매뉴얼] > 5.1 전자명세서 상태코드" 를 참조하시기 바랍니다.
    state.Add "100"
    state.Add "2**"
    state.Add "3**"
    
    '명세서 코드 배열, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표, 126-영수증
    itemCode.Add "121"
    itemCode.Add "122"
    itemCode.Add "123"
    itemCode.Add "124"
    itemCode.Add "125"
    itemCode.Add "126"
    
    '페이지 번호, 기본값 '1'
    Page = 1
    
    '페이지당 검색개수, 기본값 '500', 최대 '1000'
    PerPage = 10
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '거래처 정보, 거래처 상호 또는 거래처 사업자등록번호 기재, 미기재시 전체조회
    QString = ""
    
    Set docSearchList = statementService.Search(txtCorpNum.Text, DType, SDate, EDate, state, itemCode, Page, PerPage, Order, QString)
     
    If docSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code (응답코드) : " + CStr(docSearchList.code) + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(docSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(docSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(docSearchList.pageNum) + vbCrLf
    tmp = tmp + "perCount (페이지 개수) : " + CStr(docSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + docSearchList.message + vbCrLf + vbCrLf
    

    tmp = tmp + "itemCode(명세서 코드) | itemKey(팝빌번호) | invoiceNum(팝빌 승인번호) | mgtKey(문서번호) | taxType(세금형태) | " + vbCrLf
    tmp = tmp + "writeDate(작성일자) | regDT(임시저장일시) | senderCorpName(발신자 상호) | senderCorpNum(발신자 사업자번호) | " + vbCrLf
    tmp = tmp + "senderPrintYN(발신자 인쇄여부) | receiverCorpName(수신자 상호) | receiverCorpNum(수신자 사업자번호) | " + vbCrLf
    tmp = tmp + "receiverPrintYN(수신자 인쇄여부) | supplyCostTotal(공급가액 합계) | taxTotal(세액 합계) | purposeType(영수/청구) | " + vbCrLf
    tmp = tmp + "issueDT(발행일시) | stateCode(상태코드) | stateDT(상태 변경일시) | stateMemo(상태메모) | " + vbCrLf
    tmp = tmp + "openYN(개봉 여부) | openDT(개봉 일시)" + vbCrLf + vbCrLf

    Dim info As PBDocInfo
    
    For Each info In docSearchList.list
        tmp = tmp + CStr(info.itemCode) + " | " + info.itemKey + " | " + info.invoiceNum + " | " + info.mgtKey + " | " + info.taxType + " | " + info.writeDate + " | "
        tmp = tmp + info.regDT + " | " + info.senderCorpName + " | " + info.senderCorpNum + CStr(info.senderPrintYN) + " | " + info.receiverCorpName + " | "
        tmp = tmp + info.receiverCorpNum + CStr(info.receiverPrintYN) + " | " + info.supplyCostTotal + " | " + info.taxTotal + " | " + info.purposeType + " | "
        tmp = tmp + info.issueDT + CStr(info.stateCode) + " | " + info.stateDT + " | " + info.stateMemo + CStr(info.openYN) + " | " + info.openDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 전자명세서의 상태에 대한 변경이력을 확인합니다.
' - https://docs.popbill.com/statement/vb/api#GetLogs
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
    
    tmp = "DocLogType(로그타입) | Log(이력정보) | ProcType(처리형태) | ProcCorpName(처리회사명) | ProcMemo(처리메모) | RegDT(등록일시) | IP(아이피)" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' "승인대기", "발행완료" 상태의 전자명세서와 관련된 발행 안내 메일을 재전송 합니다.
' - https://docs.popbill.com/statement/vb/api#SendEmail
'=========================================================================
Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
  
    '수신자 메일주소
    receiverEmail = "test@test.com"
    
    Set Response = statementService.SendEmail(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서와 관련된 안내 SMS(단문) 문자를 재전송하는 함수로, 팝빌 사이트 [문자·팩스] > [문자] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
' - 메시지는 최대 90byte까지 입력 가능하고, 초과한 내용은 자동으로 삭제되어 전송합니다. (한글 최대 45자)
' - 함수 호출시 포인트가 과금됩니다.
' - https://docs.popbill.com/statement/vb/api#SendSMS
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
    
    ' 메시지 내용, 최대 90Byte (한글 45자), 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "전자명세서를 발행하였습니다. 메일을 확인하여 주시기바랍니다"
    
    Set Response = statementService.SendSMS(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서를 팩스로 전송하는 함수로, 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
' - 함수 호출시 포인트가 과금됩니다.
' - https://docs.popbill.com/statement/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    
    '발신자 번호
    senderNum = "070-4304-2991"
    
    '수신자 팩스번호
    receiverNum = "070-111-222"
  
    Set Response = statementService.SendFax(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서를 팩스로 전송하는 함수로, 팝빌에 데이터를 저장하는 과정이 없습니다.
' - 팝빌 사이트 [문자·팩스] > [팩스] > [전송내역] 메뉴에서 전송결과를 확인 할 수 있습니다.
' - 함수 호출시 포인트가 과금됩니다.
' - 팩스 발행 요청시 작성한 문서번호는 팩스전송 파일명으로 사용됩니다.
' - 팩스 전송결과를 확인하기 위해서는 선팩스 전송 요청 시 반환받은 접수번호를 이용하여 팩스 API의 전송결과 확인 (GetFaxDetail) API를 이용하면 됩니다.
' - https://docs.popbill.com/statement/vb/api#FAXSend
'=========================================================================
Private Sub btnFAXSend_Click()
    Dim Statement As New PBStatement
    Dim ReceiptNum As String
    Dim newDetail As PBDocDetail
    Dim i
    
    '팩스 발신번호
    Statement.sendNum = "07043042991"
    
    '팩스 수신번호
    Statement.receiveNum = "070111222"
       
    '[필수] 기재상 작성일자, 날자형식(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[필수] {영수, 청구} 중 기재
    Statement.purposeType = "영수"
    
    '[필수] 과세형태, {과세, 영세, 면세} 중 기재
    Statement.taxType = "과세"
    
    '맞춤양식코드, 공백처리시 기본양식으로 작성
    Statement.formCode = txtFormCode.Text
    
    '[필수] 전자명세서 종류코드
    Statement.itemCode = selectedItemCode
    
    '[필수] 문서번호, 최대 24자리, 영문, 숫자 '-', '_'를 조합하여 사업자별로 중복되지 않도록 구성
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               발신자 정보
    '=========================================================================
    
    '발신자 사업자번호, '-' 제외 10자리
    Statement.senderCorpNum = txtCorpNum.Text
    
    '발신자 종사업장 식별번호, 필요시 기재, 형식은 숫자 4자리
    Statement.senderTaxRegID = ""
    
    '발신자 상호
    Statement.senderCorpName = "발신자 상호"
    
    '발신자 상호명
    Statement.senderCEOName = "발신자 대표자 성명"
    
    '발신자 주소
    Statement.senderAddr = "발신자 주소"
    
    '발신자 종목
    Statement.senderBizClass = "발신자 종목"
    
    '발신자 업태
    Statement.senderBizType = "발신자 업태,업태2"
    
    '발신자 담당자성명
    Statement.senderContactName = "발신자 담당자명"
    
    '발신자 이메일
    Statement.senderEmail = "test@test.com"
    
    '발신자 연락처
    Statement.senderTEL = "070-7070-0707"
    
    '발신자 휴대전화 번호
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        수신자 정보
    '=========================================================================
    
    '수신자 사업자번호, '-' 제외 10자리
    Statement.receiverCorpNum = "8888888888"
    
    '수신자 상호
    Statement.receiverCorpName = "수신자 상호"
    
    '수신자 대표자 성명
    Statement.receiverCEOName = "수신자 대표자 성명"
    
    '수신자 주소
    Statement.receiverAddr = "수신자 주소"
    
    '수신자 종목
    Statement.receiverBizClass = "수신자 종목 "
    
    '수신자 업태
    Statement.receiverBizType = "수신자 업태"
    
    '수신자 담당자명
    Statement.receiverContactName = "수신자 담당자명"
    
    '수신자 메일주소
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     전자명세서 기재사항
    '=========================================================================
    
    '[필수] 공급가액 합계
    Statement.supplyCostTotal = "100000"
    
    '[필수] 세액 합계
    Statement.taxTotal = "10000"
    
    '[필수] 합계금액, 공급가액 합계 + 세액 합계
    Statement.totalAmount = "110000"
        
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
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '일련번호 1부터 순차 기재
        newDetail.purchaseDT = "20210902"       '거래일자(yyyyMMdd)
        newDetail.itemName = "품명" + CStr(i)   '품목명
        newDetail.spec = "규격"                 '규격
        newDetail.unit = "단위"                 '단위
        newDetail.qty = "1" '수량               '소수점 2자리까지
        newDetail.unitCost = "100000"           '소수점 2자리까지
        newDetail.supplyCost = "100000"         '공급가액
        newDetail.tax = "10000"                 '세액
        newDetail.remark = "비고"               '비고
        newDetail.spare1 = "spare1"             '여분1
        newDetail.spare2 = "spare2"             '여분2
        newDetail.spare3 = "spare3"             '여분3
        newDetail.spare4 = "spare4"             '여분4
        newDetail.spare5 = "spare5"             '여분5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '전자명세서 추가속성
    ' - 추가속성에 관한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
    '   5.2. 기본양식 추가속성 테이블"을 참조하시기 바랍니다.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '전잔액
    Statement.propertyBag.Add "Deposit", "10000"     '입금액
    Statement.propertyBag.Add "Balance", "100000"    '현잔액
    
    ReceiptNum = statementService.FAXSend(txtCorpNum.Text, Statement)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
End Sub

'=========================================================================
' 하나의 전자명세서에 다른 전자명세서를 첨부합니다.
' - https://docs.popbill.com/statement/vb/api#AttachStatement
'=========================================================================
Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '첨부할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표, 126-영수증
    SubItemCode = 121
    
    '첨부할 전자명세서 문서번호
    SubMgtKey = "20210902-01"
    
    Set Response = statementService.AttachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 하나의 전자명세서에 첨부된 다른 전자명세서를 해제합니다.
' - https://docs.popbill.com/statement/vb/api#DetachStatement
'=========================================================================
Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '첨부해제할 전자명세서 종류코드, 121-거래명세서, 122-청구서, 123-견적서, 124-발주서, 125-입금표, 126-영수증
    SubItemCode = 121
    
    '첨부해제할 전자명세서 문서번호
    SubMgtKey = "20210902-01"
      
    Set Response = statementService.DetachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 전자명세서 관련 메일 항목에 대한 발송설정을 확인합니다.
' - https://docs.popbill.com/statement/vb/api#ListEmailConfig
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
            tmp = tmp + "수신자에게 전자명세서가 발행 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_ACCEPT" Then
            tmp = tmp + "발신자에게 전자명세서가 승인 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_DENY" Then
            tmp = tmp + "발신자에게 전자명세서가 거부 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL" Then
            tmp = tmp + "수신자에게 전자명세서가 취소 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL_ISSUE" Then
            tmp = tmp + "수신자에게 전자명세서가 발행취소 되었음을 알려주는 메일 : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' 전자명세서 관련 메일 항목에 대한 발송설정을 수정합니다.
' - https://docs.popbill.com/statement/vb/api#UpdateEmailConfig
'
' 메일전송유형
' SMT_ISSUE : 공급받는자에게 전자명세서가 발행 되었음을 알려주는 메일입니다.
' SMT_ACCEPT : 공급자에게 전자명세서가 승인 되었음을 알려주는 메일입니다.
' SMT_DENY : 공급자에게 전자명세서가 거부 되었음을 알려주는 메일입니다.
' SMT_CANCEL : 공급받는자에게 전자명세서가 취소 되었음을 알려주는 메일입니다.
' SMT_CANCEL_ISSUE : 공급받는자에게 전자명세서가 발행취소 되었음을 알려주는 메일입니다.
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
' 팝빌 사이트와 동일한 전자명세서 1건의 상세 정보 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetPopUpURL
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim url As String
  
    url = statementService.GetPopUpURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 전자명세서 1건을 인쇄하기 위한 페이지의 팝업 URL을 반환하며, 페이지내에서 인쇄 설정값을 "공급자" / "공급받는자" / "공급자+공급받는자"용 중 하나로 지정할 수 있습니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetPrintURL
'=========================================================================
Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = statementService.GetPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' "공급받는자" 용 전자명세서 1건을 인쇄하기 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetEPrintURL
'=========================================================================
Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = statementService.GetEPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 다수건의 전자명세서를 인쇄하기 위한 페이지의 팝업 URL을 반환합니다. (최대 100건)
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetMassPrintURL
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    '전자명세서 문서번호 배열 (최대 100건)
    KeyList.Add "20210902-01"
    KeyList.Add "20210902-02"
    KeyList.Add "20210902-03"
    KeyList.Add "20210902-04"
    
    url = statementService.GetMassPrintURL(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 안내메일과 관련된 전자명세서를 확인 할 수 있는 상세 페이지의 팝업 URL을 반환하며, 해당 URL은 메일 하단의 파란색 버튼의 링크와 같습니다.
' - 함수 호출로 반환 받은 URL에는 유효시간이 없습니다.
' - https://docs.popbill.com/statement/vb/api#GetMailURL
'=========================================================================
Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = statementService.GetMailURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 로그인 상태로 팝빌 사이트의 전자명세서 매출문서함 메뉴에 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 로그인 상태로 팝빌 사이트의 전자명세서 임시문서함 메뉴에 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/statement/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(statementService.LastErrCode) + vbCrLf + "응답메시지 : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

Private Sub Form_Load()

    '전자명세서 객체 초기화
    statementService.Initialize LinkID, SecretKey
    
    '연동환경설정값, True-개발용 False-상업용
    statementService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-권장
    statementService.IPRestrictOnOff = True
    
    ' 팝빌 API 서비스 고정 IP 사용여부, True-사용, False-미사용, 기본값(False)
    statementService.UseStaticIP = False
    
    ' 로컬시스템 시간 사용여부 True-사용, Fasle-미사용, 기본값(False)
    statementService.UseLocalTimeYN = False
    
    
End Sub

