VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ڸ��� SDK ����"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   17370
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   13800
      TabIndex        =   81
      Top             =   165
      Width           =   3255
   End
   Begin VB.CommandButton btnDetachStatement 
      Caption         =   "���ڸ��� ÷������"
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Caption         =   " ���ڸ��� ���� ��� "
      Height          =   7020
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   16755
      Begin VB.Frame Frame9 
         Caption         =   "��ù��� ���μ���"
         Height          =   2655
         Left            =   5040
         TabIndex        =   58
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton btnDelete_sub 
            Caption         =   "����"
            Height          =   495
            Left            =   1560
            Style           =   1  '�׷���
            TabIndex        =   61
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton btnCancelIssue_sub 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   480
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   60
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton btnRegistIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ù���"
            Height          =   405
            Left            =   360
            Style           =   1  '�׷���
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
            BackStyle       =   1  '�������� ����
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
         Caption         =   " ���� ����"
         Height          =   2535
         Left            =   240
         TabIndex        =   46
         Top             =   4200
         Width           =   2010
         Begin VB.CommandButton btnSearch 
            Caption         =   "��� ��ȸ"
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   1580
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� Ȯ��"
            Height          =   390
            Left            =   240
            TabIndex        =   50
            Top             =   270
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� �뷮 Ȯ��"
            Height          =   390
            Left            =   240
            TabIndex        =   49
            Top             =   705
            Width           =   1590
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �����̷�"
            Height          =   390
            Left            =   240
            TabIndex        =   48
            Top             =   2000
            Width           =   1590
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "������ Ȯ��"
            Height          =   390
            Left            =   240
            TabIndex        =   47
            Top             =   1150
            Width           =   1590
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   " �ΰ� ����"
         Height          =   2295
         Left            =   2520
         TabIndex        =   42
         Top             =   4200
         Width           =   4980
         Begin VB.CommandButton btnUpdateemailconfig 
            Caption         =   "�˸����� ���۸�� ����"
            Height          =   375
            Left            =   2520
            TabIndex        =   75
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton btnListemailconfig 
            Caption         =   "�˸����� ���۸�� ��ȸ"
            Height          =   375
            Left            =   2520
            TabIndex        =   74
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton btnAttachStatement 
            Caption         =   "���ڸ��� ÷��"
            Height          =   375
            Left            =   2520
            TabIndex        =   64
            Top             =   300
            Width           =   2295
         End
         Begin VB.CommandButton btnFAXSEnd 
            Caption         =   "���ѽ� ����"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   1640
            Width           =   2115
         End
         Begin VB.CommandButton btnSendFAX 
            Caption         =   "�ѽ� ����"
            Height          =   390
            Left            =   225
            TabIndex        =   45
            Top             =   1200
            Width           =   2115
         End
         Begin VB.CommandButton btnSendSMS 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   225
            TabIndex        =   44
            Top             =   735
            Width           =   2115
         End
         Begin VB.CommandButton btnSendEmail 
            Caption         =   "�̸��� ����"
            Height          =   390
            Left            =   225
            TabIndex        =   43
            Top             =   300
            Width           =   2115
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   " ��Ÿ URL "
         Height          =   1290
         Left            =   13920
         TabIndex        =   39
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton btnGetURL_PBOX 
            Caption         =   "���� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   41
            Top             =   705
            Width           =   1500
         End
         Begin VB.CommandButton btnGetURL_TBOX 
            Caption         =   "�ӽ� ������"
            Height          =   390
            Left            =   210
            TabIndex        =   40
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   " ����/�μ�"
         Height          =   2565
         Left            =   7920
         TabIndex        =   33
         Top             =   4200
         Width           =   5250
         Begin VB.CommandButton btnGetViewURL 
            Caption         =   "���ڸ��� ���� URL (�޴�, ��ưx)"
            Height          =   615
            Left            =   3120
            TabIndex        =   76
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "���ڸ��� ���ϸ�ũ URL"
            Height          =   390
            Left            =   210
            TabIndex        =   38
            Top             =   2040
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�뷮 �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   37
            Top             =   1570
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "���ڸ��� �μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���ڸ��� ���� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintURL 
            Caption         =   "������ �μ� �˾� URL"
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
         Caption         =   " ÷������ "
         Height          =   1335
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   4560
         Begin VB.CommandButton btnDeleteFile 
            Caption         =   "���� ����"
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
            Text            =   "���Ͼ��̵�"
            Top             =   840
            Width           =   2820
         End
         Begin VB.CommandButton btnGetFiles 
            Caption         =   "÷�� ���"
            Height          =   390
            Left            =   1800
            TabIndex        =   19
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton btnAttachFile 
            Caption         =   "���� ÷��"
            Height          =   390
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ӽ�����- ���� ���μ���"
         Height          =   2700
         Left            =   7800
         TabIndex        =   12
         Top             =   480
         Width           =   3510
         Begin VB.CommandButton btnIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   480
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   29
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton btnCancelIssue 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   480
            Left            =   360
            Style           =   1  '�׷���
            TabIndex        =   28
            Top             =   2055
            Width           =   975
         End
         Begin VB.CommandButton btnRegister 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���"
            Height          =   375
            Left            =   1305
            Style           =   1  '�׷���
            TabIndex        =   15
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnUpdate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            Height          =   375
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "����"
            Height          =   495
            Left            =   2280
            Style           =   1  '�׷���
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
            BackStyle       =   0  '����
            Caption         =   "�ӽ�����"
            Height          =   180
            Left            =   465
            TabIndex        =   16
            Top             =   555
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
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
         Text            =   "�ŷ�����"
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton btnCheckMgtKeyInUse 
         Caption         =   "������ȣ ��뿩�� Ȯ��"
         Height          =   375
         Left            =   2565
         TabIndex        =   10
         Top             =   1830
         Width           =   2190
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "���� ���� : "
         Height          =   180
         Left            =   1530
         TabIndex        =   32
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����ڵ�( FormCode ) : "
         Height          =   180
         Left            =   615
         TabIndex        =   22
         Top             =   1050
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "������ȣ( MgtKey) : "
         Height          =   180
         Left            =   915
         TabIndex        =   9
         Top             =   1455
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2730
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   16755
      Begin VB.Frame Frame16 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2370
         Left            =   14040
         TabIndex        =   69
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "�������� ����Ʈ"
         Height          =   2370
         Left            =   11760
         TabIndex        =   68
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   360
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ȸ������ ���� "
         Height          =   2370
         Left            =   9600
         TabIndex        =   55
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   2370
         Left            =   6840
         TabIndex        =   26
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton btnGetSealURL 
            Caption         =   "�ΰ� �� ÷�ι��� ��� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   "�˺� �α��� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ����� ���� "
         Height          =   2370
         Left            =   4800
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ���� "
         Height          =   2370
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   2625
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������ "
         Height          =   2370
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   360
            Left            =   75
            TabIndex        =   51
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
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
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
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
' �˺� ���ڸ��� API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2021-10-07
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================

Option Explicit

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'���ڸ��� ���� ��ü ����
Private statementService As New PBDocService

Private Function selectedItemCode() As Integer
    selectedItemCode = 121
 
    If cboItemCode.Text = "�ŷ�����" Then selectedItemCode = 121
    If cboItemCode.Text = "û����" Then selectedItemCode = 122
    If cboItemCode.Text = "������" Then selectedItemCode = 123
    If cboItemCode.Text = "���ּ�" Then selectedItemCode = 124
    If cboItemCode.Text = "�Ա�ǥ" Then selectedItemCode = 125
    If cboItemCode.Text = "������" Then selectedItemCode = 126
    
End Function

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/statement/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = ""
    
    Set info = statementService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� ������ ���ڸ��� 1���� �� ���� ������(����Ʈ ���, ���� �޴� �� ��ư ����)�� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetViewURL
'=========================================================================
Private Sub btnGetViewURL_Click()
    Dim url As String
  
    url = statementService.GetViewURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
    
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/statement/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.corpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.bizType = "����"
    
    '����, �ִ� 100��
    joinData.bizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = statementService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ��� ����� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = statementService.GetUnitCost(txtCorpNum.Text, selectedItemCode)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� ���ڸ��� API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = statementService.GetChargeInfo(txtCorpNum.Text, selectedItemCode)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
        
    url = statementService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �ΰ� �� ÷�ι��� ��� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
'=========================================================================
Private Sub btnGetSealURL_Click()

    Dim url As String
    
    url = statementService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/statement/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "VB6STATE_01"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
        
    Set Response = statementService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = statementService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/statement/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = statementService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = statementService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
' - https://docs.popbill.com/statement/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = statementService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = statementService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = statementService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = statementService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)�� �̿��Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(statementService.LastErrCode) + "] " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = statementService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ��Ʈ�ʰ� ���ڸ��� ���� �������� �Ҵ��ϴ� ������ȣ�� ��뿩�θ� Ȯ���մϴ�.
' - �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
' - https://docs.popbill.com/statement/vb/api#CheckMgtKeyInUse
'=========================================================================
Private Sub btnCheckMgtKeyInUse_Click()
    Dim Response As PBResponse
   
    Set Response = statementService.CheckMgtKeyInUse(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ۼ��� ���ڸ��� �����͸� �˺��� ����� ���ÿ� �����Ͽ�, "����Ϸ�" ���·� ó���մϴ�.
' - �˺� ����Ʈ [���ڸ���] > [ȯ�漳��] > [���ڸ��� ����] �޴��� ����� �ڵ����� �ɼ� ������ ���� ���ڸ����� "����Ϸ�" ���°� �ƴ� "���δ��" ���·� ���� ó�� �� �� �ֽ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#RegistIssue
'=========================================================================
Private Sub btnRegistIssue_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    Dim emailSubject As String
    
    Statement.memo = "��ù��� �޸�"
    
    '[�ʼ�] ����� �ۼ�����, ��������(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    

    '=========================================================================
    '                               �߽��� ����
    '=========================================================================
    
    '�߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '�߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '�߽��� ��ȣ
    Statement.senderCorpName = "�߽��� ��ȣ"
    
    '�߽��� ��ȣ��
    Statement.senderCEOName = "�߽��� ��ǥ�� ����"
    
    '�߽��� �ּ�
    Statement.senderAddr = "�߽��� �ּ�"
    
    '�߽��� ����
    Statement.senderBizClass = "�߽��� ����"
    
    '�߽��� ����
    Statement.senderBizType = "�߽��� ����,����2"
    
    '�߽��� ����ڼ���
    Statement.senderContactName = "�߽��� ����ڸ�"
    
    '�߽��� �̸���
    Statement.senderEmail = "test@test.com"
    
    '�߽��� ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '�߽��� �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '������ ��ȣ
    Statement.receiverCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.receiverCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.receiverAddr = "������ �ּ�"
    
    '������ ����
    Statement.receiverBizClass = "������ ���� "
    
    '������ ����
    Statement.receiverBizType = "������ ����"
    
    '������ ����ڸ�
    Statement.receiverContactName = "������ ����ڸ�"
    
    '������ �����ּ�
    Statement.receiverEmail = "test@test.com"
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"
        
    '���� �� �Ϸù�ȣ �׸�
    Statement.serialNum = "123"
    
    '���� �� ��� �׸�
    Statement.remark1 = "���1"
    Statement.remark2 = "���2"
    Statement.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Statement.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Statement.bankBookYN = False
    
    '����� �˸����� �߼ۿ���
    Statement.smssendYN = True
    
    '���׸� �߰�. (�迭 ���� ���� ����)
    '�Ϸù�ȣ(serialNum)�� 1���� ���������� �����Ͻñ� �ٶ��ϴ�
    Set Statement.detailList = New Collection
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20210902"       '�ŷ�����(yyyyMMdd)
        newDetail.itemName = "ǰ��" + CStr(i)   'ǰ���
        newDetail.spec = "�԰�"                 '�԰�
        newDetail.unit = "����"                 '����
        newDetail.qty = "1" '����               '�Ҽ��� 2�ڸ�����
        newDetail.unitCost = "100000"           '�Ҽ��� 2�ڸ�����
        newDetail.supplyCost = "100000"         '���ް���
        newDetail.tax = "10000"                 '����
        newDetail.remark = "���"               '���
        newDetail.spare1 = "spare1"             '����1
        newDetail.spare2 = "spare2"             '����2
        newDetail.spare3 = "spare3"             '����3
        newDetail.spare4 = "spare4"             '����4
        newDetail.spare5 = "spare5"             '����5
        
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '���ܾ�
    Statement.propertyBag.Add "Deposit", "10000"     '�Աݾ�
    Statement.propertyBag.Add "Balance", "100000"    '���ܾ�
    
    '�ȳ����� ����, �̱���� �⺻������� ����.
    emailSubject = ""
    
    Set Response = statementService.RegistIssue(txtCorpNum.Text, Statement, txtUserID.Text, emailSubject)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message + vbCrLf + "�˺� ���ι�ȣ : " + Response.invoiceNum)
End Sub

'=========================================================================
' �߽��ڰ� ������ ���ڸ����� ��������մϴ�.
' - https://docs.popbill.com/statement/vb/api#Cancel
'=========================================================================
Private Sub btnCancelIssue_sub_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "���ڸ��� ������� �޸�"
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ڸ����� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "���", "���ΰź�", "�������"
' - ���ڸ����� �����ϸ� ���� ������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#Delete
'=========================================================================
Private Sub btnDelete_sub_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ۼ��� ���ڸ��� �����͸� �˺��� �����մϴ�.
' - https://docs.popbill.com/statement/vb/api#Register
'=========================================================================
Private Sub btnRegister_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[�ʼ�] ����� �ۼ�����, ��������(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               �߽��� ����
    '=========================================================================
    
    '�߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '�߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '�߽��� ��ȣ
    Statement.senderCorpName = "�߽��� ��ȣ"
    
    '�߽��� ��ȣ��
    Statement.senderCEOName = "�߽��� ��ǥ�� ����"
    
    '�߽��� �ּ�
    Statement.senderAddr = "�߽��� �ּ�"
    
    '�߽��� ����
    Statement.senderBizClass = "�߽��� ����"
    
    '�߽��� ����
    Statement.senderBizType = "�߽��� ����,����2"
    
    '�߽��� ����ڼ���
    Statement.senderContactName = "�߽��� ����ڸ�"
    
    '�߽��� �̸���
    Statement.senderEmail = "test@test.com"
    
    '�߽��� ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '�߽��� �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '������ ��ȣ
    Statement.receiverCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.receiverCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.receiverAddr = "������ �ּ�"
    
    '������ ����
    Statement.receiverBizClass = "������ ���� "
    
    '������ ����
    Statement.receiverBizType = "������ ����"
    
    '������ ����ڸ�
    Statement.receiverContactName = "������ ����ڸ�"
    
    '������ �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"
        
    '���� �� �Ϸù�ȣ �׸�
    Statement.serialNum = "123"
    
    '���� �� ��� �׸�
    Statement.remark1 = "���1"
    Statement.remark2 = "���2"
    Statement.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Statement.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Statement.bankBookYN = False
    
    '����� �˸����� �߼ۿ���
    Statement.smssendYN = True
  
    '���׸� �߰�. (�迭 ���� ���� ����)
    '�Ϸù�ȣ(serialNum)�� 1���� ���������� �����Ͻñ� �ٶ��ϴ�
    Set Statement.detailList = New Collection

    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20210902"       '�ŷ�����(yyyyMMdd)
        newDetail.itemName = "ǰ��" + CStr(i)   'ǰ���
        newDetail.spec = "�԰�"                 '�԰�
        newDetail.unit = "����"                 '����
        newDetail.qty = "1" '����               '�Ҽ��� 2�ڸ�����
        newDetail.unitCost = "100000"           '�Ҽ��� 2�ڸ�����
        newDetail.supplyCost = "100000"         '���ް���
        newDetail.tax = "10000"                 '����
        newDetail.remark = "���"               '���
        newDetail.spare1 = "spare1"             '����1
        newDetail.spare2 = "spare2"             '����2
        newDetail.spare3 = "spare3"             '����3
        newDetail.spare4 = "spare4"             '����4
        newDetail.spare5 = "spare5"             '����5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '���ܾ�
    Statement.propertyBag.Add "Deposit", "10000"     '�Աݾ�
    Statement.propertyBag.Add "Balance", "100000"    '���ܾ�
    
    
    Set Response = statementService.Register(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ڸ����� �����մϴ�.���� ���ڸ����� [����]�մϴ�.' 1���� ���ڸ����� �����մϴ�.
' - https://docs.popbill.com/statement/vb/api#Update
'=========================================================================
Private Sub btnUpdate_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[�ʼ�] ����� �ۼ�����, ��������(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               �߽��� ����
    '=========================================================================
    
    '�߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '�߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '�߽��� ��ȣ
    Statement.senderCorpName = "�߽��� ��ȣ_����"
    
    '�߽��� ��ȣ��
    Statement.senderCEOName = "�߽��� ��ǥ�� ����"
    
    '�߽��� �ּ�
    Statement.senderAddr = "�߽��� �ּ�_����"
    
    '�߽��� ����
    Statement.senderBizClass = "�߽��� ����_����"
    
    '�߽��� ����
    Statement.senderBizType = "�߽��� ����,����2"
    
    '�߽��� ����ڼ���
    Statement.senderContactName = "�߽��� ����ڸ�"
    
    '�߽��� �̸���
    Statement.senderEmail = "test@test.com"
    
    '�߽��� ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '�߽��� �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '������ ��ȣ
    Statement.receiverCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.receiverCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.receiverAddr = "������ �ּ�"
    
    '������ ����
    Statement.receiverBizClass = "������ ���� "
    
    '������ ����
    Statement.receiverBizType = "������ ����"
    
    '������ ����ڸ�
    Statement.receiverContactName = "������ ����ڸ�"
    
    '������ �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"
        
    '���� �� �Ϸù�ȣ �׸�
    Statement.serialNum = "123"
    
    '���� �� ��� �׸�
    Statement.remark1 = "���1"
    Statement.remark2 = "���2"
    Statement.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Statement.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Statement.bankBookYN = False
    
    '����� �˸����� �߼ۿ���
    Statement.smssendYN = True
    
    '���׸� �߰�. (�迭 ���� ���� ����)
    '�Ϸù�ȣ(serialNum)�� 1���� ���������� �����Ͻñ� �ٶ��ϴ�
    Set Statement.detailList = New Collection
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20210902"       '�ŷ�����(yyyyMMdd)
        newDetail.itemName = "ǰ��" + CStr(i)   'ǰ���
        newDetail.spec = "�԰�"                 '�԰�
        newDetail.unit = "����"                 '����
        newDetail.qty = "1" '����               '�Ҽ��� 2�ڸ�����
        newDetail.unitCost = "100000"           '�Ҽ��� 2�ڸ�����
        newDetail.supplyCost = "100000"         '���ް���
        newDetail.tax = "10000"                 '����
        newDetail.remark = "���"               '���
        newDetail.spare1 = "spare1"             '����1
        newDetail.spare2 = "spare2"             '����2
        newDetail.spare3 = "spare3"             '����3
        newDetail.spare4 = "spare4"             '����4
        newDetail.spare5 = "spare5"             '����5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '���ܾ�
    Statement.propertyBag.Add "Deposit", "10000"     '�Աݾ�
    Statement.propertyBag.Add "Balance", "100000"    '���ܾ�
    
    Set Response = statementService.Update(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ڸ����� �����Ͽ�, "����Ϸ�" ���·� ó���մϴ�.
' - �˺� ����Ʈ [���ڸ���] > [ȯ�漳��] > [���ڸ��� ����] �޴��� ����� �ڵ����� �ɼ� ������ ���� ���ڸ����� "����Ϸ�" ���°� �ƴ� "���δ��" ���·� ���� ó�� �� �� �ֽ��ϴ�.
' - ���ڸ��� ���� �Լ� ȣ��� ����Ʈ�� ���ݵǸ�, �����ڿ��� ���� �ȳ� ������ �߼۵˴ϴ�.
' - https://docs.popbill.com/statement/vb/api#StmIssue
'=========================================================================
Private Sub btnIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "���ڸ��� ���� �޸�"
    
    Set Response = statementService.Issue(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �߽��ڰ� ������ ���ڸ����� ��������մϴ�.
' - https://docs.popbill.com/statement/vb/api#Cancel
'=========================================================================
Private Sub btnCancelIssue_Click()
    Dim Response As PBResponse
    Dim memo As String
    
    '�޸�
    memo = "���� ��� �޸�"
    
    Set Response = statementService.Cancel(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, memo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���� ������ ������ ���ڸ����� �����մϴ�.
' - ���� ������ ����: "�ӽ�����", "���", "���ΰź�", "�������"
' - ���ڸ����� �����ϸ� ���� ������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#Delete
'=========================================================================
Private Sub btnDelete_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' "�ӽ�����" ������ ������ 1���� ������ ÷���մϴ�. (�ִ� 5��)
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
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)

End Sub

'=========================================================================
' ���ڸ����� ÷�ε� ���ϸ���� Ȯ���մϴ�.
' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ���ϻ���(DeleteFile API) ȣ��� �̿��� �� �ֽ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#GetFiles
'=========================================================================
Private Sub btnGetFiles_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim file As PBAttachFile
    
    Set resultList = statementService.GetFiles(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "serialNum(�Ϸù�ȣ) | attachedfile(���Ͼ��̵�) | displayName(÷�����ϸ�) |  RegDT(÷���Ͻ�)" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' "�ӽ�����" ������ ���ڸ����� ÷�ε� 1���� ������ �����մϴ�.
' - ������ �ĺ��ϴ� ���Ͼ��̵�� ÷������ ���(GetFiles API) �� �����׸� �� ���Ͼ��̵�(AttachedFile) ���� ���� Ȯ���� �� �ֽ��ϴ�.
' - https://docs.popbill.com/statement/vb/api#DeleteFile
'=========================================================================
Private Sub btnDeleteFile_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.DeleteFile(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, txtFileID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ����� 1���� ���� �� ������� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetInfo
'=========================================================================
Private Sub btnGetInfo_Click()
    Dim docInfo As PBDocInfo
    Dim tmp As String
    
    Set docInfo = statementService.GetInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemCode (���� �ڵ�) :" + CStr(docInfo.itemCode) + vbCrLf
    tmp = tmp + "itemKey (�˺���ȣ) : " + docInfo.itemKey + vbCrLf
    tmp = tmp + "invoiceNum (�˺� ���ι�ȣ) : " + docInfo.invoiceNum + vbCrLf
    tmp = tmp + "mgtKey (������ȣ) : " + docInfo.mgtKey + vbCrLf
    tmp = tmp + "taxType (��������) : " + docInfo.taxType + vbCrLf
    tmp = tmp + "writeDate (�ۼ�����) : " + docInfo.writeDate + vbCrLf
    tmp = tmp + "regDT (�ӽ������Ͻ�) : " + docInfo.regDT + vbCrLf
    tmp = tmp + "senderCorpName (�߽��� ��ȣ) : " + docInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCorpNum (�߽��� ����ڹ�ȣ) : " + docInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderPrintYN (�߽��� �μ⿩��) :" + CStr(docInfo.senderPrintYN) + vbCrLf
    tmp = tmp + "receiverCorpName (������ ��ȣ) : " + docInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCorpNum (������ ����ڹ�ȣ) : " + docInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverPrintYN (������ �μ⿩��) :" + CStr(docInfo.receiverPrintYN) + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + docInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + docInfo.taxTotal + vbCrLf
    tmp = tmp + "purposeType (����/û��) : " + docInfo.purposeType + vbCrLf
    tmp = tmp + "issueDT (�����Ͻ�) : " + docInfo.issueDT + vbCrLf
    tmp = tmp + "stateCode (�����ڵ�) :" + CStr(docInfo.stateCode) + vbCrLf
    tmp = tmp + "stateDT (���� �����Ͻ�) : " + docInfo.stateDT + vbCrLf
    tmp = tmp + "stateMemo (���¸޸�) : " + docInfo.stateMemo + vbCrLf
    tmp = tmp + "openYN (���� ����) :" + CStr(docInfo.openYN) + vbCrLf
    tmp = tmp + "openDT (���� �Ͻ�) : " + docInfo.openDT + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �ټ����� ���ڸ��� ���� �� ������� Ȯ���մϴ�. (1ȸ ȣ�� �� �ִ� 1,000�� Ȯ�� ����)
' - https://docs.popbill.com/statement/vb/api#GetInfos
'=========================================================================
Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBDocInfo
    
    '���ڸ��� ������ȣ �迭 (�ִ� 1000��)
    KeyList.Add "20150113-01"
    KeyList.Add "20150113-02"
    KeyList.Add "20150113-03"
    KeyList.Add "20150113-04"
    
    Set resultList = statementService.GetInfos(txtCorpNum.Text, selectedItemCode, KeyList)
            
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
        
    tmp = "itemCode(���� �ڵ�) | itemKey(�˺���ȣ) | invoiceNum(�˺� ���ι�ȣ) | mgtKey(������ȣ) | taxType(��������) | " + vbCrLf
    tmp = tmp + "writeDate(�ۼ�����) | regDT(�ӽ������Ͻ�) | senderCorpName(�߽��� ��ȣ) | senderCorpNum(�߽��� ����ڹ�ȣ) | " + vbCrLf
    tmp = tmp + "senderPrintYN(�߽��� �μ⿩��) | receiverCorpName(������ ��ȣ) | receiverCorpNum(������ ����ڹ�ȣ) | " + vbCrLf
    tmp = tmp + "receiverPrintYN(������ �μ⿩��) | supplyCostTotal(���ް��� �հ�) | taxTotal(���� �հ�) | purposeType(����/û��) | " + vbCrLf
    tmp = tmp + "issueDT(�����Ͻ�) | stateCode(�����ڵ�) | stateDT(���� �����Ͻ�) | stateMemo(���¸޸�) | " + vbCrLf
    tmp = tmp + "openYN(���� ����) | openDT(���� �Ͻ�)" + vbCrLf + vbCrLf
        
    For Each info In resultList
        tmp = tmp + CStr(info.itemCode) + " | " + info.itemKey + " | " + info.invoiceNum + " | " + info.mgtKey + " | " + info.taxType + " | " + info.writeDate + " | "
        tmp = tmp + info.regDT + " | " + info.senderCorpName + " | " + info.senderCorpNum + CStr(info.senderPrintYN) + " | " + info.receiverCorpName + " | "
        tmp = tmp + info.receiverCorpNum + CStr(info.receiverPrintYN) + " | " + info.supplyCostTotal + " | " + info.taxTotal + " | " + info.purposeType + " | "
        tmp = tmp + info.issueDT + CStr(info.stateCode) + " | " + info.stateDT + " | " + info.stateMemo + CStr(info.openYN) + " | " + info.openDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ڸ��� 1���� ������ Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetDetailInfo
'=========================================================================
Private Sub btnGetDetailInfo_Click()
    Dim docDetailInfo As PBStatement
    Dim tmp As String
    Dim key
    Dim detail As PBDocDetail
    
    Set docDetailInfo = statementService.GetDetailInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docDetailInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "itemCode (�������� �ڵ�) : " + CStr(docDetailInfo.itemCode) + vbCrLf
    tmp = tmp + "mgtKey (������ȣ) : " + docDetailInfo.mgtKey + vbCrLf
    tmp = tmp + "invoiceNum (�˺� ���ι�ȣ) : " + docDetailInfo.invoiceNum + vbCrLf
    tmp = tmp + "formCode (������ �ڵ�) : " + docDetailInfo.formCode + vbCrLf
    tmp = tmp + "writeDate (�ۼ�����) : " + docDetailInfo.writeDate + vbCrLf
    tmp = tmp + "taxType (��������) : " + docDetailInfo.taxType + vbCrLf
    tmp = tmp + "purposeType (����/û��) : " + docDetailInfo.purposeType + vbCrLf
    tmp = tmp + "serialNum (�Ϸù�ȣ) : " + docDetailInfo.serialNum + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + docDetailInfo.taxTotal + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + docDetailInfo.supplyCostTotal + vbCrLf
    tmp = tmp + "totalAmount (�հ�ݾ�) : " + docDetailInfo.totalAmount + vbCrLf
    tmp = tmp + "remark1 (���1) : " + docDetailInfo.remark1 + vbCrLf
    tmp = tmp + "remark2 (���2) : " + docDetailInfo.remark2 + vbCrLf
    tmp = tmp + "remark3 (���3) : " + docDetailInfo.remark3 + vbCrLf
        
    tmp = tmp + "senderCorpNum (�߽��� ����ڹ�ȣ) : " + docDetailInfo.senderCorpNum + vbCrLf
    tmp = tmp + "senderTaxRegID (�߽��� ��������ȣ) : " + docDetailInfo.senderTaxRegID + vbCrLf
    tmp = tmp + "senderCorpName (�߽��� ��ȣ) : " + docDetailInfo.senderCorpName + vbCrLf
    tmp = tmp + "senderCEOName (�߽��� ��ǥ�ڼ���) : " + docDetailInfo.senderCEOName + vbCrLf
    tmp = tmp + "senderAddr (�߽��� �ּ�) : " + docDetailInfo.senderAddr + vbCrLf
    tmp = tmp + "senderBizClass (�߽��� ����) : " + docDetailInfo.senderBizClass + vbCrLf
    tmp = tmp + "senderBizType (�߽��� ����) : " + docDetailInfo.senderBizType + vbCrLf
    tmp = tmp + "senderContactName (�߽��� ����) : " + docDetailInfo.senderContactName + vbCrLf
    tmp = tmp + "senderDeptName (�߽��� �μ�) : " + docDetailInfo.senderDeptName + vbCrLf
    tmp = tmp + "senderTEL (�߽��� ����ó) : " + docDetailInfo.senderTEL + vbCrLf
    tmp = tmp + "senderHP (�߽��� �޴���ȭ) : " + docDetailInfo.senderHP + vbCrLf
    tmp = tmp + "senderEmail (�߽��� �̸���) : " + docDetailInfo.senderEmail + vbCrLf
    tmp = tmp + "senderFAX (�߽��� �ѽ�) : " + docDetailInfo.senderFAX + vbCrLf

    tmp = tmp + "receiverCorpNum (������ ����ڹ�ȣ) : " + docDetailInfo.receiverCorpNum + vbCrLf
    tmp = tmp + "receiverTaxRegID (������ ��������ȣ) : " + docDetailInfo.receiverTaxRegID + vbCrLf
    tmp = tmp + "receiverCorpName (������ ��ȣ) : " + docDetailInfo.receiverCorpName + vbCrLf
    tmp = tmp + "receiverCEOName (������ ��ǥ�ڼ���) : " + docDetailInfo.receiverCEOName + vbCrLf
    tmp = tmp + "receiverAddr (������ �ּ�) : " + docDetailInfo.receiverAddr + vbCrLf
    tmp = tmp + "receiverBizClass (������ ����) : " + docDetailInfo.receiverBizClass + vbCrLf
    tmp = tmp + "receiverBizType (������ ����) : " + docDetailInfo.receiverBizType + vbCrLf
    tmp = tmp + "receiverContactName (������ ����) : " + docDetailInfo.receiverContactName + vbCrLf
    tmp = tmp + "receiverDeptName (������ �μ�) : " + docDetailInfo.receiverDeptName + vbCrLf
    tmp = tmp + "receiverTEL (������ ����ó) : " + docDetailInfo.receiverTEL + vbCrLf
    tmp = tmp + "receiverHP (������ �޴���ȭ) : " + docDetailInfo.receiverHP + vbCrLf
    tmp = tmp + "receiverEmail (������ �̸���) : " + docDetailInfo.receiverEmail + vbCrLf
    tmp = tmp + "receiverFAX (������ �ѽ�) : " + docDetailInfo.receiverFAX + vbCrLf
        
    tmp = tmp + "detailList (���׸�)" + vbCrLf
    tmp = tmp + "serialNum(�Ϸù�ȣ) | purchaseDT(�ŷ�����) | itemName(ǰ���) | spec(�԰�) | qty(����) |"
    tmp = tmp + "unitCost(�ܰ�) | supplyCost(���ް���) | tax(����) | remark(���) | spare1(����1) "
    tmp = tmp + "spare2(����2) | spare3(����3) | spare4(����4) | spare5(����5) "
    For Each detail In docDetailInfo.detailList
        tmp = tmp + vbTab + CStr(detail.serialNum) + " : " + detail.purchaseDT + " | " + detail.itemName + " | "
        tmp = tmp + detail.spec + " | " + detail.qty + " | " + " | " + detail.unitCost + " | "
        tmp = tmp + detail.supplyCost + " | " + detail.tax + " | " + " | " + detail.remark + " | "
        tmp = tmp + detail.spare1 + " | " + detail.spare2 + " | " + " | " + detail.spare3 + " | "
        tmp = tmp + detail.spare4 + " | " + detail.spare5 + vbCrLf
    Next
    
    tmp = tmp + "Properties (�߰��Ӽ�)" + vbCrLf
    For Each key In docDetailInfo.propertyBag.keys
        tmp = tmp + vbTab + key + " : " + docDetailInfo.propertyBag.Item(key) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' �˻����ǿ� �ش��ϴ� ���ڸ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
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
    
    '[�ʼ�] ��������, R-����Ͻ� W-�ۼ����� I-�����Ͻ� �� ��1
    DType = "W"
    
    '[�ʼ�] ��������, yyyyMMdd
    SDate = "20210901"
    
    '[�ʼ�] ��������, yyyyMMdd
    EDate = "20210910"
    
    '���ۻ��°� �迭, �̱���� ��ü������ȸ, �������°� 3�ڸ����� �ۼ� 2,3��° ���ϵ�ī�� ����
    '�����ڵ忡 ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] > 5.1 ���ڸ��� �����ڵ�" �� �����Ͻñ� �ٶ��ϴ�.
    state.Add "100"
    state.Add "2**"
    state.Add "3**"
    
    '���� �ڵ� �迭, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ, 126-������
    itemCode.Add "121"
    itemCode.Add "122"
    itemCode.Add "123"
    itemCode.Add "124"
    itemCode.Add "125"
    itemCode.Add "126"
    
    '������ ��ȣ, �⺻�� '1'
    Page = 1
    
    '�������� �˻�����, �⺻�� '500', �ִ� '1000'
    PerPage = 10
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '�ŷ�ó ����, �ŷ�ó ��ȣ �Ǵ� �ŷ�ó ����ڵ�Ϲ�ȣ ����, �̱���� ��ü��ȸ
    QString = ""
    
    Set docSearchList = statementService.Search(txtCorpNum.Text, DType, SDate, EDate, state, itemCode, Page, PerPage, Order, QString)
     
    If docSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code (�����ڵ�) : " + CStr(docSearchList.code) + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(docSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(docSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(docSearchList.pageNum) + vbCrLf
    tmp = tmp + "perCount (������ ����) : " + CStr(docSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + docSearchList.message + vbCrLf + vbCrLf
    

    tmp = tmp + "itemCode(���� �ڵ�) | itemKey(�˺���ȣ) | invoiceNum(�˺� ���ι�ȣ) | mgtKey(������ȣ) | taxType(��������) | " + vbCrLf
    tmp = tmp + "writeDate(�ۼ�����) | regDT(�ӽ������Ͻ�) | senderCorpName(�߽��� ��ȣ) | senderCorpNum(�߽��� ����ڹ�ȣ) | " + vbCrLf
    tmp = tmp + "senderPrintYN(�߽��� �μ⿩��) | receiverCorpName(������ ��ȣ) | receiverCorpNum(������ ����ڹ�ȣ) | " + vbCrLf
    tmp = tmp + "receiverPrintYN(������ �μ⿩��) | supplyCostTotal(���ް��� �հ�) | taxTotal(���� �հ�) | purposeType(����/û��) | " + vbCrLf
    tmp = tmp + "issueDT(�����Ͻ�) | stateCode(�����ڵ�) | stateDT(���� �����Ͻ�) | stateMemo(���¸޸�) | " + vbCrLf
    tmp = tmp + "openYN(���� ����) | openDT(���� �Ͻ�)" + vbCrLf + vbCrLf

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
' ���ڸ����� ���¿� ���� �����̷��� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetLogs
'=========================================================================
Private Sub btnGetLogs_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim log As PBDocLog
    
    Set resultList = statementService.GetLogs(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "DocLogType(�α�Ÿ��) | Log(�̷�����) | ProcType(ó������) | ProcCorpName(ó��ȸ���) | ProcMemo(ó���޸�) | RegDT(����Ͻ�) | IP(������)" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procCorpName + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' "���δ��", "����Ϸ�" ������ ���ڸ����� ���õ� ���� �ȳ� ������ ������ �մϴ�.
' - https://docs.popbill.com/statement/vb/api#SendEmail
'=========================================================================
Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
  
    '������ �����ּ�
    receiverEmail = "test@test.com"
    
    Set Response = statementService.SendEmail(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ����� ���õ� �ȳ� SMS(�ܹ�) ���ڸ� �������ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [����] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �޽����� �ִ� 90byte���� �Է� �����ϰ�, �ʰ��� ������ �ڵ����� �����Ǿ� �����մϴ�. (�ѱ� �ִ� 45��)
' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
' - https://docs.popbill.com/statement/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    Dim Contents As String
    
    '�߽Ź�ȣ
    senderNum = "070-4304-2991"
    
    '���Ź�ȣ
    receiverNum = "010-111-222"
    
    ' �޽��� ����, �ִ� 90Byte (�ѱ� 45��), ���̸� �ʰ��� ������ �����Ǿ� ���۵˴ϴ�.
    Contents = "���ڸ����� �����Ͽ����ϴ�. ������ Ȯ���Ͽ� �ֽñ�ٶ��ϴ�"
    
    Set Response = statementService.SendSMS(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ����� �ѽ��� �����ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
' - https://docs.popbill.com/statement/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    
    '�߽��� ��ȣ
    senderNum = "070-4304-2991"
    
    '������ �ѽ���ȣ
    receiverNum = "070-111-222"
  
    Set Response = statementService.SendFax(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ����� �ѽ��� �����ϴ� �Լ���, �˺��� �����͸� �����ϴ� ������ �����ϴ�.
' - �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
' - �ѽ� ���� ��û�� �ۼ��� ������ȣ�� �ѽ����� ���ϸ����� ���˴ϴ�.
' - �ѽ� ���۰���� Ȯ���ϱ� ���ؼ��� ���ѽ� ���� ��û �� ��ȯ���� ������ȣ�� �̿��Ͽ� �ѽ� API�� ���۰�� Ȯ�� (GetFaxDetail) API�� �̿��ϸ� �˴ϴ�.
' - https://docs.popbill.com/statement/vb/api#FAXSend
'=========================================================================
Private Sub btnFAXSend_Click()
    Dim Statement As New PBStatement
    Dim ReceiptNum As String
    Dim newDetail As PBDocDetail
    Dim i
    
    '�ѽ� �߽Ź�ȣ
    Statement.sendNum = "07043042991"
    
    '�ѽ� ���Ź�ȣ
    Statement.receiveNum = "070111222"
       
    '[�ʼ�] ����� �ۼ�����, ��������(yyyyMMdd)
    Statement.writeDate = "20210902"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ������ȣ, �ִ� 24�ڸ�, ����, ���� '-', '_'�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               �߽��� ����
    '=========================================================================
    
    '�߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '�߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '�߽��� ��ȣ
    Statement.senderCorpName = "�߽��� ��ȣ"
    
    '�߽��� ��ȣ��
    Statement.senderCEOName = "�߽��� ��ǥ�� ����"
    
    '�߽��� �ּ�
    Statement.senderAddr = "�߽��� �ּ�"
    
    '�߽��� ����
    Statement.senderBizClass = "�߽��� ����"
    
    '�߽��� ����
    Statement.senderBizType = "�߽��� ����,����2"
    
    '�߽��� ����ڼ���
    Statement.senderContactName = "�߽��� ����ڸ�"
    
    '�߽��� �̸���
    Statement.senderEmail = "test@test.com"
    
    '�߽��� ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '�߽��� �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '������ ��ȣ
    Statement.receiverCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.receiverCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.receiverAddr = "������ �ּ�"
    
    '������ ����
    Statement.receiverBizClass = "������ ���� "
    
    '������ ����
    Statement.receiverBizType = "������ ����"
    
    '������ ����ڸ�
    Statement.receiverContactName = "������ ����ڸ�"
    
    '������ �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"
        
    '���� �� �Ϸù�ȣ �׸�
    Statement.serialNum = "123"
    
    '���� �� ��� �׸�
    Statement.remark1 = "���1"
    Statement.remark2 = "���2"
    Statement.remark3 = "���3"
    
    '����ڵ���� �̹��� ÷�ο���
    Statement.businessLicenseYN = False
    
    '����纻 �̹��� ÷�ο���
    Statement.bankBookYN = False
    
    '����� �˸����� �߼ۿ���
    Statement.smssendYN = True
  
    '���׸� �߰�.
    Set Statement.detailList = New Collection
    
    For i = 1 To 5
        Set newDetail = New PBDocDetail
        newDetail.serialNum = i                 '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20210902"       '�ŷ�����(yyyyMMdd)
        newDetail.itemName = "ǰ��" + CStr(i)   'ǰ���
        newDetail.spec = "�԰�"                 '�԰�
        newDetail.unit = "����"                 '����
        newDetail.qty = "1" '����               '�Ҽ��� 2�ڸ�����
        newDetail.unitCost = "100000"           '�Ҽ��� 2�ڸ�����
        newDetail.supplyCost = "100000"         '���ް���
        newDetail.tax = "10000"                 '����
        newDetail.remark = "���"               '���
        newDetail.spare1 = "spare1"             '����1
        newDetail.spare2 = "spare2"             '����2
        newDetail.spare3 = "spare3"             '����3
        newDetail.spare4 = "spare4"             '����4
        newDetail.spare5 = "spare5"             '����5
        Statement.detailList.Add newDetail
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = CreateObject("Scripting.Dictionary")
    
    Statement.propertyBag.Add "CBalance", "100000"   '���ܾ�
    Statement.propertyBag.Add "Deposit", "10000"     '�Աݾ�
    Statement.propertyBag.Add "Balance", "100000"    '���ܾ�
    
    ReceiptNum = statementService.FAXSend(txtCorpNum.Text, Statement)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + ReceiptNum
End Sub

'=========================================================================
' �ϳ��� ���ڸ����� �ٸ� ���ڸ����� ÷���մϴ�.
' - https://docs.popbill.com/statement/vb/api#AttachStatement
'=========================================================================
Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ, 126-������
    SubItemCode = 121
    
    '÷���� ���ڸ��� ������ȣ
    SubMgtKey = "20210902-01"
    
    Set Response = statementService.AttachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ϳ��� ���ڸ����� ÷�ε� �ٸ� ���ڸ����� �����մϴ�.
' - https://docs.popbill.com/statement/vb/api#DetachStatement
'=========================================================================
Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '÷�������� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ, 126-������
    SubItemCode = 121
    
    '÷�������� ���ڸ��� ������ȣ
    SubMgtKey = "20210902-01"
      
    Set Response = statementService.DetachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ��� ���� ���� �׸� ���� �߼ۼ����� Ȯ���մϴ�.
' - https://docs.popbill.com/statement/vb/api#ListEmailConfig
'=========================================================================
Private Sub btnListemailconfig_Click()
    Dim resultList As Collection
    Dim i As Integer
    
    Set resultList = statementService.ListEmailConfig(txtCorpNum.Text, txtUserID.Text)
    
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
 
    Dim tmp As String
    
    tmp = "������������(EmailType) | ���ۿ���(SendYN) " + vbCrLf
    
    Dim info As PBEmailConfig
    
    For i = 1 To resultList.Count
        If resultList(i).emailType = "SMT_ISSUE" Then
            tmp = tmp + "�����ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_ACCEPT" Then
            tmp = tmp + "�߽��ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_DENY" Then
            tmp = tmp + "�߽��ڿ��� ���ڸ����� �ź� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL" Then
            tmp = tmp + "�����ڿ��� ���ڸ����� ��� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL_ISSUE" Then
            tmp = tmp + "�����ڿ��� ���ڸ����� ������� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' ���ڸ��� ���� ���� �׸� ���� �߼ۼ����� �����մϴ�.
' - https://docs.popbill.com/statement/vb/api#UpdateEmailConfig
'
' ������������
' SMT_ISSUE : ���޹޴��ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' SMT_ACCEPT : �����ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
' SMT_DENY : �����ڿ��� ���ڸ����� �ź� �Ǿ����� �˷��ִ� �����Դϴ�.
' SMT_CANCEL : ���޹޴��ڿ��� ���ڸ����� ��� �Ǿ����� �˷��ִ� �����Դϴ�.
' SMT_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڸ����� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
'=========================================================================
Private Sub btnUpdateemailconfig_Click()
    Dim Response As PBResponse
    Dim emailType As String
    Dim sendYN As Boolean
    
    '���� ���� ����
    emailType = "SMT_ISSUE"

    '���� ���� (True = ����, False = ������)
    sendYN = True
    
    Set Response = statementService.UpdateEmailConfig(txtCorpNum.Text, emailType, sendYN, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ����Ʈ�� ������ ���ڸ��� 1���� �� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetPopUpURL
'=========================================================================
Private Sub btnGetPopUpURL_Click()
    Dim url As String
  
    url = statementService.GetPopUpURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ���ڸ��� 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�ϸ�, ������������ �μ� �������� "������" / "���޹޴���" / "������+���޹޴���"�� �� �ϳ��� ������ �� �ֽ��ϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetPrintURL
'=========================================================================
Private Sub btnGetPrintURL_Click()
    Dim url As String
    
    url = statementService.GetPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' "���޹޴���" �� ���ڸ��� 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetEPrintURL
'=========================================================================
Private Sub btnGetEPrintUrl_Click()
    Dim url As String
    
    url = statementService.GetEPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �ټ����� ���ڸ����� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetMassPrintURL
'=========================================================================
Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    '���ڸ��� ������ȣ �迭 (�ִ� 100��)
    KeyList.Add "20210902-01"
    KeyList.Add "20210902-02"
    KeyList.Add "20210902-03"
    KeyList.Add "20210902-04"
    
    url = statementService.GetMassPrintURL(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �ȳ����ϰ� ���õ� ���ڸ����� Ȯ�� �� �� �ִ� �� �������� �˾� URL�� ��ȯ�ϸ�, �ش� URL�� ���� �ϴ��� �Ķ��� ��ư�� ��ũ�� �����ϴ�.
' - �Լ� ȣ��� ��ȯ ���� URL���� ��ȿ�ð��� �����ϴ�.
' - https://docs.popbill.com/statement/vb/api#GetMailURL
'=========================================================================
Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = statementService.GetMailURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڸ��� ���⹮���� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_PBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �α��� ���·� �˺� ����Ʈ�� ���ڸ��� �ӽù����� �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/statement/vb/api#GetURL
'=========================================================================
Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

Private Sub Form_Load()

    '���ڸ��� ��ü �ʱ�ȭ
    statementService.Initialize LinkID, SecretKey
    
    '����ȯ�漳����, True-���߿� False-�����
    statementService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-����
    statementService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, False-�̻��, �⺻��(False)
    statementService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� True-���, Fasle-�̻��, �⺻��(False)
    statementService.UseLocalTimeYN = False
    
    
End Sub

