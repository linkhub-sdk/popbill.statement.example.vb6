VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� ���ڸ��� SDK ����"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   17565
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnDetachStatement 
      Caption         =   "���ڸ��� ÷������"
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Top             =   8040
      Width           =   2235
   End
   Begin VB.Frame Frame7 
      Caption         =   " ���ڸ��� ���� ��� "
      Height          =   7380
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   13995
      Begin VB.Frame Frame9 
         Caption         =   "��ù��� ���μ���"
         Height          =   2655
         Left            =   5040
         TabIndex        =   58
         Top             =   480
         Width           =   2535
         Begin VB.CommandButton btnDelete_2 
            Caption         =   "����"
            Height          =   495
            Left            =   1560
            Style           =   1  '�׷���
            TabIndex        =   61
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton btnCancelISsue_2 
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
         Caption         =   " ���� ���� "
         Height          =   2535
         Left            =   240
         TabIndex        =   46
         Top             =   4200
         Width           =   2010
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� �����ȸ"
            Height          =   375
            Left            =   210
            TabIndex        =   63
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton btnGetInfo 
            Caption         =   "���� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   50
            Top             =   270
            Width           =   1590
         End
         Begin VB.CommandButton btnGetInfos 
            Caption         =   "���� ����(�뷮)"
            Height          =   390
            Left            =   210
            TabIndex        =   49
            Top             =   705
            Width           =   1590
         End
         Begin VB.CommandButton btnGetLogs 
            Caption         =   "���� �̷�"
            Height          =   390
            Left            =   210
            TabIndex        =   48
            Top             =   1140
            Width           =   1590
         End
         Begin VB.CommandButton btnGetDetailInfo 
            Caption         =   "���� �� ����"
            Height          =   390
            Left            =   210
            TabIndex        =   47
            Top             =   1590
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
            Width           =   2235
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
         Left            =   11400
         TabIndex        =   39
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton btnGetURL_SBOX 
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
         Caption         =   " ���� ���� "
         Height          =   2565
         Left            =   7920
         TabIndex        =   33
         Top             =   4200
         Width           =   3210
         Begin VB.CommandButton btnGetMailURL 
            Caption         =   "�̸���(���޹޴���) ��ũ URL"
            Height          =   390
            Left            =   195
            TabIndex        =   38
            Top             =   1590
            Width           =   2745
         End
         Begin VB.CommandButton btnGetMassPrintURL 
            Caption         =   "�ٷ� �μ� �˾� URL"
            Height          =   390
            Left            =   195
            TabIndex        =   37
            Top             =   1140
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPrintURL 
            Caption         =   "�μ� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   36
            Top             =   705
            Width           =   2745
         End
         Begin VB.CommandButton btnGetPopUpURL 
            Caption         =   "���� ���� ���� �˾� URL"
            Height          =   390
            Left            =   210
            TabIndex        =   35
            Top             =   270
            Width           =   2745
         End
         Begin VB.CommandButton btnGetEPrintURL 
            Caption         =   "������ �μ� �˾� URL"
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
            Height          =   525
            Left            =   322
            Style           =   1  '�׷���
            TabIndex        =   29
            Top             =   1365
            Width           =   1020
         End
         Begin VB.CommandButton btnCancel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�������"
            Height          =   480
            Left            =   285
            Style           =   1  '�׷���
            TabIndex        =   28
            Top             =   2055
            Width           =   1095
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
            Left            =   2355
            Style           =   1  '�׷���
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
      Height          =   2250
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   16755
      Begin VB.Frame Frame16 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1695
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
         Height          =   1695
         Left            =   11760
         TabIndex        =   68
         Top             =   240
         Width           =   2175
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
            Width           =   1905
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " ȸ������ ���� "
         Height          =   1695
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
         Height          =   1770
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
         Height          =   1770
         Left            =   4800
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ���� "
         Height          =   1770
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   2625
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "��� �ܰ� Ȯ��"
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������ "
         Height          =   1770
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1635
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   360
            Left            =   75
            TabIndex        =   51
            Top             =   735
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
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
' - VB6 SDK ����ȯ�� ������� �ȳ� :
' - ������Ʈ ���� : 2017-08-30
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

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
' ���ڸ����� ÷�������� ����մϴ�.
' - ÷������ ����� ���ڸ����� [�ӽ�����] ������ ��쿡�� �����մϴ�.
' - ÷�������� �ִ� 5������ ����� �� �ֽ��ϴ�.
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
' ���ڸ����� �ٸ� ���ڸ��� 1���� ÷���մϴ�.
'=========================================================================

Private Sub btnAttachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubItemCode = 121
    
    '÷���� ���ڸ��� ������ȣ
    SubMgtKey = "20151223-01"
      
    Set Response = statementService.AttachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ڸ����� [�������] ó���մϴ�.
'=========================================================================

Private Sub btnCancel_Click()
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
' 1���� ���ڸ����� [�������] ó���մϴ�.
'=========================================================================

Private Sub btnCancelISsue_2_Click()
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
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
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
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
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
' ���ڸ��� ������ȣ �ߺ����θ� Ȯ���մϴ�.
' - ������ȣ�� 1~24�ڸ��� ����, ���� '-', '_' �������� ������ �� �ֽ��ϴ�.
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
' 1���� ���ڸ����� �����մϴ�.
' - ���ڸ����� �����ϸ� ���� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������]
'=========================================================================

Private Sub btnDelete_2_Click()
    Dim Response As PBResponse
    
    Set Response = statementService.Delete(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' 1���� ���ڸ����� �����մϴ�.
' - ���ڸ����� �����ϸ� ���� ����������ȣ(mgtKey)�� ������ �� �ֽ��ϴ�.
' - ���������� ���� ���� : [�ӽ�����], [�������]
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
' ���ڸ����� ÷�ε� ������ �����մϴ�.
' - ������ �ĺ��ϴ� ���Ͼ��̵�� ÷������ ���(GetFileList API) �� �����׸�
'   �� ���Ͼ��̵�(AttachedFile) ���� ���� Ȯ���� �� �ֽ��ϴ�.
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
' ���ڸ����� ÷�ε� �ٸ� ���ڸ����� ÷�������մϴ�.
'=========================================================================

Private Sub btnDetachStatement_Click()
    Dim Response As PBResponse
    Dim SubItemCode As Integer
    Dim SubMgtKey As String
    
    '÷���� ���ڸ��� �����ڵ�, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    SubItemCode = 121
    
    '÷�������� ���ڸ��� ������ȣ
    SubMgtKey = "20151223-01"
      
    Set Response = statementService.DetachStatement(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, SubItemCode, SubMgtKey)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺��� ������� �ʰ� ���ڸ����� �ѽ������մϴ�.
' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnFAXSEnd_Click()
    Dim Statement As New PBStatement
    Dim ReceiptNum As String
    Dim newDetail As PBDocDetail
    Dim i
    
    '�ѽ� �߽Ź�ȣ
    Statement.sendNum = "07043042991"
    
    '�ѽ� ���Ź�ȣ
    Statement.receiveNum = "070000111"
       
    '[�ʼ�] ����� �ۼ�����, ��¥����(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ����������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '������ ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '������ ��ȣ
    Statement.senderCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.senderCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.senderAddr = "������ �ּ�"
    
    '������ ����
    Statement.senderBizClass = "������ ����"
    
    '������ ����
    Statement.senderBizType = "������ ����,����2"
    
    '������ ����ڼ���
    Statement.senderContactName = "������ ����ڸ�"
    
    '������ �̸���
    Statement.senderEmail = "test@test.com"
    
    '������ ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '������ �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ���޹޴��� ����
    '=========================================================================
    
    '���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '���޹޴��� ��ȣ
    Statement.receiverCorpName = "���޹޴��� ��ȣ"
    
    '���޹޴��� ��ǥ�� ����
    Statement.receiverCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Statement.receiverAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Statement.receiverBizClass = "���޹޴��� ���� "
    
    '���޹޴��� ����
    Statement.receiverBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Statement.receiverContactName = "���޹޴��� ����ڸ�"
    
    '���޹޴��� �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
        
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
    
    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
        newDetail.itemName = "ǰ��" + CStr(i)
        newDetail.spec = "�԰�"
        newDetail.unit = "����"
        newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "���"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    ReceiptNum = statementService.FAXSend(txtCorpNum.Text, Statement)
    
    If ReceiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + ReceiptNum
End Sub



'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
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
' ����ȸ���� ���ڸ��� API ���� ���������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = statementService.GetChargeInfo(txtCorpNum.Text, selectedItemCode)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (����ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub



'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
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
' ���ڸ��� 1���� �������� ��ȸ�մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] > 4.1.
'   ���ڸ��� ����" �� �����Ͻñ� �ٶ��ϴ�.
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

    '''  �󼼳��� ���� '''
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
' ���ڸ��� �μ�(���޹޴���) URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetEPrintURL_Click()
    Dim url As String
    
    url = statementService.GetEPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���ڸ����� ÷�ε� ������ ����� Ȯ���մϴ�.
' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ���ϻ���(DeleteFile API)
'   ȣ��� �̿��� �� �ֽ��ϴ�.
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
    
    tmp = "serialNum | attachedfile | displayName |  RegDT" + vbCrLf
    
    For Each file In resultList
        tmp = tmp + CStr(file.serialNum) + " | " + file.AttachedFile + " | " + file.DisplayName + " | " + file.regDT + vbCrLf
        txtFileID.Text = file.AttachedFile
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 1���� ���ڸ��� ����/��� ������ Ȯ���մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] > 3.3.1.
'   GetInfo (���� Ȯ��)"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetInfo_Click()
    Dim docInfo As PBDocInfo
    Dim tmp As String
    
    Set docInfo = statementService.GetInfo(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
     
    If docInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
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
' �ټ����� ���ڸ��� ����/��� ������ Ȯ���մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] > 3.3.2.
'   GetInfos (���� �뷮 Ȯ��)"�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetInfos_Click()
    Dim resultList As Collection
    Dim KeyList As New Collection
    Dim tmp As String
    Dim info As PBDocInfo
    
    '���ڸ��� ������ȣ �迭, �ִ� 1000��
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    Set resultList = statementService.GetInfos(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
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
' ���ڸ��� ���� �����̷��� Ȯ���մϴ�.
' - ���� �����̷� Ȯ��(GetLogs API) �����׸� ���� �ڼ��� ������
'   "[���ڸ��� API �����Ŵ���] > 3.3.4 GetLogs (���� �����̷� Ȯ��)"
'   �� �����Ͻñ� �ٶ��ϴ�.
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
    
    tmp = "DocLogType | Log | ProcType |  ProcMemo | RegDT | IP" + vbCrLf
    
    For Each log In resultList
        tmp = tmp + CStr(log.docLogType) + " | " + log.log + " | " + log.procType + " | " + log.procMemo + " | " + log.regDT + " | " + log.ip + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���޹޴��� ���ϸ�ũ URL�� ��ȯ�մϴ�.
' - ���ϸ�ũ URL�� ��ȿ�ð��� �������� �ʽ��ϴ�.
'=========================================================================

Private Sub btnGetMailURL_Click()
    Dim url As String
    
    url = statementService.GetMailURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ټ����� ���ڸ��� �μ��˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetMassPrintURL_Click()
    Dim url As String
    Dim KeyList As New Collection
    
    '�μ��� ���ڸ��� ������ȣ �迭, �ִ� 100��
    KeyList.Add "20161011-01"
    KeyList.Add "20161011-02"
    KeyList.Add "20161011-03"
    KeyList.Add "20161011-04"
    
    url = statementService.GetMassPrintURL(txtCorpNum.Text, selectedItemCode, KeyList)
     
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = statementService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = statementService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = statementService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ΰ� �� ÷�ι��� ��� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetSealURL_Click()
    Dim url As String
    
    url = statementService.GetSealURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1���� ���ڸ��� ���� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetPopUpURL_Click()
    Dim url As String
    
    url = statementService.GetPopUpURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� ���ڸ��� �μ��˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetPrintURL_Click()
    Dim url As String
  
    url = statementService.GetPrintURL(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub



'=========================================================================
' �˺� > ���ڸ��� > ���⹮���� �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_SBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "SBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���ڸ��� > �ӽ�(����)������ �˾� URL�� ��ȯ�մϴ�.
' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
'=========================================================================

Private Sub btnGetURL_TBOX_Click()
    Dim url As String
    
    url = statementService.GetURL(txtCorpNum.Text, txtUserID.Text, "TBOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 1���� [�ӽ�����] ������ ���ڸ����� ����ó���մϴ�.
' - ����� ����Ʈ�� �����˴ϴ�.
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
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1231212312"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = statementService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
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
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ڸ��� ���� �������� �׸� ���� ���ۿ��θ� ������� ��ȯ�մϴ�
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
            tmp = tmp + "���޹޴��ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_ACCEPT" Then
            tmp = tmp + "�����ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_DENY" Then
            tmp = tmp + "�����ڿ��� ���ڸ����� �ź� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڸ����� ��� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
        
        If resultList(i).emailType = "SMT_CANCEL_ISSUE" Then
            tmp = tmp + "���޹޴��ڿ��� ���ڸ����� ������� �Ǿ����� �˷��ִ� ���� : " + resultList(i).emailType + " | "
            tmp = tmp + CStr(resultList(i).sendYN) + vbCrLf
        End If
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ���ڸ��� ���� �������� �׸� ���� ���ۿ��θ� �����մϴ�.
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
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = statementService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.id = "testkorea_20161011"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = statementService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub btnRegister_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[�ʼ�] ����� �ۼ�����, ��¥����(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ����������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '������ ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '������ ��ȣ
    Statement.senderCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.senderCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.senderAddr = "������ �ּ�"
    
    '������ ����
    Statement.senderBizClass = "������ ����"
    
    '������ ����
    Statement.senderBizType = "������ ����,����2"
    
    '������ ����ڼ���
    Statement.senderContactName = "������ ����ڸ�"
    
    '������ �̸���
    Statement.senderEmail = "test@test.com"
    
    '������ ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '������ �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ���޹޴��� ����
    '=========================================================================
    
    '���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '���޹޴��� ��ȣ
    Statement.receiverCorpName = "���޹޴��� ��ȣ"
    
    '���޹޴��� ��ǥ�� ����
    Statement.receiverCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Statement.receiverAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Statement.receiverBizClass = "���޹޴��� ���� "
    
    '���޹޴��� ����
    Statement.receiverBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Statement.receiverContactName = "���޹޴��� ����ڸ�"
    
    '���޹޴��� �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
        
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

    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
        newDetail.itemName = "ǰ��" + CStr(i)
        newDetail.spec = "�԰�"
        newDetail.unit = "����"
        newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "���"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.Register(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' 1���� ���ڸ����� ��ù��� ó���մϴ�.
'=========================================================================

Private Sub btnRegistIssue_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    Statement.memo = "��ù��� �޸�"
    
    '[�ʼ�] ����� �ۼ�����, ��¥����(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ����������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '������ ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '������ ��ȣ
    Statement.senderCorpName = "������ ��ȣ"
    
    '������ ��ǥ�� ����
    Statement.senderCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.senderAddr = "������ �ּ�"
    
    '������ ����
    Statement.senderBizClass = "������ ����"
    
    '������ ����
    Statement.senderBizType = "������ ����,����2"
    
    '������ ����ڼ���
    Statement.senderContactName = "������ ����ڸ�"
    
    '������ �̸���
    Statement.senderEmail = "test@test.com"
    
    '������ ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '������ �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ���޹޴��� ����
    '=========================================================================
    
    '���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '���޹޴��� ��ȣ
    Statement.receiverCorpName = "���޹޴��� ��ȣ"
    
    '���޹޴��� ��ǥ�� ����
    Statement.receiverCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Statement.receiverAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Statement.receiverBizClass = "���޹޴��� ���� "
    
    '���޹޴��� ����
    Statement.receiverBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Statement.receiverContactName = "���޹޴��� ����ڸ�"
    
    '���޹޴��� �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
        
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
    
    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
        newDetail.itemName = "ǰ��" + CStr(i)
        newDetail.spec = "�԰�"
        newDetail.unit = "����"
        newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "���"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.RegistIssue(txtCorpNum.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˻������� ����Ͽ� ���ڸ��� ����� ��ȸ�մϴ�.
' - �����׸� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
'   3.3.3. Search (��� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
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
    
    '[�ʼ�] ��������, R-����Ͻ� W-�ۼ����� I-�����Ͻ� �� ��1
    DType = "W"
    
    '[�ʼ�] ��������, yyyyMMdd
    SDate = "20160901"
    
    '[�ʼ�] ��������, yyyyMMdd
    EDate = "20161031"
    
    '���ۻ��°� �迭, �̱���� ��ü������ȸ, �������°� 3�ڸ����� �ۼ�
    '2,3��° ���ϵ�ī�� ����
    state.Add "100"
    state.Add "2**"
    state.Add "3**"
    
    '���������ڵ� �迭, 121-�ŷ�����, 122-û����, 123-������, 124-���ּ�, 125-�Ա�ǥ,126-������
    itemCode.Add "121"
    itemCode.Add "122"
    itemCode.Add "123"
    itemCode.Add "124"
    itemCode.Add "125"
    itemCode.Add "126"
    
    '������ ��ȣ
    Page = 1
    
    '������ ��ϰ���, �ִ� 1000��
    PerPage = 15
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '�ŷ�ó ����, �ŷ�ó ��ȣ �Ǵ� �ŷ�ó ����ڵ�Ϲ�ȣ ����, �̱���� ��ü��ȸ
    QString = ""
        
    Set docSearchList = statementService.Search(txtCorpNum.Text, DType, SDate, EDate, state, itemCode, Page, PerPage, Order, QString)
     
    If docSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
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
' ���� �ȳ������� �������մϴ�.
'=========================================================================

Private Sub btnSendEmail_Click()
    Dim Response As PBResponse
    Dim receiverEmail As String
    
    '���Ÿ����ּ�
    receiverEmail = "test@test.com"
  
    Set Response = statementService.SendEmail(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, receiverEmail)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ����� �ѽ������մϴ�.
' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnSendFAX_Click()
    Dim Response As PBResponse
    Dim senderNum As String
    Dim receiverNum As String
    
    '�߽Ź�ȣ
    senderNum = "070-4304-2991"
    
    '�����ѽ���ȣ
    receiverNum = "070-111-222"
    
    Set Response = statementService.SendFax(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˸����ڸ� �����մϴ�. (�ܹ�/SMS- �ѱ� �ִ� 45��)
' - �˸����� ���۽� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [���۳���] �ǿ���
'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
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
    
    '���ڸ޽��� ����, 90Byte �ʰ��� ������ �����Ǿ� ���۵�
    Contents = "���ڸ����� �����Ͽ����ϴ�. ������ Ȯ���Ͽ� �ֽñ�ٶ��ϴ�"
    
    Set Response = statementService.SendSMS(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, senderNum, receiverNum, Contents)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ���ڸ��� ����ܰ��� Ȯ���մϴ�.
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
' 1���� ���ڸ����� �����մϴ�.
' - [�ӽ�����] ������ ���ڸ����� ������ �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnUpdate_Click()
    Dim Statement As New PBStatement
    Dim newDetail As PBDocDetail
    Dim Response As PBResponse
    Dim i
    
    '[�ʼ�] ����� �ۼ�����, ��¥����(yyyyMMdd)
    Statement.writeDate = "20170223"
    
    '[�ʼ�] {����, û��} �� ����
    Statement.purposeType = "����"
    
    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    Statement.taxType = "����"
    
    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    Statement.formCode = txtFormCode.Text
    
    '[�ʼ�] ���ڸ��� �����ڵ�
    Statement.itemCode = selectedItemCode
    
    '[�ʼ�] ����������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����
    Statement.mgtKey = txtMgtKey.Text
    
    
    '=========================================================================
    '                               ������ ����
    '=========================================================================
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.senderCorpNum = txtCorpNum.Text
    
    '������ ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    Statement.senderTaxRegID = ""
    
    '������ ��ȣ
    Statement.senderCorpName = "������ ��ȣ_����"
    
    '������ ��ǥ�� ����
    Statement.senderCEOName = "������ ��ǥ�� ����"
    
    '������ �ּ�
    Statement.senderAddr = "������ �ּ�_����"
    
    '������ ����
    Statement.senderBizClass = "������ ����_����"
    
    '������ ����
    Statement.senderBizType = "������ ����,����2"
    
    '������ ����ڼ���
    Statement.senderContactName = "������ ����ڸ�"
    
    '������ �̸���
    Statement.senderEmail = "test@test.com"
    
    '������ ����ó
    Statement.senderTEL = "070-7070-0707"
    
    '������ �޴���ȭ ��ȣ
    Statement.senderHP = "010-000-2222"
    
    
    '=========================================================================
    '                        ���޹޴��� ����
    '=========================================================================
    
    '���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    Statement.receiverCorpNum = "8888888888"
    
    '���޹޴��� ��ȣ
    Statement.receiverCorpName = "���޹޴��� ��ȣ"
    
    '���޹޴��� ��ǥ�� ����
    Statement.receiverCEOName = "���޹޴��� ��ǥ�� ����"
    
    '���޹޴��� �ּ�
    Statement.receiverAddr = "���޹޴��� �ּ�"
    
    '���޹޴��� ����
    Statement.receiverBizClass = "���޹޴��� ���� "
    
    '���޹޴��� ����
    Statement.receiverBizType = "���޹޴��� ����"
    
    '���޹޴��� ����ڸ�
    Statement.receiverContactName = "���޹޴��� ����ڸ�"
    
    '���޹޴��� �����ּ�
    Statement.receiverEmail = "test@receiver.com"
    
    
    '=========================================================================
    '                     ���ڸ��� �������
    '=========================================================================
    
    '[�ʼ�] ���ް��� �հ�
    Statement.supplyCostTotal = "100000"
    
    '[�ʼ�] ���� �հ�
    Statement.taxTotal = "10000"
    
    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    Statement.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
        
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
    
    For i = 1 To 20
    
        Set newDetail = New PBDocDetail
        
        newDetail.serialNum = i             '�Ϸù�ȣ 1���� ���� ����
        newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
        newDetail.itemName = "ǰ��" + CStr(i)
        newDetail.spec = "�԰�"
        newDetail.unit = "����"
        newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
        newDetail.supplyCost = "100000"
        newDetail.tax = "10000"
        newDetail.remark = "���"
        newDetail.spare1 = "spare1"
        newDetail.spare2 = "spare2"
        newDetail.spare3 = "spare3"
        newDetail.spare4 = "spare4"
        newDetail.spare5 = "spare5"
        
        Statement.detailList.Add newDetail
        
    Next
    
    '=========================================================================
    '���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
    '=========================================================================
    
    Set Statement.propertyBag = New Dictionary
    
    Statement.propertyBag.Add "CBalance", "100000"
    Statement.propertyBag.Add "Deposit", "10000"
    Statement.propertyBag.Add "Balance", "100000"
    
    Set Response = statementService.Update(txtCorpNum.Text, selectedItemCode, txtMgtKey.Text, Statement)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-4304-2991"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
                
    Set Response = statementService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = statementService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(statementService.LastErrCode) + vbCrLf + "����޽��� : " + statementService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub



Private Sub Form_Load()
    '���ڸ��� ��ü �ʱ�ȭ
    statementService.Initialize LinkID, SecretKey
    
    '����ȯ�漳����, True-���߿� False-�����
    statementService.IsTest = True
End Sub

