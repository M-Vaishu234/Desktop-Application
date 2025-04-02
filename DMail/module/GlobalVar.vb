Imports System.Configuration

Public Class GlobalVar

    Public Shared CompanyCode As Long = 47
    Public Shared CompanyName As String = ""
    Public Shared ConnectToApi As Long = 0
    Public Shared UserCode As Long
    Public Shared UserName As String = ""
    Public Shared DatabaseName As String
    Public Shared MainDBName As String
    Public Shared PortId As String = ""
    Public Shared POLName As String = ""
    Public Shared IECCode As String = ""

    Public Shared ICE_UserId As String

    Public Shared LicenseType As String
    Public Shared IsRead As Boolean = False
    Public Shared token As String
    Public Shared lf = System.Convert.ToChar(System.Convert.ToUInt32("0A", 16))
    Public Shared mailId As String = ""
    Public Shared FwdIDname As String = ""
    Public Shared FwdIDaddress As String = ""

    Public Shared Branch_Key As Long = 0
    Public Shared Branch_No As String = ""
    Public Shared AD_Key As Long = 0
    Public Shared AD_Code As String = ""
    Public Shared Port_Code As String = ""
    Public Shared Port_Name As String = ""

    Public Shared ThumbPrint As String = ""
    Public Shared Signature As String = ""


    Public Shared chkFile As Integer = 0
    Public Shared fileRead As Boolean = False
    Public Shared chkDSR As Integer = 0
    Public Shared DsrRead As Boolean = False
    Public Shared chkAck As Integer = 0
    Public Shared AckRead As Boolean = False
    Public Shared chkleo As Integer = 0
    Public Shared LeoRead As Boolean = False
    Public Shared chkGate As Integer = 0
    Public Shared GRead As Boolean = False

    Public Class GetCurrentTime

        Public Property CurrentDay() As Integer
            Get
                Return m_CurrentDay
            End Get
            Set(ByVal value As Integer)
                m_CurrentDay = value
            End Set
        End Property
        Private m_CurrentDay As Integer

        Public Property CurrentMonth() As Integer
            Get
                Return m_CurrentMonth
            End Get
            Set(ByVal value As Integer)
                m_CurrentMonth = value
            End Set
        End Property
        Private m_CurrentMonth As Integer

        Public Property CurrentYear() As Integer
            Get
                Return m_CurrentYear
            End Get
            Set(ByVal value As Integer)
                m_CurrentYear = value
            End Set
        End Property
        Private m_CurrentYear As Integer

        Public Property CurrentHour() As Integer
            Get
                Return m_CurrentHour
            End Get
            Set(ByVal value As Integer)
                m_CurrentHour = value
            End Set
        End Property
        Private m_CurrentHour As Integer

        Public Property CurrentMin() As Integer
            Get
                Return m_CurrentMin
            End Get
            Set(ByVal value As Integer)
                m_CurrentMin = value
            End Set
        End Property
        Private m_CurrentMin As Integer

        Public Property CurrentSec() As Integer
            Get
                Return m_CurrentSec
            End Get
            Set(ByVal value As Integer)
                m_CurrentSec = value
            End Set
        End Property
        Private m_CurrentSec As Integer

        Public Property CurrentTime() As Date
            Get
                Return m_CurrentTime
            End Get
            Set(ByVal value As Date)
                m_CurrentTime = value
            End Set
        End Property
        Private m_CurrentTime As Date

    End Class

    Public Class clsAckHead
        Public Property strFilter As String
        Public Property AckKey As Long
        Public Property DtReceived As String
        Public Property DtFetched As String
        Public Property DtRead As String
        Public Property FName As String
        Public Property msgtype As String
        Public Property repEvent As String
        Public Property SendId As String
        Public Property JobNo As Long
        Public Property Status As Integer
        Public Property db_name As String
        Public Property GroupType As String
    End Class


    Public Class cls_Ack
        Public Property headerField As HeaderField
        Public Property master As Master
        Public Property errorDetails As ErrorDetails
    End Class

    Public Class Master
        Public Property decRef As DecRef
        Public Property authPrsn As AuthPrsn
        Public Property vesselDtls As VesselDtls
        Public Property voyageDtls As VoyageDtls
        Public Property mastrCnsgmtDec As MastrCnsgmtDec()
    End Class

    Public Class HeaderField
        Public Property senderID As String
        Public Property receiverID As String
        Public Property messageID As String
        Public Property sequenceOrControlNumber As Integer
        Public Property reportingEvent As String
    End Class

    Public Class DecRef
        Public Property msgTyp As String
        Public Property prtofRptng As String
        Public Property jobNo As Integer
        Public Property jobDt As String
        Public Property rptngEvent As String
        Public Property errorCode As ErrorCode()
    End Class

    Public Class AuthPrsn
        Public Property errorCode As ErrorCode()
    End Class


    Public Class VesselDtls
        Public Property errorCode As ErrorCode()
    End Class

    Public Class VoyageDtls
        Public Property errorCode As ErrorCode()
    End Class

    Public Class ErrorCode
        Public Property pathName As String
        Public Property errorCode As String
        Public Property errorMessage As String
    End Class

    Public Class MCRef
        Public Property lineNo As Integer
        Public Property errorCode As ErrorCode()
    End Class

    Public Class PrevRef
        Public Property errorCode As ErrorCode()
    End Class

    Public Class LocCstm
        Public Property errorCode As ErrorCode()
    End Class

    Public Class Trnshpr
        Public Property errorCode As ErrorCode()
    End Class

    Public Class TrnsprtDoc
        Public Property errorCode As ErrorCode()
    End Class

    Public Class TrnsprtDocMsr
        Public Property errorCode As ErrorCode()
    End Class

    Public Class ItemDtl
        Public Property crgoItemSeqNmbr As Integer
        Public Property errorCode As ErrorCode()
    End Class

    Public Class TrnsprtEqmt
        Public Property eqmtSeqNo As Integer
        Public Property errorCode As ErrorCode()
    End Class

    Public Class Itnry
        Public Property prtOfCallSeqNmbr As Integer
        Public Property errorCode As ErrorCode()
    End Class

    Public Class McResponse
        Public Property cinType As Object
        Public Property mcinPcin As Object
    End Class

    Public Class MastrCnsgmtDec
        Public Property MCRef As MCRef
        Public Property prevRef As PrevRef
        Public Property locCstm As LocCstm
        Public Property trnshpr As Trnshpr
        Public Property trnsprtDoc As TrnsprtDoc
        Public Property trnsprtDocMsr As TrnsprtDocMsr
        Public Property itemDtls As ItemDtl()
        Public Property trnsprtEqmt As TrnsprtEqmt()
        Public Property itnry As Itnry()
        Public Property mcResponse As McResponse
    End Class

    Public Class Schema
        Public Property loadingURI As String
        Public Property pointer As String
    End Class

    Public Class Instance
        Public Property pointer As String
    End Class

    Public Class ErrorMessage
        Public Property level As String
        Public Property schema As Schema
        Public Property instance As Instance
        Public Property domain As String
        Public Property keyword As String
        Public Property message As String
        Public Property value As String
        Public Property found As Integer
        Public Property maxLength As Integer
    End Class

    Public Class ErrorDetails
        Public Property status As String
        Public Property errorCode As String
        Public Property errorMessage As ErrorMessage()
    End Class


End Class

Public Class Globalitemvar
    Public Shared codetype, gcode, gadd, gcity, gcountry, gpin, mcountry, sstate, tcountry, astatus, hawb, totpackage As String
End Class

Public Class Globalcessvar
    Public Shared sno, cunit, tval, cval, tunit, cqty, cflag, cadv As String
End Class




