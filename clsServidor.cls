VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mID As String
Private mHostname As String
Private mFQDN As String
Private mDescricao As String
Private mAmbiente As String
Private mSistemaOperacional As String
Private mVersaoSO As String
Private mArquitetura As String
Private mFabricante As String
Private mModelo As String
Private mSerial As String
Private mPatrimonio As String
Private mVirtual As Boolean
Private mHypervisor As String
Private mHostFisico As String
Private mCluster As String
Private mDatastore As String
Private mCPU As String
Private mMemoriaGB As Double
Private mDiscoGB As Double
Private mDataCadastro As Date
Private mDataAtualizacao As Date
Private mObservacoes As String
Private mIsNew As Boolean
Private mIsLoaded As Boolean
Private mIsDirty As Boolean

Private Sub Class_Initialize()
    Clear
End Sub

Public Sub Clear()
 mID = "": mHostname = "": mFQDN = "": mDescricao = "": mAmbiente = "": mSistemaOperacional = "": mVersaoSO = "": mArquitetura = "": mFabricante = "": mModelo = "": mSerial = "": mPatrimonio = "": mVirtual = False: mHypervisor = "": mHostFisico = "": mCluster = "": mDatastore = "": mCPU = "": mMemoriaGB = 0: mDiscoGB = 0: mDataCadastro = 0: mDataAtualizacao = 0: mObservacoes = ""
 mIsNew = True: mIsLoaded = False: mIsDirty = False
End Sub

Public Property Get ID() As String: ID = mID: End Property
Public Property Let ID(ByVal Value As String): mID = Trim$(Value): End Property
Public Property Get Hostname() As String: Hostname = mHostname: End Property
Public Property Let Hostname(ByVal Value As String): mHostname = UCase$(Trim$(Value)): End Property
Public Property Get FQDN() As String: FQDN = mFQDN: End Property
Public Property Let FQDN(ByVal Value As String): mFQDN = LCase$(Trim$(Value)): End Property
Public Property Get Descricao() As String: Descricao = mDescricao: End Property
Public Property Let Descricao(ByVal Value As String): mDescricao = Trim$(Value): End Property
Public Property Get Ambiente() As String: Ambiente = mAmbiente: End Property
Public Property Let Ambiente(ByVal Value As String): mAmbiente = Trim$(Value): End Property
Public Property Get SistemaOperacional() As String: SistemaOperacional = mSistemaOperacional: End Property
Public Property Let SistemaOperacional(ByVal Value As String): mSistemaOperacional = Trim$(Value): End Property
Public Property Get VersaoSO() As String: VersaoSO = mVersaoSO: End Property
Public Property Let VersaoSO(ByVal Value As String): mVersaoSO = Trim$(Value): End Property
Public Property Get Arquitetura() As String: Arquitetura = mArquitetura: End Property
Public Property Let Arquitetura(ByVal Value As String): mArquitetura = Trim$(Value): End Property
Public Property Get Fabricante() As String: Fabricante = mFabricante: End Property
Public Property Let Fabricante(ByVal Value As String): mFabricante = Trim$(Value): End Property
Public Property Get Modelo() As String: Modelo = mModelo: End Property
Public Property Let Modelo(ByVal Value As String): mModelo = Trim$(Value): End Property
Public Property Get Serial() As String: Serial = mSerial: End Property
Public Property Let Serial(ByVal Value As String): mSerial = Trim$(Value): End Property
Public Property Get Patrimonio() As String: Patrimonio = mPatrimonio: End Property
Public Property Let Patrimonio(ByVal Value As String): mPatrimonio = Trim$(Value): End Property
Public Property Get Virtual() As Boolean: Virtual = mVirtual: End Property
Public Property Let Virtual(ByVal Value As Boolean): mVirtual = Value: End Property
Public Property Get Hypervisor() As String: Hypervisor = mHypervisor: End Property
Public Property Let Hypervisor(ByVal Value As String): mHypervisor = Trim$(Value): End Property
Public Property Get HostFisico() As String: HostFisico = mHostFisico: End Property
Public Property Let HostFisico(ByVal Value As String): mHostFisico = Trim$(Value): End Property
Public Property Get Cluster() As String: Cluster = mCluster: End Property
Public Property Let Cluster(ByVal Value As String): mCluster = Trim$(Value): End Property
Public Property Get Datastore() As String: Datastore = mDatastore: End Property
Public Property Let Datastore(ByVal Value As String): mDatastore = Trim$(Value): End Property
Public Property Get CPU() As String: CPU = mCPU: End Property
Public Property Let CPU(ByVal Value As String): mCPU = Trim$(Value): End Property
Public Property Get MemoriaGB() As Double: MemoriaGB = mMemoriaGB: End Property
Public Property Let MemoriaGB(ByVal Value As Double): mMemoriaGB = Value: End Property
Public Property Get DiscoGB() As Double: DiscoGB = mDiscoGB: End Property
Public Property Let DiscoGB(ByVal Value As Double): mDiscoGB = Value: End Property
Public Property Get DataCadastro() As Date: DataCadastro = mDataCadastro: End Property
Public Property Let DataCadastro(ByVal Value As Date): mDataCadastro = Value: End Property
Public Property Get DataAtualizacao() As Date: DataAtualizacao = mDataAtualizacao: End Property
Public Property Let DataAtualizacao(ByVal Value As Date): mDataAtualizacao = Value: End Property
Public Property Get Observacoes() As String: Observacoes = mObservacoes: End Property
Public Property Let Observacoes(ByVal Value As String): mObservacoes = Trim$(Value): End Property
Public Property Get IsNew() As Boolean: IsNew = mIsNew: End Property
Public Property Get IsLoaded() As Boolean: IsLoaded = mIsLoaded: End Property
Public Property Get IsDirty() As Boolean: IsDirty = mIsDirty: End Property

Public Function Validate() As Boolean
 Validate = (Len(mHostname) > 0 And Len(mSistemaOperacional) > 0)
End Function

Public Function IsVirtual() As Boolean
 IsVirtual = mVirtual
End Function

Public Sub Touch()
 If mDataCadastro = 0 Then mDataCadastro = Now
 mDataAtualizacao = Now
 MarkDirty
End Sub

Public Sub MarkDirty()
 mIsDirty = True
End Sub

Public Sub AcceptChanges()
 mIsNew = False
 mIsLoaded = True
 mIsDirty = False
End Sub
