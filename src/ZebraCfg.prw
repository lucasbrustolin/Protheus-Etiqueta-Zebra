#Include "Protheus.ch"
#INCLUDE "FWMVCDEF.CH"
#include "rwmake.ch"
#INCLUDE 'FWEditPanel.CH'
#INCLUDE "FWBrowse.ch"

//-- CONTROLE VLD ARQUIVOS
#DEFINE ARQUIVOS_PASTAS_PREVIEW_BROWSE  1
#DEFINE ARQUIVOS_PASTAS_PREVIEW_ZPLVIEW 2

//--  MODO IMPRESSAO
#DEFINE IMP_SPOOL_PREVIEW  1
#DEFINE IMP_PREVIEW 2

#DEFINE Enter Chr(13) + Chr(10)

#DEFINE OPER_COPY 9

Static __ClearGrid  := .F.
Static __lFlagUrl   := .F.
//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@type				: Funcao de Usuario
@Sample				: U_ZEBRACFG()
@description	    : Monta Browse CRUD MVC da tabela de Configuracao de Etiquetas (Zebra) 						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/010/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
User Function ZEBRACFG()

Local oDlg      := Nil 
Local cImageAtu := ""
Local nAltura	:= 0
Local nLargura	:= 0

//-----------------------
Private oBrowse     := Nil
Private aRotina     := FwLoadMenuDef("ZEBRACFG")
Private __oPrint	:= Nil
Private __oImagem   := Nil 
Private __oWebPage  := Nil
Private __URLPage   := ""  
Private __aIMGZebra := {}
Private __lExecSx7  := .F. //--  Controle do gatilho - Carregamento do Grid PDK
Private __ZebraCfg  := ""

//-------------------------------------------------------------------------------+
// Valida a existencia de funcoes e tabelas para a utilização deste configurador |
//-------------------------------------------------------------------------------+
If !( FindFunction("U_MSCBPrinter") )
    FwAlertWarning("O configurador exige a inclusao da rotina U_MSCBPrinter. " + Enter + ;
                    "Conte o suporte Protheus!","ZebraCfg")
    Return( .F. )
ElseIf ( !AliasInDic("PDI") .Or. !AliasInDic("PDJ") .Or. !AliasInDic("PDK") )
    FwAlertWarning("Tabelas utilizadas neste configurador nao foram incluidas." + Enter + ;
                    "Tabelas: PDI\PDJ\PDK " + Enter + ;
                    "Contate o suporte Protheus!","ZebraCfg")
    Return( .F. )
EndIf 

//--------------------------------------------+
// Copia os arquivos do servidor para client  |
//-------------------------------------------+
CheckDirFile(ARQUIVOS_PASTAS_PREVIEW_BROWSE,@__aIMGZebra) 
    
If !( CheckDirFile(ARQUIVOS_PASTAS_PREVIEW_ZPLVIEW) )
    FwAlertWarning("A ausencia dos arquivos e\ou pastas impactam diretamente no Preview da etiqieta!","ZebraCfg")
EndIf 


//-- Faz a inclusao dos 3 modelos de exemplo
FWMsgRun(, {|oSay| GravaModel(oSay) },"Modelo de Etiqueta","Validando modelo..." )


// Define as Colunas
aColunas := {}

oSize := FwDefSize():New(.F.)

oSize:AddObject( "CABECALHO",(oSize:aWindSize[4]),(oSize:aWindSize[3]) , .F., .F. ) // Não dimensionavel
oSize:aMargins 	:= { 3, 3, 3, 3  } 	// Espaco ao lado dos objetos 0, entre eles 3		
oSize:lProp 		:= .F. 			// Proporcional             
oSize:Process() 	   				// Dispara os calculos  


DEFINE MSDIALOG oDlg TITLE OemToAnsi( "Configurador Etiquetas (Zebra)" ) ;
From oSize:aWindSize[1],oSize:aWindSize[2] TO (oSize:aWindSize[3]),(oSize:aWindSize[4] ) OF oMainWnd ;
PIXEL STYLE nOR( WS_VISIBLE, WS_POPUP )	


// Cria o conteiner onde serão colocados os paineis
oTela		:= FWFormContainer():New( oDlg )
cIdCab	  	:= oTela:CREATEVERTICALBOX( 60 )
cIdGrid  	:= oTela:CREATEVERTICALBOX( 40 )     


oTela:Activate( oDlg, .F. )

//Cria os paineis onde serao colocados os browses
oPanelUp  	:= oTela:GeTPanel( cIdCab )
oPanelDown	:= oTela:GeTPanel( cIdGrid )

Define Font oFont Name 'Courier New' Size 0, -12

oBrowse:= FWBrowse():New()
oBrowse:SetOwner(oPanelUp)
oBrowse:SetDataTable(.T.)
oBrowse:SetAlias("PDI")
oBrowse:SetDescription("Configurador Etiquetas (Zebra)")

oBrowse:SetColumns( GetColumns() )

oBrowse:DisableReport()
oBrowse:SetUseFilter() // Habilita a utilização do Filtro de registros
oBrowse:SetLocate() // Habilita a Localização de registros
oBrowse:SetSeek() // Habilita a Pesquisa de registros
oBrowse:SetDBFFilter()
// Define ação na troca de linha (Change)
oBrowse:SetChange({|oBrowse| ChangePict(oBrowse) })

oBrowse:Activate()


nAltura		:= (oPanelDown:nHeight/2) - 05			//Altura padrão dos paineis
nLargura	:= (oPanelDown:nWidth/2) - 10			//Largura padrão dos paineis
//310 x 260 

 oTScrollBox :=  TScrollBox():New(oPanelDown,05,05,nAltura,nLargura,.T.,.T.,.T.) 

__oImagem := TBitmap():New(05,05,260,184,,cImageAtu,.T.,oTScrollBox /*oPanelDown*/,;
{|| .T. },,.F.,.F.,,,.F.,,.T.,,.F.)
__oImagem:lAutoSize := .T.


// // relaciona os paineis aos componentes
oBar := FWButtonBar():New()
oBar:Init( oPanelUp , 25 , 25 , CONTROL_ALIGN_TOP , .T. )

oBar:AddBtnImage( "ADICIONAR_001.PNG" 		,'Incluir Etiqueta'   ,{|| MenuDef(3) } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "ALTERA.PNG"        		,'Alterar Etiqueta'   ,{|| MenuDef(4) } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "BMPVISUAL.PNG" 		    ,'Visualizar Etiqueta',{|| MenuDef(1) } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "EXCLUIR.PNG"       		,'Excluir Etiqueta'   ,{|| MenuDef(5) } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "SDUCOPYTO.PNG"       	,'Copiar Etiqueta'    ,{|| MenuDef(9) } ,, .T., CONTROL_ALIGN_LEFT )

oBar:AddBtnImage( "PRINT03.PNG"       		,'Imprimir Etiqueta'        ,{|| U_IMPZEBRA( IMP_SPOOL_PREVIEW ) } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "DBG06.PNG"       		,'Testar Porta Impressao'   ,{|| MSCBTestePort() } ,, .T., CONTROL_ALIGN_LEFT )
oBar:AddBtnImage( "PMSCOLOR.PNG"       		,'Legenda'                  ,{|| LegendBrw() } ,, .T., CONTROL_ALIGN_LEFT )


oBar:AddBtnImage( "FINAL.PNG"        		,'Sair'             ,{|| oDlg:End()} ,{|| .T.}, .T., CONTROL_ALIGN_LEFT )


ACTIVATE MSDIALOG oDlg CENTERED 

Return()

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: MenuDef()
@description	    : Adicona o menu do cadastro						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------

Static Function MenuDef(nOper)  

Local oModel        := Nil 
Local cTitulo       := ""
Local cPrograma     := ""
Local nOperation    := 0
Local nRet          := 0
Local aRotina       := {}

If Empty(nOper)
    ADD OPTION aRotina TITLE "Visualizar"   ACTION "VIEWDEF.ZEBRACFG" OPERATION 2 ACCESS 0  // "Visualizar"
    ADD OPTION aRotina TITLE "Incluir"      ACTION "VIEWDEF.ZEBRACFG" OPERATION 3 ACCESS 0  // "Incluir"
    ADD OPTION aRotina TITLE "Alterar"      ACTION "VIEWDEF.ZEBRACFG" OPERATION 4 ACCESS 0  // "Alterar"
    ADD OPTION aRotina TITLE "Excluir"      ACTION "VIEWDEF.ZEBRACFG" OPERATION 5 ACCESS 0  // "Excluir"
    ADD OPTION aRotina TITLE "Copiar"       ACTION "VIEWDEF.ZEBRACFG" OPERATION 9 ACCESS 0  // "Copiar"
    ADD OPTION aRotina TITLE "Imp. Etiqueta"ACTION "U_IMPZEBRA(1)"    OPERATION 2 ACCESS 0  //"Imp. Etiqueta"
    ADD OPTION aRotina TITLE "Testar Porta" ACTION "MSCBTestePort"    OPERATION 2 ACCESS 0  //"TESTE"
Else 

    cTitulo     := "Configurador Etiquetas (Zebra)"
    cPrograma   := 'ZEBRACFG'
    nOperation  := nOper  

    Do Case    
        Case nOper = MODEL_OPERATION_VIEW
            cTitulo := "Visualizar"
        Case nOper = MODEL_OPERATION_INSERT
            cTitulo := "Inclusao"
        Case nOper = MODEL_OPERATION_UPDATE
            cTitulo := "Alteracao"
        Case nOper = MODEL_OPERATION_DELETE
            cTitulo := "Exclusao"
        Case nOper = OPER_COPY
            cTitulo := "Copia"
            nOperation := MODEL_OPERATION_INSERT
        OtherWise
    EndCase

    oModel := FWLoadModel( cPrograma )
    oModel:SetOperation( nOperation ) 
    
    If nOper == OPER_COPY
        oModel:Activate(.T.) // Ativa o modelo com os dados posicionados
        oModel:SetValue("PDIMASTER","PDI_PROPRI","U")
    Else 
        oModel:Activate()
    EndIf

    nRet := FWExecView( cTitulo , cPrograma, nOperation, /*oDlg*/, {|| .T. } ,/*bOk*/ , /*nPercReducao*/, /*aEnableButtons*/, /*bCancel*/ , /*cOperatId*/, /*cToolBar*/, oModel )
    oModel:DeActivate()
    
    If ValType(oBrowse) == "O"
        oBrowse:Refresh(.T.)
    EndIf


EndIf 

Return( aRotina )


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: ModelDef()
@description	    : Cria Modelo de dados 						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------

Static Function ModelDef()

Local oStruPDI		:= FWFormStruct( 1, "PDI" )			
Local oStruPDJ		:= FWFormStruct( 1, "PDJ" )	
Local oStruPDK		:= FWFormStruct( 1, "PDK" )	
Local bVldTudOk     := {|oModel| VldTudOk( oModel ) }
Local bLinPre       := {|oModel,nLinha,cAcao,cCampo,xValue,xOldValue| PreVldLin(oModel,nLinha,cAcao,cCampo,xValue,xOldValue) } 				
Local aRelacPDJ     := {}
Local aRelacPDK     := {}


// Define quais sao os campos obrigatórios.
oStruPDI:SetProperty("PDI_DESC" , MODEL_FIELD_OBRIGAT, .T. )
oStruPDI:SetProperty("PDI_PORTA", MODEL_FIELD_OBRIGAT, .T. )
oStruPDJ:SetProperty("PDJ_TIPO" , MODEL_FIELD_OBRIGAT, .T. )


//-- Gatilho que faz o load do grid inferior ( PDK ) 
oStruPDJ:AddTrigger( "PDJ_TIPO", "PDJ_TIPO", {|| .T. }, {|| X7InsPar() }  )

// Cria o objeto do Modelo de Dados
oModel := MPFormModel():New("MZEBRACFG",/*bPreValid*/,bVldTudOk,/*bCommit*/ )
oModel:AddFields("PDIMASTER",/*cOwner*/     , oStruPDI)
oModel:AddGrid( 'PDJDETAIL', 'PDIMASTER'	, oStruPDJ, bLinPre /*bLinePre*/ , /*bLinePost*/, /*bPreVal*/, /*bPosVal*/, /*BLoad*/ )
oModel:AddGrid( 'PDKDETAIL', 'PDJDETAIL'	, oStruPDK, /*bLinePre*/ , /*bLinePost*/, /*bPreVal*/, /*bPosVal*/, /*BLoad*/ )

//Relacionamento da tabela Etapa com Projeto
aAdd(aRelacPDJ,{ 'PDJ_FILIAL'	, 'xFilial( "PDJ" )'	})
aAdd(aRelacPDJ,{ 'PDJ_CODPDI'	, 'PDI_COD' 		    })

// Faz relaciomaneto entre os compomentes do model
oModel:SetRelation( 'PDJDETAIL', aRelacPDJ , PDJ->( IndexKey( 1 ) )  )

//Relacionamento da tabela Etapa com Projeto
aAdd(aRelacPDK,{ 'PDK_FILIAL'	, 'xFilial( "PDK" )'	})
aAdd(aRelacPDK,{ 'PDK_CODPDI'	, 'PDI_COD' 		    })
aAdd(aRelacPDK,{ 'PDK_ITEM'	    , 'PDJ_ITEM' 		    })

// Faz relaciomaneto entre os compomentes do model
oModel:SetRelation( 'PDKDETAIL', aRelacPDK , PDK->( IndexKey( 1 ) )  )

//Deixa o prrenchimento das tabelas opcional
oModel:GetModel( 'PDKDETAIL' ):SetOptional( .T. )

//-- Permite apagar todas as linhas do grid
oModel:GetModel( 'PDKDETAIL' ):SetDelAllLine(.T.)

//-- Regra de manutencao: Bloqueia inclusao e delecao de linhas 
oModel:GetModel( 'PDKDETAIL' ):SetNoInsertLine( .T. )
oModel:GetModel( 'PDKDETAIL' ):SetNoDeleteLine( .T. )

//-- Descricao de cada submodelo (componentes)
oModel:SetDescription("Modelo Configurador Etiquetas (Zebra)") 
oModel:GetModel( 'PDJDETAIL' ):SetDescription("Layout Etiqueta")
oModel:GetModel( 'PDKDETAIL' ):SetDescription("Definicoes Layout") 

oModel:SetOnDemand( .T. )

oModel:SetVldActivate({|oModel| ActiveModel(oModel)  })

__ClearGrid := .F.

Return( oModel )

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: ViewDef()
@description	    : Cria Modelo de Visualização						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------

Static Function ViewDef() 

Local oView		:= FWFormView():New()	
Local oModel	:= FwLoadModel("ZEBRACFG") 				// Cria um objeto de Modelo de dados 
Local oStruPDI	:= FWFormStruct( 2, "PDI" )			
Local oStruPDJ	:= FWFormStruct( 2, "PDJ" )	
Local oStruPDK	:= FWFormStruct( 2, "PDK" )					

oView:SetModel(oModel)							

//-- PDI
// oStruPDI:RemoveField("PDI_MEMO")
//-- PDJ
oStruPDJ:RemoveField("PDJ_CODPDI")
//-- PDK
oStruPDK:RemoveField("PDK_CODPDI")
oStruPDK:RemoveField("PDK_ITEM")
oStruPDK:RemoveField("PDK_TIPO")
oStruPDK:RemoveField("PDK_VALID")

//-- Executa ação no setfocus no campo
oView:SetFieldAction("PDK_VALOR", { |oView| F3GetFile(oView) } )

//-- Inclusao dos componentes Field e Grids				
oView:AddField( "VIEW_PDI"  , oStruPDI  , "PDIMASTER")		
oView:AddGrid(  "VIEW_PDJ"  , oStruPDJ  , 'PDJDETAIL' )
oView:AddGrid(  "VIEW_PDK"  , oStruPDK  , 'PDKDETAIL' )
                            
oView:AddOtherObject("OTHER_PANEL",{|oPanel| WebPreview(oPanel) } )

oView:CreateHorizontalBox( "BOX", 100)  

//--
oView:CreateFolder( "PASTA", "BOX" )
oView:AddSheet( "PASTA", "ABA01", "Configuracao Modelo")   
oView:AddSheet( "PASTA", "ABA02", "Parametrizacao MSCB")   
oView:AddSheet( "PASTA", "ABA03", "Preview" ) 

oView:CreateHorizontalBox( 'ID_ABA01'   , 100,,, 'PASTA', 'ABA01' )

oView:CreateHorizontalBox( 'ID_ABA02_SUPERIOR'   , 50,,, 'PASTA', 'ABA02' ) 
oView:CreateHorizontalBox( 'ID_ABA02_INFERIOR'   , 50,,, 'PASTA', 'ABA02' ) 

oView:CreateHorizontalBox( 'ID_ABA03'   , 100,,, 'PASTA', 'ABA03' ) 


// Relaciona o identificador (ID) da View com o "box" para exibição
oView:SetOwnerView("VIEW_PDI","ID_ABA01")
oView:SetOwnerView("VIEW_PDJ","ID_ABA02_SUPERIOR")
oView:SetOwnerView("VIEW_PDK","ID_ABA02_INFERIOR")
oView:SetOwnerView('OTHER_PANEL','ID_ABA03')

oView:SetDescription( "Modelo Configurador Etiquetas (Zebra)" ) 

// Liga a identificacao do componente
oView:EnableTitleView( 'VIEW_PDI' )
oView:EnableTitleView( 'VIEW_PDJ' )
oView:EnableTitleView( 'VIEW_PDK' )

//-- AutoIncremental do Item
oView:AddIncrementField('VIEW_PDJ','PDJ_ITEM' )

//-- Altera visual do Grid
oView:SetViewProperty("VIEW_PDJ", "GRIDROWHEIGHT", {14})  
oView:SetViewProperty("VIEW_PDK", "GRIDROWHEIGHT", {14})  

oView:SetViewProperty( "VIEW_PDJ", "SETCSS", { GetCssGrid() } )
oView:SetViewProperty( "VIEW_PDK", "SETCSS", { GetCssGrid() } )

//-- Para não tirar a ordem dos registros  
oView:SetViewProperty("VIEW_PDJ", "ENABLENEWGRID")
oView:SetViewProperty("VIEW_PDK", "ENABLENEWGRID")
oView:SetViewProperty("VIEW_PDJ", "GRIDNOORDER")
oView:SetViewProperty("VIEW_PDK", "GRIDNOORDER")


//-- Valida folder ao seleciona-la
oView:SetVldFolder({|cFolderID, nOldSheet, nSelSheet| ValidFolder(cFolderID, nOldSheet, nSelSheet)})

// oView:SetViewProperty("VIEW_PDI","SETLAYOUT" ,{ FF_LAYOUT_HORZ_DESCR_TOP , -1 } )
// Exibe interface como fosse um webpage
// oView:SetContinuousForm(.T.)
// oView:SetViewProperty("VIEW_PDI", "SIZEMEMO", {"PDI_MEMO" , {10, 400}})

//-- Exibe Load ao carregar tela
// oView:SetProgressBar( .T. )

    
//fecha a tela após clicar no botao confirmar
oView:SetCloseOnOk({||.T.}) 

Return( oView )

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ActiveModel 
@Sample				: ActiveModel()
@description	    : Validacao da ativação do modelo de dados.					
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ActiveModel(oModel)

Local nOperation	:= oModel:GetOperation()
Local lRet 			:= .T.

If nOperation == MODEL_OPERATION_UPDATE .Or. nOperation == MODEL_OPERATION_DELETE

	If PDI->PDI_PROPRI <> "U"
		Help( ,, 'Help',"ZebraCfg", "Este modelo de etiqueta nao permite Alteração e Exclusão!", 1, 0 )
		lRet := .F.
	EndIf 

EndIf 

Return(lRet)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: X7InsPar()
@description	    : Gatilho responsavel por alimentar o grid de parametros						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function X7InsPar()

Local aSaveLines	:= FWSaveRows()
Local oModel        := FwModelActive()
Local oGridPDJ      := oModel:GetModel("PDJDETAIL")
Local oGridPDK      := oModel:GetModel("PDKDETAIL")
Local cTipo         := oGridPDJ:GetValue("PDJ_TIPO")
Local cFunDesc      := ""


//-- Gatilha a descricao do tipo de funcao
cFunDesc := AllTrim(Posicione("SX5",1,xFilial("SX5")+"W6"+ cTipo ,"X5_DESCRI"))
oGridPDJ:SetValue("PDJ_FUNCAO",cFunDesc)

//-- Gatilha os parametros no grid PDK 
If ( Empty(cTipo) .Or.  ValType(__lExecSx7) == "L" .And.  __lExecSx7 )

    //-- Faz o desbloqueio do grid, permitindo inclusao e delecao
    oModel:GetModel( 'PDKDETAIL' ):SetNoInsertLine( .F. )
    oModel:GetModel( 'PDKDETAIL' ):SetNoDeleteLine( .F. )

    // --------------------------------------+
    // Se for informado outro Tipo de funcao  |
    // recarrega o grid de parametros         |
    // ---------------------------------------+
    If ( __ClearGrid )
        //-- Limpa Grid com as definicoes dos parametros
        ClearGrid( oModel, "PDKDETAIL" )
        __ClearGrid := .F. 
    EndIf 

    oGridPDK:GoLine( oGridPDK:Length() )
    If oGridPDK:IsDeleted() 
        oGridPDK:AddLine()
    EndIf 

    /*
    10 - MSCBBOX
    11 - MSCBLineH 
    12 - MSCBLineV                                                 
    20 - MSCBSAY                                                
    30 - MSCBSAYBAR                                                                                     
    50 - MSCBGRAFIC                                             
    */

    Do Case 
        // -------------+
        // 10 - MSCBBOX | 
        //  ------------+
        Case cTipo == "10" 

            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")          

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X2 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")
    
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y2 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Numero com a expessura em pixel")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")        

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Cor Branco ou Preto  (W ou B) ")
            oGridPDK:SetValue("PDK_VALOR"   ,"B")
            oGridPDK:SetValue("PDK_TIPO"    ,"C")      

        // ----------------------------------+
        // 11 - MSCBLineH  | 12 - MSCBLineV  | 
        // ----------------------------------+
        Case cTipo $ ("11|12|") 

            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")          

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X2 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")
    
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Numero com a expessura em pixel")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")        

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Cor Branco ou Preto  (W ou B) ")
            oGridPDK:SetValue("PDK_VALOR"   ,"B")
            oGridPDK:SetValue("PDK_TIPO"    ,"C")            

        // ---------------+
        // 20 - MSCBSAY   | 
        //  -------------+        
        Case cTipo == "20"

            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")   

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y1 em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Texto a ser impresso")
            oGridPDK:SetValue("PDK_VALOR"   , "Informe o texto")
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Tipo de Rotacao N,R,I,B")
            oGridPDK:SetValue("PDK_VALOR"   , "N")
            oGridPDK:SetValue("PDK_TIPO"    , "C")
            

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Fonte 0,1,2,3,4,5,6,7,8,9,20,21,22,27, A...H")
            oGridPDK:SetValue("PDK_VALOR"   , "1")
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Tamanho da Fonte ")
            oGridPDK:SetValue("PDK_VALOR"   , "0")
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Imprime em reverso quando tiver sobre um box preto?")
            oGridPDK:SetValue("PDK_VALOR"   ,".F.")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Serializa o codigo")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Incrementa quando for serial posito ou negativo")
            oGridPDK:SetValue("PDK_VALOR"   , "")
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Coloca zeros a esquerda no numero serial")
            oGridPDK:SetValue("PDK_VALOR"   , "")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Permite brancos a esquerda e direita ")
            oGridPDK:SetValue("PDK_VALOR"   , "")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

        // ------------------+
        // 30 - MSCBSAYBAR   | 
        //  ----------------+           
        Case cTipo == "30"

            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"1")
            oGridPDK:SetValue("PDK_TIPO"    ,"N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Texto a ser impresso")
            oGridPDK:SetValue("PDK_VALOR"   , "Informe o texto")     
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Tipo de Rotacao N,R,I,B")
            oGridPDK:SetValue("PDK_VALOR"   , "N")     
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Modelo Codigo Barras MB01,MB02,MB03,MB04,MB05,MB06,MB07,MB08")
            oGridPDK:SetValue("PDK_VALOR"   , "MB07")    
            oGridPDK:SetValue("PDK_TIPO"    , "C") 

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Altura codigo barras em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   , "01")   
            oGridPDK:SetValue("PDK_TIPO"    , "N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Imprimi Digito Verificacao")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Imprime a linha de código")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")       
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Imprime a linha de código acima das barras")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")   
            oGridPDK:SetValue("PDK_TIPO"    , "L")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Utilizado no code128")
            oGridPDK:SetValue("PDK_VALOR"   , "")   
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Largura da barra mais fina em pontos default 3")
            oGridPDK:SetValue("PDK_VALOR"   , "")   
            oGridPDK:SetValue("PDK_TIPO"    , "N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Relacao entre as barras finas e grossas em pontos default 2")
            oGridPDK:SetValue("PDK_VALOR"   , "2")   
            oGridPDK:SetValue("PDK_TIPO"    , "N")

            // Compacta			Array of Record			Parâmetro fora de uso
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Parâmetro lCompacta fora de uso")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")   
            oGridPDK:SetValue("PDK_TIPO"    , "L")       
            
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Serializa o código")
            oGridPDK:SetValue("PDK_VALOR"   , ".F.")   
            oGridPDK:SetValue("PDK_TIPO"    , "L")      
  
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Incrementa quando for serial positivo ou negativo")
            oGridPDK:SetValue("PDK_VALOR"   , "")   
            oGridPDK:SetValue("PDK_TIPO"    , "C")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   , "Permite brancos a esquerda e direita ")
            oGridPDK:SetValue("PDK_VALOR"   , ".T.")   
            oGridPDK:SetValue("PDK_TIPO"    , "L")

        // -----------------+
        // 50 - MSCBGRAFIC | 
        //  ---------------+  
        Case cTipo == "50"

            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao X em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    , "N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Posicao Y em Milimetros")
            oGridPDK:SetValue("PDK_VALOR"   ,"01")
            oGridPDK:SetValue("PDK_TIPO"    , "N")

            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Definir imagem Grafica")
            oGridPDK:SetValue("PDK_VALOR"   ,"SIGA.GRF")
            oGridPDK:SetValue("PDK_TIPO"    , "I")
            
            oGridPDK:AddLine()
            oGridPDK:SetValue("PDK_PARAM"   ,"Imprime em reverso quando tiver sobre um box preto")
            oGridPDK:SetValue("PDK_VALOR"   ,".F.")
            oGridPDK:SetValue("PDK_TIPO"    , "L")

        // --------------------------------------------------------------+
        // Quando o tipo informado for diferente de SX5W6, limpa o grid. | 
        //  -------------------------------------------------------------+  
        OtherWise 

            //-- Limpa Grid com as definicoes dos parametros
            ClearGrid( oModel, "PDKDETAIL" )
                
    EndCase
    
    
    //-- Retoma o bloqueio da manutencao: Bloqueia inclusao e delecao de linhas 
    oModel:GetModel( 'PDKDETAIL' ):SetNoInsertLine( .T. )
    oModel:GetModel( 'PDKDETAIL' ):SetNoDeleteLine( .T. )

EndIf 

FwRestRows(aSaveLines)

Return(cTipo)



//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: ()
@description	    : 						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ClearGrid(oModel, cSubModel )

Local nI,nJ      := 0
Local nTamHeader := 0

Default oModel   := FWModelActive()
	
oSubModel	:= oModel:GetModel(cSubModel)
nTamHeader	:= LEN(oSubModel:aHeader)

//  EXEMPLO 01 DE DELECAO DE GRID
// oSubModel:GoLine(1)

// ASIZE(oSubModel:aDataModel, 1)
// ASIZE(oSubModel:aCols, 1)	

// For nI := 1 To nTamHeader
//     oModel:ClearField( cSubModel , oSubModel:aHeader[nI][2])
// Next nI

// //AddLine força um refresh no grid, os aSizes removem a nova linha em branco.
// oSubModel:AddLine()
// ASIZE(oSubModel:aDataModel, 1)
// ASIZE(oSubModel:aCols, 1)
    
// EXEMPLO 02 DE DELECAO DE GRID
// oSubModel:cleardata(.F.) // Limpa o Grid
// oSubModel:InitLine()

// EXEMPLO 03 DE DELECAO DE GRID
//---------------------------------------------------------------------+
// Grid: Limpa o Grid para ser recarregado com os trechos da linha.    |
//---------------------------------------------------------------------+	
// For nI := oSubModel:Length() To 1 Step -1
//     //-- Exclui tudo menos a primeira que nao é permitido.
//     If nI > 1
//         oSubModel:GoLine(nI)
//         oSubModel:DeleteLine()
//     // Else 
//     //     For nJ:= 1 To nTamHeader
//     //         oModel:ClearField( cSubModel , oSubModel:aHeader[nJ][2])
//     //     Next 
//     EndIf 

// Next 

oSubModel:DelAllLine()

Return Nil


Static Function GetCssGrid()

Local cGridCSS := ""


cGridCSS := " QTableView { "                        +;
            " background-color: #FFFFFF; "          +; // Branco
            " color: #4D4D4D; "                     +; // Cinza Escuro
            " alternate-background-color: #FFFFFF; "+; // //FAFAFA
            " selection-background-color: #B0E0E6; "+; // //0091FF
            " selection-color: #000000; "           +; // Preto
            " border: 1px solid #C0C0C0; "          +; // D3D3D3
            " font: 10px Helvetica; "              +;				
            "} "

cGridCSS += " QHeaderView::Section { "       +;
            " background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFFFFF, stop:0.3 #F2F2F2, stop:1 #D9D9D9); "+; // Cinza
            " color: #000000; "             +; // Preto
            " border: 1px solid #D3D3D3; "  +;
            " font: 11px Helvetica; "       +;
            " font-weight: bold; "          +;
            " height: 10px; "

// cGridCSS := " QTableView { "                        +;
//             " selection-background-color: #B0E0E6; "+; // //0091FF
//             " border: 2px solid #808080; "          +; // C0C0C0
//             " font: 10px Helvetica; "              +;				
//             "} "

// cGridCSS += " QHeaderView::Section { "       +;
//             " background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFFFFF, stop:0.3 #F2F2F2, stop:1 #D9D9D9); "+; // Cinza
//             " color: #000000; "             +; // Preto
//             " border: 1px solid #D3D3D3; "  +;
//             " font: 11px Helvetica; "       +;
//             " font-weight: bold; "          +;
//             " height: 10px; "

Return(cGridCSS)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: PreVldLin()
@description	    : Rotina executada na pre validacao do grid de funcoes, seta o flag 
					  para limpar o grid de parametros.						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function PreVldLin(oModel,nLinha,cAcao,cCampo,xValue,xOldValue)

Local lRet := .T.

    //  ------------------------------------+
    //  Regra p/ gatilhar o grid PDK        |
    //  ------------------------------------+
    If  cAcao   == "SETVALUE"       .And. ;
        cCampo  == "PDJ_TIPO"       .And. ;
        ValType(__lExecSx7) == "L"  .And. ;
        !Empty(xValue )             .And. ; 
        xValue  <> xOldValue 

        __lExecSx7 := __ClearGrid  := .T.        
    EndIf 


Return(lRet)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ZEBRACFG 
@Sample				: VldTudOk()
@description	    : Chama a rotina para montagem do bloco de codigo. Bloco com as funcoes 
					  que desenham a atiqueta.  						
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function VldTudOk( oModel )

Local nOperation := oModel:GetOperation()
Local cMemo := ""   
Local lRet  := .T.

If nOperation ==  MODEL_OPERATION_INSERT .Or. nOperation == MODEL_OPERATION_UPDATE

    // ----------------------
    // MONTA CODIGO DA ETIQUETA E GRAVA NO MEMO
    // ----------------------
    If ( lRet := CreateBlock( oModel, @cMemo ) )
        If !( lRet := oModel:GetModel("PDIMASTER"):SetValue("PDI_MEMO",cMemo  ) )
            Help( ,, 'Help',"ZebraCfg", "Nao foi possivel gravar a configuracao da etiqueta!", 1, 0 )
        EndIf 
    Else
        Help( ,, 'Help',"ZebraCfg", "Nao foi possivel gerar o corpo da etiqueta. Verifique a parametrizacao!", 1, 0 )
    EndIf 

EndIf

Return(lRet)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} CreateBlock 
@Sample				: CreateBlock()
@description	    : Cria bloco com as funcoes da atiqueta.
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function CreateBlock( oModel, cMemo )

Local aSaveLines:= FWSaveRows()
Local oField    := oModel:GetModel("PDIMASTER")
Local oGridPDJ  := oModel:GetModel("PDJDETAIL")
Local oGridPDK  := oModel:GetModel("PDKDETAIL")
Local cModelo   := oField:GetValue("PDI_MODELO")
Local cPorta    := oField:GetValue("PDI_PORTA")
Local cTipo     := ""
Local cFuncao   := ""
Local cParametro:= ""
Local cVariaveis:= ""
Local aImage    := {} 
Local cImage    := ""
Local cFormul   := ""
Local cConteudo := ""
Local cCorpo    := ""
Local cAux      := ""
Local n1,n2,nPos:= 0
Local lRet      := .F.

Default cMemo := ""

cModelo := AllTrim(cModelo)
cPorta  := AllTrim(cPorta)

// ---------------------------------------------------------+
// VARRE GRID DE PARAMETROS PARA MONTAR O CORPO DA ETIQUETA |
// ---------------------------------------------------------+
For n1 := 1 To oGridPDJ:Length()

    oGridPDJ:GoLine( n1 )

    If !( oGridPDJ:IsDeleted() )

        cTipo   := AllTrim(oGridPDJ:GetValue("PDJ_TIPO"))
        cFuncao := DePara( cTipo )
                
        For n2 := 1 To oGridPDK:Length()
            
            oGridPDK:GoLine( n2 )

            If !( oGridPDK:IsDeleted() )

                cConteudo := cFormul := AllTrim(oGridPDK:GetValue("PDK_FORMUL") ) 
                
                If !Empty( cFormul )
                    cConteudo   := " __ZPAR" + cValToChar(n1) + cValToChar(n2) 
                    cVariaveis  += " __ZPAR" + cValToChar(n1) + cValToChar(n2) + " := &( '" + cFormul + "' ) " + Enter 

                    cParametro += cConteudo 
                Else
                                 
                    //-- Converte Parametros
                    cConteudo  := AllTrim(oGridPDK:GetValue("PDK_VALOR"))  

                    If !Empty(cConteudo)
                        If oGridPDK:GetValue("PDK_TIPO") == "C" //-- Tipo Char Texto
                            cParametro += "'" + cConteudo  +"'"
                        ElseIf oGridPDK:GetValue("PDK_TIPO") == "I" //-- Tipo Imagem - Char Texto ()
                            
                            //-- Nome do imagem + extensão
                            If aScan(aImage, cConteudo) == 0
                                aAdd( aImage, cConteudo  )
                            EndIf 

                            //-- Nome do imagem  sem extensão
                            nPos := At(".",cConteudo)
                            nPos -= 1    
                            cConteudo := SubStr(cConteudo,1, nPos)
                            cParametro += "'" + cConteudo  +"'"
                            
                        Else 
                            cParametro += cConteudo                        
                        EndIf 
                    Else 
                        cParametro += "Nil" 
                    EndIf 
                EndIf 

                If oGridPDK:Length() > n2
                    cParametro += "," 
                EndIf 
            EndIf 
        Next 

        //-----------
        cCorpo      += cFuncao + "(" + cParametro + ")" + Enter 
        cFuncao     := ""
        cParametro  := ""

    EndIf 

Next 

// -----------------------------------------+
// MONTA A ESTRUTURA COMPLETA DA ETIQUETA   |
// -----------------------------------------+
If !Empty(cCorpo)
	
	//-- Instancia objeto MSCBPrinter 
	cMemo += " __oPrint := MSCBPrinter():New() " + Enter  

    //-- - Configura Impressora 
    cMemo += " __oPrint:MSCBPRINTER( '" + cModelo + "','" + cPorta + "',,,.F.,,,,) " + Enter

    //-- Seta ou visualiza o status do sistema com a impressora 
    cMemo += " __oPrint:MSCBCHKStatus(.F.) " + Enter


    For n1 := 1 To Len( aImage ) 
        cImage += " __oPrint:MSCBLOADGRF('" + aImage[n1] + "')" + Enter
    Next

    If !Empty(cImage)
        cMemo += cImage
    EndIf 

    //-- Inicio da Imagem da Etiqueta
    cMemo += " __oPrint:MSCBBEGIN(1,6) " + Enter 

    cMemo += cCorpo 
    lRet  := .T.


    //-- Fim da Imagem da Etiqueta
    cMemo += " __oPrint:MSCBEND() " + Enter 

    //-- Finaliza a conexão com a impressora 
    cMemo += " __oPrint:MSCBClosePrinter() " + Enter

EndIf 


cAux  := cVariaveis + Enter + cMemo
cMemo := cAux

FWRestRows(aSaveLines)

Return(lRet)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} DePara 
@Sample				: DePara()
@description	    : Faz o De \ Para. Recebe o codigo da funcao e devolve sua Syntax.  
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function DePara( cTipo )

Local cFuncao   := ""

Do Case 
    Case cTipo == "10"
        cFuncao  := " __oPrint:MSCBBOX"
    Case cTipo == "11"
        cFuncao  := " __oPrint:MSCBLINEH"
    Case cTipo == "12"
        cFuncao  := " __oPrint:MSCBLINEV"
    Case cTipo == "20"
        cFuncao := " __oPrint:MSCBSAY"
    Case cTipo == "30"
        cFuncao  := " __oPrint:MSCBSAYBAR"
    Case cTipo == "50"
        cFuncao := " __oPrint:MSCBGRAFIC"
EndCase 

Return(cFuncao)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} IMPZEBRA 
@Sample				: IMPZEBRA()
@description	    : Executa a impressao da etiqueta
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
User Function IMPZEBRA(nModo)

Local oModel    := Nil 
Local cPDIMEMO  := "" 
Local cErro     := ""

Private  __lRet := .T.

Default nModo       := IMP_SPOOL_PREVIEW 

If Select("PDI") > 0


    If nModo == IMP_SPOOL_PREVIEW

        If FwAlertYesNo("Confirma a impressao da etiqueta?","ZebraCfg")

            FWMsgRun(, {|oSay| __lRet := ExecBloco(PDI->PDI_MEMO,@cErro,oSay) },"Etiqueta Zebra","Imprimindo..." )

            If ( __lRet ) 
                FwAlertInfo("Impressao finalizada","Impressao")
            Else 
                FwAlertInfo(cErro,"Falha na impressao:")
            EndIf 
        EndIf 
    
    ElseIf nModo == IMP_PREVIEW

        oModel := FwModelActive()

        If ValType( oModel  ) == "O" 
            cPDIMEMO := oModel:GetModel("PDIMASTER"):GetValue("PDI_MEMO")
        Else 
            cPDIMEMO := PDI->PDI_MEMO
        EndIf 

        __lRet := ExecBloco(cPDIMEMO,@cErro,Nil,IMP_PREVIEW)
    
    EndIf 

EndIf 


Return(__lRet)


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ExecBloco 
@Sample				: ExecBloco()
@description	    : Executa o bloco de codigo da etiqueta, macro execução.
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ExecBloco(cMemoCod,cErro,oSay,nModo)

Local aLinha    := {}
Local xAux      := ""
Local cBrkLine  := ""
Local cRet      := ""
Local cRetM     := ""
Local cMacro    := ""
Local bErroA    := {|| Nil }   
Local cPnlCmdI  := ""
Local nX        := 0
Local lRet      := .T.

Private  __cErroBlock  := ""

Default cErro   := ""


aLinha := StrTokArr(cMemoCod,Enter)
bErroA := ErrorBlock( { |oErro| CheckErro( oErro,.T. ) } ) 

Begin Sequence
    
    For nX := 1 to Len(aLinha)
        If !Empty(aLinha[nX])
            aLinha[nX]:= Alltrim(aLinha[nX])
            
            If Right(aLinha[nX],1) == ";"
                cBrkLine += SubStr(aLinha[nX],1,len(aLinha[nX])-1)
                Loop
            EndIf

            cMacro := cBrkLine + aLinha[nX]

            // --------------------------------------------------------------------------+
            //-- Tratamento para NAO gerar Spool (Envio da etiqueta para impressora)     |
            //-- Pois os metodos closePrinter e LoadGRF chamam a impressora.             |
            // --------------------------------------------------------------------------+
            If  "__oPrint:MSCBClosePrinter"  $ cMacro  .Or. ;
                "__oPrint:MSCBLOADGRF"       $ cMacro  .And. nModo == IMP_PREVIEW

                Loop

            ElseIf "__oPrint:MSCBEND()" $ cMacro

                // --------------------------------------------------------------------------+
                //-- Tratamento para NAO gerar Spool (Envio da etiqueta para impressora) e   |
                //-- gerar apenas o codigo ZPL,DPL,IPL e assim enviar para o Preview.        |
                // --------------------------------------------------------------------------+
                If nModo == IMP_PREVIEW
                    //-- Substitui o metodo para gerar apenas o ZPL
                    cMacro := "__oPrint:MSCBEND2()"
                EndIf 

                __ZebraCfg := xAux:= &( cMacro )
            Else 
                xAux:= &( cMacro )
                cBrkLine := ""
            EndIf 
            
            If Valtype(xAux) == "C"
                cRet:= xAux
            Else
                If ValType(xAux) <> "U"

                    If ValType(xAux) == "A"
                        cRet := VarInfo("A",xAux,,.F.)
                    ElseIf ValType(xAux) == "N"
                        cRet := Alltrim(Str(xAux))
                    ElseIf ValType(xAux)== "B"
                        cRet := GetCbSource(xAux)
                    Else
                        cRet := AllToChar(xAux)
                    EndIf
                
                EndIf
            EndIf
            cRetM += Valtype(xAux)+ " -> " + cRet + CRLF
        EndIf
    Next

    // -------------------------------------------------------+
    //  ATUALIZA PAGINA WEB DO PREVIEW COM O NOVO CODIGO ZPL  |
    // -------------------------------------------------------+
    __ZebraCfg := AllTrim(__ZebraCfg)
    UpdateHtml(__ZebraCfg)
    __ZebraCfg := ""

End Sequence

ErrorBlock( bErroA ) 

If !Empty(__cErroBlock) 
    __cErroBlock    += " Comando Executado" + CRLF
    __cErroBlock    += cMacro 
    cPnlCmdI        := __cErroBlock 
    cErro           := __cErroBlock
    __cErroBlock    := ""
    lRet            := .F.
Else
    cPnlCmdI := cRetM
EndIf

Return( lRet )

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} CheckErro 
@Sample				: CheckErro()
@description	    : Tratamento de erro do bloco de codigo.
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function CheckErro(oErroArq, lTrataVar)

Local nI := 0

Default lTrataVar   := .T.


If lTrataVar
    If "variable does not exist " $ oErroArq:description
        __cErroBlock := Alltrim(SubStr(oErroArq:description,24)) + " := '' " + CRLF
    Else
        If oErroArq:GenCode > 0
            __cErroBlock := '(' + Alltrim( Str( oErroArq:GenCode ) ) + ') : ' + AllTrim( oErroArq:Description ) + CRLF
        EndIf  
    EndIf
Else
    If oErroArq:GenCode > 0
        __cErroBlock := '(' + Alltrim( Str( oErroArq:GenCode ) ) + ') : ' + AllTrim( oErroArq:Description ) + CRLF
    EndIf 
    nI := 2
    While ( !Empty(ProcName(nI)) )
        __cErroBlock += Trim(ProcName(nI)) +"(" + Alltrim(Str(ProcLine(nI))) + ") " + CRLF
        nI++
    End             
EndIf

    Break

Return 


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} F3GetFile 
@Sample				: F3GetFile()
@description	    : Rotina executada quando o campo recebe foco. Chama janela para 
 					  selecao de imagem.
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function F3GetFile(oView) 

Local oModel    := Nil 
Local oGridPDK  := Nil 
Local aDirImage := {}
Local cDescPar  := ""
Local cImagem   := ""
Local nPos      := 0

Default oView   := FwViewActive()

//Inicializa variáveis
oModel    := FWModelActive()

If ValType(oModel) == "O"

    oGridPDK  := oModel:GetModel("PDKDETAIL")
    cDescPar  := oGridPDK:GetValue("PDK_PARAM")
    cDescPar  := AllTrim(cDescPar)

    //-- Se parametro for definicao de imagem chama interface de selecao
    If ( cDescPar == "Definir imagem Grafica")

        //-- Interface selecao dir img
        If GetImage( @cImagem )

            //-- Recupera apenas o nome do arquivo
            aDirImage   := Separa(cImagem,"\")
            nPos        := aScan(aDirImage, {|x|  "." $ x  }  )
            cImagem     := IIF( nPos > 0, aDirImage[nPos],cImagem)


            oGridPDK:SetValue("PDK_VALOR",cImagem)
            oView:Refresh()
        EndIf 

    EndIf 

EndIf 
    
Return( Nil )

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} GetImage 
@type				: Funcao estatica
@Sample				: GetImage(cImagem)
@description		: Exibe a janela para a escolher a imagem do logo da etiqueta
@Param				: cIamgem
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function GetImage( cImagem )

Local cFolderSrv:= "\zebra"
Local cTipo     := ""
Local lRet      := .F. 


// ---------------------------------------------+ 
// Cria pasta no servidor caso nao exista       |
// ---------------------------------------------+
If !ExistDir(cFolderSrv)
    If ( MakeDir(cFolderSrv) <> 0 )
        cFolderSrv := ""
        ConOut("Não foi possível criar o diretório padrão de imagens. Erro: " + cValToChar( FError() ) )
    EndIf
EndIf 

cTipo := cTipo + "Imagem (*.PCX)    | *.PCX | "
cTipo := cTipo + "Imagem (*.BMP)    | *.BMP | "
cTipo := cTipo + "Imagem (*.GRF)    | *.GRF | "

cImagem := AllTrim( cGetFile( cTipo, 'Selecionar Imagem', 3, cFolderSrv, .T. ) ) 

// ---------------------------------------------+ 
// Envia imagem do diretorio local para servidor|
// ---------------------------------------------+
If !Empty( cImagem )

    If (lRet := File( cImagem ) )

        If(  ":" $ cImagem ) 

            lRet := CpyT2S(  cImagem, cFolderSrv, /*lCompress*/ , /*lChangeCase*/  )

            If !( lRet )
                FWAlertHelp("O sistema falhou ao transferir imagem para servidor!","Contate o adm do sistema.","ZebraCfg")
            EndIf 
        EndIf
    Else 
        FWAlertHelp("Arquivo de imagem nao localizado!","Selecione uma imagem valida.","ZebraCfg")
    EndIf
EndIf 

Return( lRet )
//------------------------------------------------------------------------------------------
/*/{Protheus.doc} WebTracking 
@type				: Funcao estatica
@Sample				: WebPreview(oPanel)
@description		: Adiciona o componente para exibição da pagina Web. (Preview)
@Param				: oPanel - Painel Container MVC
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function WebPreview(oPanel)

Local cPathLoc  := "c:\zebra\"
Local cHTMLSrv  := "zebra\" + AllTrim(PDI->PDI_COD) +  ".htm

__lFlagUrl := .F.

 //-- Informa URL do registro do modelo 
__URLPage := cPathLoc + AllTrim(PDI->PDI_COD) +  ".htm" 

// COPIA O ARQUIVO DO SERVIDOR PARA UMA PASTA LOCAL
CPYS2T(cHTMLSrv, cPathLoc)

oPanel:Align:= CONTROL_ALIGN_ALLCLIENT


// ------------------------------------------
// Cria navegador embedado
// ------------------------------------------
If GetRpoRelease() > "12.1.017"

    PRIVATE oWebChannel := TWebChannel():New()
    nPort := oWebChannel::connect()

    __oWebPage := TWebEngine():New(oPanel, 0, 0, 100, 100,, nPort)
    __oWebPage:bLoadFinished := {|self,__URLPage| conout("Termino da carga do pagina: " + __URLPage) }
    __oWebPage:navigate(__URLPage)
    __oWebPage:Align := CONTROL_ALIGN_ALLCLIENT
Else 
    __oWebPage := TIBrowser():New(00,400,180,100, "" ,oPanel)
    __oWebPage:Align:= CONTROL_ALIGN_ALLCLIENT
    __oWebPage:GoHome()
EndIf 

Return()

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} UpdateHtml 
@type				: Funcao estatica
@Sample				: UpdateHtml()
@description		: Insere codigo ZPL na pagina web Preview
@Param				: oPanel - Painel Container MVC
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function UpdateHtml(__ZebraCfg)

Local oHtml     := Nil 
Local cModelo 	:= "\zebra\ZebraCfgModelo.html"
Local cHtmlDest := "zebra\" + AllTrim(PDI->PDI_COD) +  ".htm" //Destino deve ser .htm pois o metodo :SaveFile salva somente neste formato.
Local cPathLoc  := "c:\zebra\"


oHTML := TWFHTML():New( cModelo )
oHTML:ValByName( "cCodZpl", __ZebraCfg )
oHTML:SaveFile(cHtmlDest)

// COPIA O ARQUIVO DO SERVIDOR PARA UMA PASTA LOCAL
CPYS2T(cHtmlDest, cPathLoc)


Return()

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ChangePict 
@type				: Funcao estatica
@Sample				: ChangePict(oBrowse)
@description		: Faz a troca das imagens ao navegar pelas linha do Browse
@Param				: oBrowse
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ChangePict( oBrowse )

Local cAlias    := ""
Local cImage    := ""
Local nPos      := 0

Default oBrowse := Nil 

If ValType(oBrowse) == "O" .And. ValType(__oImagem) == "O" .And. Valtype(__aIMGZebra) == "A"

    
    cAlias := oBrowse:GetAlias() 

    nPos := aScan(__aIMGZebra, {|x| (cAlias)->PDI_COD $ x[1]   } )

    If nPos > 0
        cImage := __aIMGZebra[nPos][1]
        __oImagem:Load(NIL, "C:\zebra\imagens\" + cImage )
    Else
        __oImagem:Load(NIL,"C:\zebra\imagens\000000.png")
    EndIf 

    oTScrollBox:Reset()
    oBrowse:Refresh(,.F.,.F.)
    
EndIf 

Return( oBrowse )

//---------------------------------------------------------------------
/*/{Protheus.doc} GetColumns
@Sample	GetColumns()
	Rotina responsavel por montar a estrutura das colunas do Browse.
	
@Param		cAlias
@Return 	aColumns Estrutura de colunas do Browse - FwFormBrowse
@Author		lucas.Brustolin
@Since		05/10/2019	
@Version	12.1.17
/*/
//--------------------------------------------------------------------- 
Static Function GetColumns()

Local aArea	:= GetArea()
Local cCampo	:= ""
Local aCampos	:= {}
Local aColumns	:= {}
Local nX		:= 0
Local nLinha	:= 0
Local cIniBrw	:= ""
Local aCpoQry	:= {}
Local cAlias    := "PDI"


aCampos := {'PDI_COD'	, ; 
			'PDI_DESC'  , ; 
			'PDI_MODELO', ; 	
			'PDI_PORTA' }
			
DbSelectArea("SX3")
DbSetOrder(2)//X3_CAMPO

AAdd(aColumns,FWBrwColumn():New())
nLinha := Len(aColumns)
aColumns[nLinha]:SetData(&(  "{ || IIF( PDI->PDI_PROPRI == 'U','BMPUSER', 'ENGRENAGEM') } "))
aColumns[nLinha]:SetTitle("")
aColumns[nLinha]:SetType("C")
aColumns[nLinha]:SetPicture("@BMP")
aColumns[nLinha]:SetSize(1)
aColumns[nLinha]:SetDecimal(0)
aColumns[nLinha]:SetDoubleClick({|| LegendBrw() })
aColumns[nLinha]:SetImage(.T.)


For nX := 1 To Len(aCampos)
	If SX3->(DbSeek(AllTrim(aCampos[nX])))
		If (X3USO(SX3->X3_USADO) .AND. SX3->X3_BROWSE == "S" .AND. SX3->X3_TIPO <> "M") .OR. SX3->X3_CAMPO = "PDI_FILIAL"
			AAdd(aColumns,FWBrwColumn():New())
			nLinha	:= Len(aColumns)
			cCampo 	:= AllTrim(SX3->X3_CAMPO)
			cIniBrw := AllTrim(SX3->X3_INIBRW)
			aColumns[nLinha]:SetType(SX3->X3_TIPO)
			If SX3->X3_CONTEXT <> "V"
				aAdd(aCpoQry,cCampo)
				If SX3->X3_TIPO = "D"
					aColumns[nLinha]:SetData( &("{|| sTod("  + "('"+cAlias+"')->" + cCampo + ") }") )
				ElseIf !Empty(X3CBox())
					aColumns[nLinha]:SetData( &("{|| X3Combo('" +  cCampo + "',('"+cAlias+"')->" + cCampo + ") }") )
				Else
					aColumns[nLinha]:SetData( &("{|| " + "('"+cAlias+"')->" + cCampo + " }") )
				EndIf
			Else
				aColumns[nLinha]:SetData( &("{|| U_MLRetBrw(" + "'"+cIniBrw+"','"+cAlias+"'" + ") }") )
			EndIf
			aColumns[nLinha]:SetTitle(X3Titulo())
			aColumns[nLinha]:SetSize(SX3->X3_TAMANHO)
			aColumns[nLinha]:SetDecimal(SX3->X3_DECIMAL)

		EndIf

	EndIf
Next nX


RestArea(aArea)

Return(aColumns)

//---------------------------------------------------------------------
/*/{Protheus.doc} LegendBrw
@Sample	LegendBrw()
	Monta interface com a legenda do browse
@Param		Null
@Return 	Null
@Author		lucas.Brustolin
@Since		31/10/2019	
@Version	12.1.17
/*/
//--------------------------------------------------------------------- 
Static Function LegendBrw()

Local oLegenda  :=  FWLegend():New()

oLegenda:Add("","ENGRENAGEM"    ,"Modelo gerado pelo sistema")
oLegenda:Add("","BMPUSER" 	    ,"Modelo gerado pelo usuario")

oLegenda:Activate()
oLegenda:View()
oLegenda:DeActivate()

Return Nil


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} CheckDirFile 
@type				: Funcao estatica
@Sample				: CheckDirFile(oBrowse)
@description		: Checa a existencia das pastas\arquivos 
@Param				: nType
@Param				: __aIMGZebra
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function CheckDirFile(nType,__aIMGZebra)

Local aDiretorio    := {}
Local aFile         := {}
Local aImagens      := {}
Local cArquivo      := ""
Local cPath         := ""
Local cErro         := ""
Local nI            := 0
Local lRet          := .T.

If nType == ARQUIVOS_PASTAS_PREVIEW_BROWSE

    //-- Servidor
    aAdd( aDiretorio, "\zebra" ) 
    aAdd( aDiretorio, "\zebra\imagens\" ) 

    //-- Remote Local 
    aAdd( aDiretorio, "c:\zebra" ) 
    aAdd( aDiretorio, "c:\zebra\imagens\" ) 

ElseIf  nType == ARQUIVOS_PASTAS_PREVIEW_ZPLVIEW
    //-- Servidor
    aAdd( aDiretorio, "\zebra" ) 
    aAdd( aDiretorio, "\zebra\ZPL Viewer_files\" ) 

    //-- Remote Local 
    aAdd( aDiretorio, "c:\zebra" ) 
    aAdd( aDiretorio, "c:\zebra\ZPL Viewer_files\" ) 

EndIf 


For nI := 1 To Len( aDiretorio )

    cPath := aDiretorio[nI] 

    If !ExistDir( cPath )
        If ( MakeDir(cPath) <> 0 )

            If ( lRet )
                cErro   := "Nao foi possivel criar o diretorio padrao. Erro: " + cValToChar( FError() ) + Enter  
            EndIf 

            cErro   += Lower(cPath) + Enter
            lRet    := .F.  
        EndIf 
    EndIf 
Next 

If ( lRet )

    If nType == ARQUIVOS_PASTAS_PREVIEW_BROWSE

        aImagens := Directory("\zebra\imagens\*.")
            
        For nI := 1 To Len(aImagens)

            cArquivo := aImagens[nI][1]

            //-- Arquivo x Destino 
            aAdd( aFile, {"\zebra\imagens\" + cArquivo, "C:\zebra\imagens\"} )  
        Next

    ElseIf nType == ARQUIVOS_PASTAS_PREVIEW_ZPLVIEW

        //-- Arquivo x Destino 
        aAdd( aFile, {"\zebra\ZebraCfgModelo.html"                              , "C:\zebra"} )
        aAdd( aFile, {"\zebra\ZPL Viewer_files\analytics.js.download"           , "C:\zebra\ZPL Viewer_files\"} )
        aAdd( aFile, {"\zebra\ZPL Viewer_files\bootstrap.min.css"               , "C:\zebra\ZPL Viewer_files\"} )
        aAdd( aFile, {"\zebra\ZPL Viewer_files\bootstrap.min.js.download"       , "C:\zebra\ZPL Viewer_files\"} )
        aAdd( aFile, {"\zebra\ZPL Viewer_files\font-awesome.css"                , "C:\zebra\ZPL Viewer_files\"} )
        aAdd( aFile, {"\zebra\ZPL Viewer_files\jquery-1.11.1.min.js.download"   , "C:\zebra\ZPL Viewer_files\"} )    

    EndIf 

    For nI := 1 To Len(aFile)

        cArquivo    := aFile[nI][1]
        cPath       := aFile[nI][2]

        If File(cArquivo)
            If !( CpyS2T( cArquivo, cPath, .T.) )
                //-- Erro na Transferencia 
                cErro   := "Nao foi possível copiar o arquivo: " + cArquivo + " para: " + cPath + " . Verifique se voce possui acesso de escrita" 
                lRet    := .F. 
                Exit                
            EndIf 
        Else 
            cErro   := "Nao foi localizar o arquivo: "+ cArquivo + " . Contate o administrador do sistema"
            lRet    := .F. 
            Exit
        EndIf 
    Next 
EndIf 

If !( lRet )
    FwAlertWarning(cErro,"ZebraCfg")
Else 
    __aIMGZebra:= Directory("c:\zebra\imagens\*.")
EndIf 

Return(lRet)

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ValidFolder 
@type				: Funcao estatica
@Sample				: ValidFolder(cFolderID, nOldSheet, nSelSheet)
@description		: Executa ação no click da Folder. Carrega pagina web ao clicar na 
                      folder corresponde.
@Param				: cFolderID
@Param				: nOldSheet
@Param              : nSelSheet
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ValidFolder(cFolderID, nOldSheet, nSelSheet)

Local aArea         := GetArea()
Local oModel        := FwModelActive()
Local oView         := FwViewActive()
Local nOperation    := oModel:GetOperation()
Local cModelo       := ""
Local lZPLView      := .F.
Local lUpdate       := .F. 
Local lRet          := .T. 

    // Pastas 1 Configuracao Modelo" 
    // Pastas 2 Parametrizacao MSCB 
    // Pastas 3 Preview  

    cModelo     := oModel:GetModel("PDIMASTER"):GetValue("PDI_MODELO")
    lZPLView    := IsModelZpl( cModelo )

    // ------------------------------------------------------------------------+
    //  Ação ao clicar na pasta PreView - Carrega pagina com o novo codigo ZPL |
    // -------------------------------------------------------------------------+
    If nSelSheet == 3 .And. ( !lZPLView )

        Help(" ",1,"ZEBRACFg",,"O Preview só está disponivel para modelos de impressora Zebra",3,1) 
        lRet := .F.

    ElseIf nSelSheet == 3 .And. ( nOperation == MODEL_OPERATION_INSERT .Or. nOperation == MODEL_OPERATION_UPDATE ) 

        lUpdate       := oModel:lModify

        If ( lUpdate )  
            //-- 
            FWMsgRun( ,{|| ProcFolder(oModel,oView)  },,'Atualizando... Web Preview')        
        EndIf
        
    EndIf


    If nSelSheet == 3 .And. lZPLView .And. !( __lFlagUrl ) 
        __oWebPage:Navigate( __URLPage )
        __lFlagUrl  := .T.
    EndIf 


RestArea( aArea )
Return( lRet )

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} ProcFolder 
@type				: Funcao estatica
@Sample				: ProcFolder()
@description		: Atualiza a pasta Preview 
@Param				: 
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function ProcFolder(oModel,oView)

Local cMemo := ""

    CreateBlock( oModel, @cMemo )

    oModel:GetModel("PDIMASTER"):SetValue("PDI_MEMO",cMemo  ) 

    U_IMPZEBRA( IMP_PREVIEW )

    __oWebPage:Navigate( __URLPage )

    oView:Refresh("VIEW_PDI")
    oView:Refresh("OTHER_PANEL")

Return()


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} GravaModel 
@type				: Funcao estatica
@Sample				: GravaModel(oSay,cPK, aData, cError)
@description		: Grava os modelos de etiqueta 000001,000002 e 000003. 
@Param				: 
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 01/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function GravaModel(oSay,cPK, aData, cError) 

Local oModel    := FwLoadModel("ZEBRACFG") 
Local lXml      := .F.
Local lRet      := .T.
Local cData	    := ""
Local nI        := 0

Default cPK     := ""
Default aData   := {"000001","000002","000003"} 
Default cError  := ""


DbSelectArea("PDI")
DbSetOrder(1)

For nI := 1 To Len(aData)

    If PDI->( !DbSeek(xFilial("PDI") + aData[nI] ) )

        oSay:cCaption := "Incluindo modelo: " + aData[nI]
        ProcessMessages()

        If Empty(cPk)
            oModel:SetOperation(MODEL_OPERATION_INSERT)
        Else
            oModel:SetOperation(MODEL_OPERATION_UPDATE)
            lRet := Seek(cPK)
        EndIf

        cData := MemoRead("\zebra\"+ aData[nI] +".json")

        If ( lRet .And. !Empty(cData) )
            oModel:Activate()

            If lXml
                lRet := oModel:LoadXMLData(cData)
            Else
                lRet := oModel:LoadJsonData(cData)
            EndIf

            If lRet
                If oModel:lModify // Verifico se o modelo sofreu alguma alteração
                    If !(oModel:VldData() .And. oModel:CommitData())
                        lRet := .F.
                        cError := ErrorMessage(oModel:GetErrorMessage())
                    EndIf
                Else
                    lRet    := .F.
                    cError  := "Not Modified"
                EndIf
            Else
                cError := oModel:GetErrorMessage()
            EndIf
            oModel:DeActivate()
        Else
            cError := i18n("Invalid record '#1' on table #2", {cPK, "PDI"})
        EndIf
    EndIf
Next 

Return( lRet )


//------------------------------------------------------------------------------------------
/*/{Protheus.doc} IsModelZpl 
@type				: Funcao estatica
@Sample				: IsModelZpl(cPesq)
@description		: Verifica se o modelo de impressora configurado é ZEBRA.
@Param				: cPesq - Modelo da impressora termica
@return				: .T. / .F.
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 22/11/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
Static Function IsModelZpl(cPesq)

Local aModelos  := {}
Local lRet      := .T. 

Default cPesq := ""

If !Empty(cPesq)
    
    cPesq := AllTrim(UPPER(cPesq))

    aModelos    := {'S300'     ,;
                    'S400'		,;
                    'S500-6'	,;
                    'S500-8'	,;
                    'Z105S-6'	,;
                    'Z105S-8'	,;
                    'Z160S-6'	,;
                    'Z160S-8'	,;
                    'Z140XI'	,;
                    'S600'		,;
                    'Z4M'		,;
                    'Z90XI'	    ,;
                    'Z170XI'	,;
                    'ZEBRA'     ,;
                    'QL320'    } 


    lRet := ( aScan( aModelos, cPesq ) > 0 )

EndIf 

Return( lRet )
