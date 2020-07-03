#include 'totvs.ch'

//-- User para aparecer no inspetor de objetos
User Function MSCBPrinter() ; Return

//------------------------------------------------------------------------------------------
/*/{Protheus.doc} MSCBPrinter 
@type				: Classe
@Sample				: MSCBPrinter():New()
@description	    : Classe cuja os metodos representam as funcoes MSCB(s) para impressao de etiquetas.
                      Criada exclusivamente para obter os recursos das classes filhas como a MSCBZPL, MSCBEPL
					  e entre outras que são instanciadas pela funcao MSCBPrinter.
					  Também foi adicionado um novo metodo de finalização que impede a impressão da etiqueta e
					  gera apenas os codidos ZPL,EPL... 
					  				
@Param				: Nulo
@return				: Nulo
@ ------------------|----------------
@author				: Lucas.Brustolin
@since				: 14/10/2019
@version			: Protheus 12.1.17
/*/
//------------------------------------------------------------------------------------------
CLASS  MSCBPrinter From LongClassName

	Data oMSCB	As Object 
	Data cType	As String

	Method New() CONSTRUCTOR
	Method MSCBPrinter(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) CONSTRUCTOR
	Method GetType() // Retorna o modelo de impressora
	Method MSCBBEGIN(nxQtde,nVeloc,nTamanho,lSalva)
	Method MSCBEND()
	Method MSCBEND2()
	Method MSCBWrite(cConteudo,cModo)
	Method MSCBIsPrinter() 
	Method MSCBClosePrinter()
	Method MSCBCHKStatus(lStatus)
	Method MSCBSAY(nXmm,nYmm,cTexto,cRotacao,cFonte,cTam,lReverso,lSerial,cIncr,lZerosL,lNoAlltrim) 
	Method MSCBVar(cVar,cDados)
	Method MSCBSAYMEMO(nXmm,nYmm,nLMemomm,nQLinhas,cTexto,cRotacao,cFonte,cTam,lReverso,cAlign) 
	Method MSCBSAYBAR(nXmm,nYmm,cConteudo,cRotacao,cTypePrt,nAltura,lDigVer,lLinha,lLinBaixo,cSubSetIni,nLargura,nRelacao,lCompacta,lSerial,cIncr,lZerosL)
	Method MSCBBOX(nX1mm,nY1mm,nX2mm,nY2mm,nExpessura,cCor) 
	Method MSCBLineH(nX1mm,nY1mm,nX2mm,nExpessura,cCor)
	Method MSCBLineV(nX1mm,nY1mm,nY2mm,nExpessura,cCor) 
	Method MSCBGRAFIC(nXmm,nYmm,cArquivo,lReverso) 
	Method MSCBLOADGRF(cImagem)

ENDCLASS


Method New() CLASS MSCBPrinter

	Self:oMSCB	:= Nil 
	Self:cType	:= ""
	
Return( Self )

Method MSCBPrinter(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) CLASS MSCBPrinter

Local aArea:= GetArea()

If MSCbModelo('ZPL',ModelPrt)
	Self:cType := 'ZPL'
   	Self:oMSCB := MSCBZPL():New(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) 
ElseIf MSCbModelo('DPL',ModelPrt)
   	Self:oMSCB := MSCBDPL():New(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) 
   	Self:cType := 'DPL'
ElseIf MSCbModelo('EPL',ModelPrt)
   	Self:oMSCB := MSCBEPL():New(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) 
   	Self:cType := 'EPL'
ElseIf MSCbModelo('IPL',ModelPrt)
   Self:oMSCB := MSCBIPL():New(ModelPrt,cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni) 
   Self:cType := 'IPL'
Else         
	// modelo nao encontado, portanto default zebra com densidade 6  
	Self:oMSCB := MSCBZPL():New("S500-6",cPorta,nDensidade,nTamanho,lSrv,nPorta,cServer,cEnv,nMemoria,cFila,lDrvWin,cPathIni)    
	Self:cType := 'ZPL'
EndIf
                
Self:oMSCB:Setup()

RestArea(aArea)

Return( Self )

Method GetType() CLASS MSCBPrinter
Return(Self:cType)

Method MSCBBEGIN(nxQtde,nVeloc,nTamanho,lSalva) Class MSCBPrinter
	Self:oMSCB:CBBegin(nxQtde,nVeloc,nTamanho,lSalva)
Return('')

Method MSCBEND() Class MSCBPrinter
Return (Self:oMSCB:CBEnd() )


Method MSCBWrite(cConteudo,cModo) CLASS MSCBPrinter

	cModo := If(cModo==NIL,"WRITE",cModo)
	
	If (cModo=="ABRE")
	   Self:oMSCB:cResult :=''
	ElseIf (cModo=="WRITE")
	   Self:oMSCB:cResult +=cConteudo
	ElseIf (cModo=="FECHA")
	   Self:oMSCB:Envia()
	EndIf

Return('')


Method MSCBIsPrinter() CLASS MSCBPrinter

	If ValType(Self:oMSCB) <> "O"
	   Return .F.
	EndIf

Return( Self:oMSCB:IsPrinted )

Method MSCBClosePrinter() CLASS MSCBPrinter
Return( Self:oMSCB:Close() )


Method MSCBCHKStatus(lStatus) CLASS MSCBPrinter

	If lStatus <> NIL
	   Self:oMSCB:lCHKStatus := lStatus
	EndIf         
	     
	If  Self:oMSCB:lDrvWin
	   Self:oMSCB:lCHKStatus := .F.
	EndIf
	
	If Self:oMSCB:lSpool 
	   MSCBGrvSpool(5,,Self:oMSCB)
	EndIf
	
Return( Self:oMSCB:lCHKStatus )

Method MSCBSAY(nXmm,nYmm,cTexto,cRotacao,cFonte,cTam,lReverso,lSerial,cIncr,lZerosL,lNoAlltrim) CLASS MSCBPrinter
	Self:oMSCB:Say(nXmm,nYmm,cTexto,cRotacao,cFonte,cTam,lReverso,lSerial,cIncr,lZerosL,lNoAlltrim)
Return('')

Method MSCBVar(cVar,cDados) CLASS MSCBPrinter
	Self:oMSCB:Var(cVar,cDados)
RETURN('')

Method MSCBSAYMEMO(nXmm,nYmm,nLMemomm,nQLinhas,cTexto,cRotacao,cFonte,cTam,lReverso,cAlign) CLASS MSCBPrinter
	Self:oMSCB:Memo(nXmm,nYmm,nLMemomm,nQLinhas,cTexto,cRotacao,cFonte,cTam,lReverso,cAlign)
Return('')

Method MSCBSAYBAR(nXmm,nYmm,cConteudo,cRotacao,cTypePrt,nAltura,lDigVer,lLinha,lLinBaixo,cSubSetIni,nLargura,nRelacao,lCompacta,lSerial,cIncr,lZerosL) CLASS MSCBPrinter
	Self:oMSCB:Bar(nXmm,nYmm,cConteudo,cRotacao,cTypePrt,nAltura,lDigVer,lLinha,lLinBaixo,cSubSetIni,nLargura,nRelacao,lCompacta,lSerial,cIncr,lZerosL)
Return('')

Method MSCBBOX(nX1mm,nY1mm,nX2mm,nY2mm,nExpessura,cCor) CLASS MSCBPrinter
	Self:oMSCB:Box(nX1mm,nY1mm,nX2mm,nY2mm,nExpessura,cCor)
Return('')

Method MSCBLineH(nX1mm,nY1mm,nX2mm,nExpessura,cCor) CLASS MSCBPrinter
	Self:oMSCB:LineH(nX1mm,nY1mm,nX2mm,nExpessura,cCor)
Return('')

Method MSCBLineV(nX1mm,nY1mm,nY2mm,nExpessura,cCor) CLASS MSCBPrinter
	Self:oMSCB:LineV(nX1mm,nY1mm,nY2mm,nExpessura,cCor)
Return('')

Method MSCBGRAFIC(nXmm,nYmm,cArquivo,lReverso) CLASS MSCBPrinter
	Self:oMSCB:GRAFIC(nXmm,nYmm,cArquivo,lReverso)
Return('')

Method MSCBLOADGRF(cImagem) CLASS MSCBPrinter
Return( Self:oMSCB:LOADGRF(cImagem) )

Method MSCBEND2() CLASS MSCBPrinter

Local cConteudo:= ""

	Self:oMSCB:cResult += "^XZ" + Chr(13) + Chr(10)
	
	If Len(Self:oMSCB:cResult) > Self:oMSCB:nMemory
	   cConteudo := Self:oMSCB:cResult
	EndIF   

Return cConteudo