'---------------------------------------------------------------------------
'     ###          #       ###   ###     #####                        ##
'    #   #         #      #   #  # #      #  #       0ut0fBound        #
'    #   #         #      #   #  #        #  #                         #
'    #   # ## ##  ###     #   # ###       ###   ###  ## ##  #####   ####
'    #   #  #  #   #      #   #  #        # ## #   #  #  #   #  #  #   #
'    #   #  #  #   #      #   #  #        #  # #   #  #  #   #  #  #   #
'    #   #  #  #   #      #   #  #        #  # #   #  #  #   #  #  #   #
'     ###   #####  ##      ###  ###      #####  ###   ##### ######  #####
'                                             http://maycon.hacknroll.com
'------------------------------------------------------------------------
'                                                        Recovery: Tiny/H
'                                                        Date: 02/07/2008
'                                                Author: Maycon M. Vitali
'                                           Contact: maycon@hacknroll.com
'------------------------------------------------------------------------

'-------------------------------
' Aonde ser� aplicado o recovery
'-------------------------------
strComputer = "."


'------------------------------------------
' Inicia os servico necessarios (RMI e FSO)
'------------------------------------------
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'----------------
' Tipos de Discos
'----------------
Const REMOVABLE_DRIVER = 2

'-------------------------------
' Tipos de Permi��es de Arquivos
'-------------------------------
Const FILE_ATTRIBUTE_NORMAL   = 128


'---------------------------------------------------------------------
' Verifica a possivel infec��o em um dispositivo, tendo como parametro
' a letra do dispositivo e verificando atrav�s da existencia dos
' arquivos deixados pela infec��o
'---------------------------------------------------------------------
function DispositivoInfectado(cDispositivo)
    DispositivoInfectado = objFSO.FileExists(cDispositivo + "\autorun.inf") AND _
                           objFSO.FileExists(cDispositivo + "\explorer.exe") AND _
                           objFSO.FileExists(cDispositivo + "\fooool.exe")
end function



'----------------------------------------------------------------------
' Esta fun��o restaura as permi��es dos arquivos de infec��o para a
' NORMAL, pois os mesmo ficam com as permi��es SYSTEM, HIDDEN e ARCHIVE
'----------------------------------------------------------------------
sub RemovePermicoes(cArquivo)
    Wscript.Echo "  > " & cArquivo
    Set ObjFile = objFSO.GetFile(cArquivo)
    objFile.Attributes = FILE_ATTRIBUTE_NORMAL
end sub



'----------------------------------------------------------------------
' Esta fun��o � respons�vel por remover os processos relacionados aos
' arquivos que identificam os virus
'----------------------------------------------------------------------
sub RemoveProcessos(cCaminho)
    Wscript.Echo "  > Caminho: " & cCaminho
    Set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process")
    For Each objProcess in colProcessList
        if objProcess.ExecutablePath = cCaminho then
            objProcess.Terminate()
            Wscript.Echo "    > PID: " & objProcess.ProcessId & " morto"
        end if
    next
end sub



'----------------------------------------------------------------
' Esta fun��o remove um arquivo ( no caso vir�tico ) passado como
' parametro
'----------------------------------------------------------------
sub RemoveArquivo(cArquivo)
	objFSO.DeleteFile(cArquivo)
	if objFSO.FileExists(cArquivo) then
	    Wscript.Echo "  > Arquivo '" & cArquivo & "' N�O removido"
    else
	    Wscript.Echo "  > Arquivo '" & cArquivo & "' removido"
	end if
end sub


'------------------------------------------------
' Fun��o respons�vel pelo processo de desinfec��o
'------------------------------------------------
function Desinfecta(cDispositivo)
    Wscript.Echo "----------------------------------"
    Wscript.Echo "======= Aplicando Recovery ======="
    Wscript.Echo "----------------------------------"

    Wscript.Echo "> Restaurando Permi��es para Original"
    RemovePermicoes cDispositivo & "\autorun.inf"
    RemovePermicoes cDispositivo & "\explorer.exe"
    RemovePermicoes cDispositivo & "\fooool.exe"
    Wscript.Echo ""

    Wscript.Echo "> Finalizando Processos Dependentes"    
    RemoveProcessos cDispositivo + "\explorer.exe"
    RemoveProcessos cDispositivo + "\fooool.exe"
    Wscript.Echo ""
	
    Wscript.Echo "> Apagando Arquivos"    
    RemoveArquivo cDispositivo + "\explorer.exe"
    RemoveArquivo cDispositivo + "\fooool.exe"
    Wscript.Echo ""
	
end function


'-------------------------------------------------------------
' Busca todos os dispositivos removiveis a procura da infeccao
'-------------------------------------------------------------
Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & REMOVABLE_DRIVER & "")


Wscript.Echo "----------------------------------"
Wscript.Echo "====== Verificando Infec��o ======"
Wscript.Echo "----------------------------------"


boolInfected = False
For Each objDisk in colDisks

    if objDisk.DeviceID <> "A:" then ' N�o vale disquete :P

        '----------------------------------------------
        ' Verifica a existencia da infeccao nos drivers
        '----------------------------------------------
        if DispositivoInfectado(objDisk.DeviceID) then

            Wscript.Echo "> Poss�vel infec��o TINY/H em (" + objDisk.DeviceID + ")"
            Wscript.Echo "  > " + objDisk.DeviceID + "\autorun.inf"
            Wscript.Echo "  > " + objDisk.DeviceID + "\explorer.exe"
            Wscript.Echo "  > " + objDisk.DeviceID + "\fooool.exe"
            Wscript.Echo ""

            Desinfecta objDisk.DeviceID



            boolInfected = True

        end if

    end if

Next

if not(boolInfected ) then
    Wscript.Echo "> Nenhum dispositivo remov�vel infectado"
end if


Wscript.Sleep 5000