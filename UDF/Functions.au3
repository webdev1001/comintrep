#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.7.18 (beta)
 Author:         Rizonesoft (Derick Payne)

 Script Function:
	Doors Functions

#ce ----------------------------------------------------------------------------


#include-once


; #FUNCTION# ====================================================================================================
; Name...........: _ReduceMemory
; Description ...: Reduce memory usage of process ID (PID) given.
; Syntax.........: _ReduceMemory($iPID = -1, $hPsAPIdll = 'psapi.dll', $hKernel32dll = 'kernel32.dll')
; Parameters ....: $BugProcess - Process to Close
; Return values .: $iPID - PID of process to reduce memory. If -1 reduce self memory usage.
;                  $hPsAPIdll - Optional handle to psapi.dll.
;                  $hKernel32dll - Optional handle To kernel32.dll.
; Requirement(s) : psapi.dll (Doesn't come with WinNT4 by default)
; Author(s) .....: w0uter,  Saunders (admin@therks.com)
; Modified.......: Derick Payne (Rizonesoft)
; Remarks .......: If @OSVersion = 'WIN_NT4' Then FileInstall('psapi.dll', @SystemDir & '\psapi.dll')
; Link ..........:
; Example .......:
; ===============================================================================================================
Func _ReduceMemory($iPID = -1, $hPsAPIdll = 'psapi.dll', $hKernel32dll = 'kernel32.dll')
    If $iPID <> -1 Then
        Local $aHandle = DllCall($hKernel32dll, 'int', 'OpenProcess', 'int', 0x1f0fff, 'int', False, 'int', $iPID)
        Local $aReturn = DllCall($hPsAPIdll, 'int', 'EmptyWorkingSet', 'long', $aHandle[0])
        DllCall($hKernel32dll, 'int', 'CloseHandle', 'int', $aHandle[0])
    Else
        Local $aReturn = DllCall($hPsAPIdll, 'int', 'EmptyWorkingSet', 'long', -1)
    EndIf

    Return $aReturn[0]
EndFunc


Func _FileWriteAccessible($sFile)
    ; Returns
    ;            1 = Success, file is writeable and deletable
    ;            0 = Failure
    ; @error
    ;            1 = Access Denied because of lacking access rights
    ;            2 = File is set "Read Only" by attribute
    ;            3 = File not found
    ;            4 = Unknown Api Error, check @extended

    Local $iSuccess = 0, $iError_Extended = 0, $iError = 0, $hFile
    ;$hFile = _WinAPI_CreateFileEx($sFile, $OPEN_EXISTING, $FILE_WRITE_DATA, BitOR($FILE_SHARE_DELETE, $FILE_SHARE_READ, $FILE_SHARE_WRITE), $FILE_FLAG_BACKUP_SEMANTICS)
    $hFile = _WinAPI_CreateFileEx($sFile, 3, 2, 7, 0x02000000)
    Switch _WinAPI_GetLastError()
        Case 0 ; ERROR_SUCCESS
            $iSuccess = 1
        Case 5 ; ERROR_ACCESS_DENIED
            If StringInStr(FileGetAttrib($sFile), "R", 2) Then
                $iError = 2
            Else
                $iError = 1
            EndIf
        Case 2 ; ERROR_FILE_NOT_FOUND
            $iError = 3
        Case Else ; w000t?
            $iError = 4
            $iError_Extended = _WinAPI_GetLastError()
    EndSwitch
    _WinAPI_CloseHandle($hFile)
    Return SetError($iError, $iError_Extended, $iSuccess)
EndFunc   ;==>_FileWriteAccessible


	; #FUNCTION# ====================================================================================================================
; Name...........: _WinAPI_CreateFileEx
; Description....: Creates or opens a file or I/O device.
; Syntax.........: _WinAPI_CreateFileEx ( $sFile, $iCreation [, $iAccess [, $iShare [, $iFlagsAndAttributes [, $tSecurity [, $hTemplate]]]]] )
; Parameters.....: $sFile               - The name of the file or device to be created or opened.
;                  $iCreation           - The action to take on a file or device that exists or does not exist. This parameter must be
;                                         one of the following values, which cannot be combined.
;
;                                         $CREATE_NEW
;                                         $CREATE_ALWAYS
;                                         $OPEN_EXISTING
;                                         $OPEN_ALWAYS
;                                         $TRUNCATE_EXISTING
;
;                  $iAccess             - The requested access to the file or device, which can be summarized as read, write, both
;                                         or neither (zero).
;
;                                         $GENERIC_READ
;                                         $GENERIC_WRITE
;
;                                         (See MSDN for more information)
;
;                  $iShare              - The requested sharing mode of the file or device, which can be read, write, both,
;                                         delete, all of these, or none. If this parameter is 0 and _WinAPI_CreateFileEx() succeeds,
;                                         the file or device cannot be shared and cannot be opened again until the handle to
;                                         the file or device is closed.
;
;                                         $FILE_SHARE_DELETE
;                                         $FILE_SHARE_READ
;                                         $FILE_SHARE_WRITE
;
;                  $iFlagsAndAttributes - The file or device attributes and flags. This parameter can be one or more of the
;                                         following values.
;
;                                         $FILE_ATTRIBUTE_READONLY
;                                         $FILE_ATTRIBUTE_HIDDEN
;                                         $FILE_ATTRIBUTE_SYSTEM
;                                         $FILE_ATTRIBUTE_DIRECTORY
;                                         $FILE_ATTRIBUTE_ARCHIVE
;                                         $FILE_ATTRIBUTE_DEVICE
;                                         $FILE_ATTRIBUTE_NORMAL
;                                         $FILE_ATTRIBUTE_TEMPORARY
;                                         $FILE_ATTRIBUTE_SPARSE_FILE
;                                         $FILE_ATTRIBUTE_REPARSE_POINT
;                                         $FILE_ATTRIBUTE_COMPRESSED
;                                         $FILE_ATTRIBUTE_OFFLINE
;                                         $FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
;                                         $FILE_ATTRIBUTE_ENCRYPTED
;
;                                         $FILE_FLAG_BACKUP_SEMANTICS
;                                         $FILE_FLAG_DELETE_ON_CLOSE
;                                         $FILE_FLAG_NO_BUFFERING
;                                         $FILE_FLAG_OPEN_NO_RECALL
;                                         $FILE_FLAG_OPEN_REPARSE_POINT
;                                         $FILE_FLAG_OVERLAPPED
;                                         $FILE_FLAG_POSIX_SEMANTICS
;                                         $FILE_FLAG_RANDOM_ACCESS
;                                         $FILE_FLAG_SEQUENTIAL_SCAN
;                                         $FILE_FLAG_WRITE_THROUGH
;
;                                         $SECURITY_ANONYMOUS
;                                         $SECURITY_CONTEXT_TRACKING
;                                         $SECURITY_DELEGATION
;                                         $SECURITY_EFFECTIVE_ONLY
;                                         $SECURITY_IDENTIFICATION
;                                         $SECURITY_IMPERSONATION
;
;                  $tSecurity           - $tagSECURITY_ATTRIBUTES structure that contains two separate but related data members:
;                                         an optional security descriptor, and a Boolean value that determines whether the returned
;                                         handle can be inherited by child processes. If this parameter is 0, the handle cannot
;                                         be inherited by any child processes the application may create and the file or device
;                                         associated with the returned handle gets a default security descriptor.
;                  $hTemplate           - Handle to a template file with the $GENERIC_READ access right. The template file supplies
;                                         file attributes and extended attributes for the file that is being created.
; Return values..: Success              - Handle to the specified file, device, named pipe, or mail slot.
;                  Failure              - 0 and sets the @error flag to non-zero.
; Author.........: Yashied
; Modified.......:
; Remarks........: When an application is finished using the object handle returned by this function, use the _WinAPI_CloseHandle()
;                  function to close the handle. This not only frees up system resources, but can have wider influence on things
;                  like sharing the file or device and committing data to disk.
; Related........:
; Link...........: @@MsdnLink@@ CreateFile
; Example........: Yes
; ===============================================================================================================================

Func _WinAPI_CreateFileEx($sFile, $iCreation, $iAccess = 0, $iShare = 0, $iFlagsAndAttributes = 0, $tSecurity = 0, $hTemplate = 0)

	Local $Ret = DllCall('kernel32.dll', 'ptr', 'CreateFileW', 'wstr', $sFile, 'dword', $iAccess, 'dword', $iShare, 'ptr', DllStructGetPtr($tSecurity), 'dword', $iCreation, 'dword', $iFlagsAndAttributes, 'ptr', $hTemplate)

	If (@error) Or ($Ret[0] = Ptr(-1)) Then
		Return SetError(1, 0, 0)
	EndIf
	Return $Ret[0]
EndFunc   ;==>_WinAPI_CreateFileEx


Func _GetExecVersioning($sExecPath, $iFlag = 6)

	If FileExists($sExecPath) Then
		Local $verReturn = FileGetVersion($sExecPath)
		Local $splReturn = StringSplit($verReturn, ".")

		If $splReturn[0] >= 4 Then
			If $iFlag = 1 Then
				Return $splReturn[1]
			ElseIf $iFlag = 2 Then
				Return $splReturn[2]
			ElseIf $iFlag = 3 Then
				Return $splReturn[3]
			ElseIf $iFlag = 4 Then
				Return $splReturn[4]
			ElseIf $iFlag = 5 Then
				Return $splReturn[1] & " : Build " & $splReturn[4]
			ElseIf $iFlag = 6 Then
				Return $verReturn
			EndIf
		EndIf
	Else
		Return "0000001" ;ERROR CODE: DOORS VERSIONING: PATH TO EXECUTABLE NOT FOUND
	EndIf
EndFunc ;==>_GetDoorsVersion