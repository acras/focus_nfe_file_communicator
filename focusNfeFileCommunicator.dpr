program focusNfeFileCommunicator;

{$APPTYPE CONSOLE}

uses
  Windows, dialogs,
  focusNfeCommunicator in 'focusNfeCommunicator.pas';

var
  mutex: THandle;
begin
  mutex := CreateMutex(nil, true, 'FocusNFeCommunicator');
  if not((Mutex = 0) OR (GetLastError = ERROR_ALREADY_EXISTS)) then
  begin
    TFocusNFeCommunicator.startProcess;
  end;
end.



