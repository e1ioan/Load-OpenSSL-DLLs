unit LoadSSLDll;

{------------------------------------------------------------------------------
  Unit Name: LoadSSLDll
  Purpose  : Provides functionality to dynamically load DLL libraries from a
             specified directory at runtime, with a specific method to load
             OpenSSL libraries based on the system architecture (32-bit or
             64-bit).

  Description:
    - This unit offers a general procedure `LoadLibrariesFromDirectory` that
      loads all DLL files from a given directory.
    - It also provides `LoadOpenSSLLibraries`, which utilizes the general
      procedure to load OpenSSL libraries from a directory determined by the
      system architecture.
    - Error handling is implemented to log any failures in loading the DLLs,
      and the application will terminate if critical libraries cannot be loaded.

  Usage:
    - Include this unit at the beginning of your project's `.dpr` file to ensure
      that the OpenSSL libraries are loaded before any other units that depend on them.
    - Example:

      ```delphi
      program MyApplication;

      uses
        LoadSSLDll, // Ensure this unit is listed first
        Forms,
        MainUnit;

      begin
        Application.Initialize;
        Application.CreateForm(TMainForm, MainForm);
        Application.Run;
      end.
      ```

    - The unit automatically loads the appropriate OpenSSL libraries during
      initialization, so no additional code is needed in your application.
    - If you need to load libraries from a different directory, you can call
      `LoadLibrariesFromDirectory` with the desired path.

------------------------------------------------------------------------------}

interface

uses
  Windows,
  System.IOUtils,
  System.Types,
  System.SysUtils;

{**
  Loads all DLL libraries from the specified directory.

  @param Directory The path to the directory containing the DLL files to load.
*}
procedure LoadLibrariesFromDirectory(const Directory: string);

{**
  Loads the OpenSSL libraries based on the system architecture (32-bit or 64-bit).
*}
procedure LoadOpenSSLLibraries;

implementation

uses
  SvcMgr,
  avglobal,
  avfiles;

procedure LoadLibrariesFromDirectory(const Directory: string);
var
  libHandles: array of HMODULE;
  fName: string;
  Index: Integer;
begin
  // Temporarily set the DLL directory to the specified path
  SetDllDirectory(PWideChar(Directory));
  Index := 0;

  // Iterate over all DLL files in the directory
  for fName in TDirectory.GetFiles(IncludeTrailingPathDelimiter(Directory), '*.dll') do
  begin
    Inc(Index);
    SetLength(libHandles, Index);

    // Load the DLL and store its handle
    libHandles[Index - 1] := LoadLibrary(PWideChar(TPath.GetFileName(fName)));
    if libHandles[Index - 1] = 0 then
    begin
      // Log an error if the DLL fails to load
      with TEventLogger.Create(globalAppName) do
      try
        LogMessage(Format('ERROR: Failed to load library %s from %s', [TPath.GetFileName(fName), Directory]), EVENTLOG_ERROR_TYPE);
      finally
        Free;
      end;
      // Terminate the application if critical DLLs cannot be loaded
      Halt(1);
    end;
  end;
  // Restore the DLL directory to its previous value
  SetDllDirectory(nil);
end;

procedure LoadOpenSSLLibraries;
var
  dllPath: string;
begin
  // Determine the OpenSSL directory based on system architecture
  {$IFDEF WIN32}
    dllPath := GlobalRootPath + 'OpenSSL_32';
  {$ELSE}
    dllPath := GlobalRootPath + 'OpenSSL_64';
  {$ENDIF}
  // Load all OpenSSL DLLs from the determined directory
  LoadLibrariesFromDirectory(dllPath);
end;

initialization
  // Automatically load OpenSSL libraries during unit initialization
  LoadOpenSSLLibraries;

end.

