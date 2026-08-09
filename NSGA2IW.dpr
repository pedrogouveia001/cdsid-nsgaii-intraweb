//program NSGA2IW;
library NSGA2IW;
uses
  ISAPIApp, {ISAPI}
  IWInitISAPI, {ISAPI}
  Forms,
  IWMain,
  Unit2 in 'Unit2.pas' {IWForm2: TIWAppForm},
  ServerController in 'ServerController.pas' {IWServerController: TIWServerControllerBase},
  UserSessionUnit in 'UserSessionUnit.pas' {IWUserSession: TIWUserSessionBase},
  lpsolve51 in 'lpsolve51.pas',
  uAPI_SolveLP in 'uAPI_SolveLP.pas',
  uCode in 'uCode.pas';

{$R *.res}
{Início ISAPI}
exports GetExtensionVersion, HttpExtensionProc, TerminateExtension;
{Fim ISAPI}


begin
  {Application.Initialize;
  Application.CreateForm(TformIWMain, formIWMain);
  Application.Run;}
  IwRun; {ISAPI}
end.
