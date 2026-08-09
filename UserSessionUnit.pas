unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to 
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, SysUtils, Classes, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection, IWCompEdit,
  IWCompLabel, IWGrids, series, IWRegion, IWCompTabControl, IWCompListbox;

type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery1: TZQuery;
    ZConnection1: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
  
  end;

implementation

uses Unit2, ServerController;

{$R *.dfm}

end.