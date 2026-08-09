unit Unit2;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompLabel,
  IWCompText, pngimage, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWExtCtrls, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, jpeg, IWCompButton,
  IWCompListbox, IWCompEdit, IWGrids, ActnList, Comobj, IWWebGrid, IWAdvWebGrid,
  IWAdvWebGridExcel, TeeGDIPlus, TeEngine, ExtCtrls, TeeProcs, Chart, Series,
  IWDBStdCtrls, WebComp, WebAdapt, Math, IdMessage, IdTCPConnection,
  IdTCPClient, Graphics,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP,
  IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdText, IWProgressIndicator,
  IWCompProgressBar, IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWTMSPopup, IWHTMLControls, Menus, IWHelpTip,
  IWCompRadioButton, DateUtils, IWTMSCheckList, Dialogs;
type
  TIWForm2 = class(TIWAppForm)
    rgnprob: TIWRegion;
    lblWarning: TIWLabel;
    btnInputManual: TIWButton;
    btnImportar: TIWButton;
    IWImage9: TIWImage;
    IWLabel70: TIWLabel;
    IWImage10: TIWImage;
    IWImage27: TIWImage;
    IWImage23: TIWImage;
    IWLabel12: TIWLabel;
    IWRegion16: TIWRegion;
    IWRegion3: TIWRegion;
    IWText2: TIWText;
    rgninput: TIWRegion;
    IWImage3: TIWImage;
    IWImage4: TIWImage;
    IWImage5: TIWImage;
    IWLabel7: TIWLabel;
    IWLink2: TIWLink;
    IWRegion4: TIWRegion;
    rgrpPeriodo: TIWRadioGroup;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    IWLabel3: TIWLabel;
    IWLabel4: TIWLabel;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWLabel9: TIWLabel;
    edtprodregular: TIWEdit;
    edthorasextras: TIWEdit;
    edtsubcontrat: TIWEdit;
    edtestoq: TIWEdit;
    edtfunc: TIWEdit;
    IWLabel10: TIWLabel;
    IWLink1: TIWLink;
    rgnadcobjetivo: TIWRegion;
    IWLabel11: TIWLabel;
    IWLabel13: TIWLabel;
    edtnomeobjetivo: TIWEdit;
    IWLink3: TIWLink;
    btnavancaretapa1: TIWButton;
    btnredefiniretapa1: TIWButton;
    btnsalvaretapa1: TIWButton;
    IWLink5: TIWLink;
    IWLink6: TIWLink;
    IWLabel14: TIWLabel;
    IWLabel15: TIWLabel;
    edtinifunc: TIWEdit;
    rgndemanda: TIWRegion;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWImage6: TIWImage;
    IWLabel16: TIWLabel;
    IWLink7: TIWLink;
    IWLink9: TIWLink;
    IWLabel25: TIWLabel;
    IWRegion2: TIWRegion;
    grddemanda: TTIWAdvWebGrid;
    btnsalvaretapa2: TIWButton;
    btnredefiniretapa2: TIWButton;
    imgdemanda: TIWImage;
    btnavancaretapa2: TIWButton;
    IWRegion20: TIWRegion;
    txtOPRSA: TIWText;
    txtCPRSA: TIWText;
    IWImage8: TIWImage;
    IWImage24: TIWImage;
    lblLRSA: TIWLabel;
    IWImage11: TIWImage;
    IWText1: TIWText;
    rgnfuncaoobjetivo: TIWRegion;
    IWImage12: TIWImage;
    IWImage13: TIWImage;
    IWImage14: TIWImage;
    IWLabel17: TIWLabel;
    IWLink8: TIWLink;
    IWLink10: TIWLink;
    IWLabel18: TIWLabel;
    btnsalvaretapa3: TIWButton;
    btnredefiniretapa3: TIWButton;
    IWRegion5: TIWRegion;
    cmbbxObjetivos: TIWComboBox;
    IWLabel19: TIWLabel;
    grdfuncaoobjetivo: TTIWAdvWebGrid;
    btnreiniciaretapa3: TIWButton;
    btnexecutar: TIWButton;
    IWLabel20: TIWLabel;
    chrtDemanda: TChart;
    LineSeries2: TLineSeries;
    LineSeries3: TLineSeries;
    Series1: TLineSeries;
    cbxproducaoregular: TIWCheckBox;
    cbxHorasExtras: TIWCheckBox;
    cbxSubcontratacao: TIWCheckBox;
    cbxEstoque: TIWCheckBox;
    lstbxobjetivos: TIWListbox;
    cmbbxobj: TIWComboBox;
    IWButton1: TIWButton;
    IWButton2: TIWButton;
    lstbxFuncoesObjetivo: TIWListbox;
    cbxcontrat: TIWCheckBox;
    cbxdem: TIWCheckBox;
    IWLabel8: TIWLabel;
    edtnomeproblema: TIWEdit;
    rgnintro: TIWRegion;
    IWLabel21: TIWLabel;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    txtwarnning: TIWText;
    IWButton3: TIWButton;
    IWLabel46: TIWLabel;
    IWImage7: TIWImage;
    IWLabel52: TIWLabel;
    IWImage16: TIWImage;
    IWImage17: TIWImage;
    iwrgnCaixaDialogoRU: TIWRegion;
    lblWg: TIWLabel;
    IWRegionru: TIWRegion;
    txtCaixaDialogoRU: TIWText;
    btnBackRU: TIWButton;
    btncontinueRU: TIWButton;
    IWRegion9: TIWRegion;
    IWRegion6: TIWRegion;
    lblSolve: TIWLabel;
    IWImage15: TIWImage;
    rgnlogin: TIWRegion;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    btnEnter: TIWButton;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    lblwarningpassword: TIWLabel;
    lblwarningemail: TIWLabel;
    IWButton4: TIWButton;
    btnrecover: TIWButton;
    txtpassrecover: TIWText;
    txtnewcode: TIWText;
    IWImage19: TIWImage;
    IWImage33: TIWImage;
    lnkesqueceusenha: TIWLink;
    IWRegion7: TIWRegion;
    IWRegion19: TIWRegion;
    iwrgnCaixaDialogoRU2: TIWRegion;
    lblWg2: TIWLabel;
    IWRegionru2: TIWRegion;
    txtCaixaDialogoRU2: TIWText;
    btnBackRU2: TIWButton;
    btncontinueRU2: TIWButton;
    IWImage18: TIWImage;
    rgnretomada: TIWRegion;
    lblChooseproj: TIWLabel;
    lstbxProjects: TIWListbox;
    btnReset3: TIWButton;
    btnLogoutRgnLogin: TIWButton;
    IWLabel26: TIWLabel;
    lblMexplanation: TIWLabel;
    btnauxchoose: TIWButton;
    IWImage25: TIWImage;
    IWImage34: TIWImage;
    IWRegion1: TIWRegion;
    IWRegion8: TIWRegion;
    IWRegion23: TIWRegion;
    IWImage20: TIWImage;
    IWLink4: TIWLink;
    grdmatrizconseq: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    TIWAdvWebGridExcelIO2: TTIWAdvWebGridExcelIO;
    grdplanos: TTIWAdvWebGrid;
    lnkPlanos: TIWLink;
    procedure rgninputCreate(Sender: TObject);
    procedure rgnprobCreate(Sender: TObject);
    procedure btnInputManualClick(Sender: TObject);
    procedure IWLink2Click(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure edtprodregularAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edthorasextrasAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edtsubcontratAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure edtestoqAsyncKeyPress(Sender: TObject; EventParams: TStringList);
    procedure edtfuncAsyncKeyPress(Sender: TObject; EventParams: TStringList);
    procedure edtinifuncAsyncKeyPress(Sender: TObject;
      EventParams: TStringList);
    procedure rgndemandaCreate(Sender: TObject);
    procedure rgnfuncaoobjetivoCreate(Sender: TObject);
    procedure btnsalvaretapa1Click(Sender: TObject);
    procedure rgrpPeriodoClick(Sender: TObject);
    procedure btnredefiniretapa1Click(Sender: TObject);
    procedure btnavancaretapa1Click(Sender: TObject);
    procedure cbxproducaoregularClick(Sender: TObject);
    procedure cbxHorasExtrasClick(Sender: TObject);
    procedure cbxSubcontratacaoClick(Sender: TObject);
    procedure cbxEstoqueClick(Sender: TObject);
    procedure IWButton1Click(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure btnsalvaretapa2Click(Sender: TObject);
    procedure btnredefiniretapa2Click(Sender: TObject);
    procedure btnavancaretapa2Click(Sender: TObject);
    procedure btnsalvaretapa3Click(Sender: TObject);
    procedure btnredefiniretapa3Click(Sender: TObject);
    procedure btnreiniciaretapa3Click(Sender: TObject);
    procedure btnexecutarClick(Sender: TObject);
    procedure Exportar;
    procedure SalvarBD;
    procedure cbxcontratClick(Sender: TObject);
    procedure cbxdemClick(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure rgnintroCreate(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure btnRegisterClick(Sender: TObject);
    procedure btncontinueRUClick(Sender: TObject);
    procedure btnBackRUClick(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure rgnloginCreate(Sender: TObject);
    procedure lnkesqueceusenhaClick(Sender: TObject);
    procedure btnBackRU2Click(Sender: TObject);
    procedure btncontinueRU2Click(Sender: TObject);
    procedure IWButton4Click(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure rgnretomadaCreate(Sender: TObject);
    procedure IWLink5Click(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure btnauxchooseClick(Sender: TObject);
    procedure IWLink6Click(Sender: TObject);
    procedure lnkplanosClick(Sender: TObject);
  public
  MatrizFinal: Array of Array of String;
  Demanda: Array of Real;
  Limites: Array of Real;
  Custos: Array of Real;
  Periodo : integer;
  ObjSelecionados : array of string;
  ParametrosFuncaoObjetivo: Array of Array of Real;
  Populacao, Population, NovaGeracao: Array of Array of Array of Real;
  // MatrizAvaliacao
  PopulationPerformance, PerformancePopulacao: Array of Array of Real;
  PopulationRanking, OrdemPopulacao: array of array of integer;
  PopulationDistance, DistanciaPopulacional: Array of Real;
  NinicialFunc, Nger, G1, G2: integer;
  DemandaMedia: Real;
  N1, analyst : integer;
  Problema, Id_User : string;
  primeiroacesso : boolean;
  email, Id_Email : string;
  senha, validation: string;
  nameproblem, idproblemasuser, datainicioproblema, datafinalproblema : array of string;
  ordemsel : array of integer;
  auxiliar : integer;
  datainicio : string;
  retomada : boolean;
  end;

implementation

uses UserSessionUnit, ServerController;

{$R *.dfm}


procedure TIWForm2.btnauxchooseClick(Sender: TObject);
var
i,j : integer;
Id_Problema : string;
begin
with UserSession.ZQuery1 do
  begin
   SetLength(Limites, 5);
   Setlength(Custos, 6);

  SQL.Text := 'SELECT * FROM restrictions WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        + '"';
      ExecSQL;
      open;
      Periodo := strtoint(FieldByName('Total_Period').AsString);
      Limites[0] := strtofloat(FieldByName('Regular_Production').AsString);
      Limites[1] := strtofloat(FieldByName('Overtime_Production').AsString);
      Limites[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
      Limites[3] := strtofloat(FieldByName('Stock').AsString);
      Limites[4] := strtofloat(FieldByName('Max_Employees').AsString);
      NinicialFunc := strtoint(FieldByName('Ini_Employees').AsString);
      close;


    SetLength(Demanda, Periodo);
    for i := 0 to Periodo - 1 do
    begin
    SQL.Text := 'SELECT * FROM demand WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        + '" AND Period ="' + inttostr(i+1) + '"';
      ExecSQL;
      open;
      Demanda[i] := strtofloat(FieldByName('Demand').AsString);
      close;
    end;






      SetLength(Custos, 6);


      SQL.Text := 'SELECT * FROM obj_functions WHERE Id_Problem ="' + idproblemasuser[lstbxProjects.ItemIndex]
        +'"';
      ExecSQL;
      open;
      First;
      i := 0;
      j := 0;
      while not eof do
      begin
      SetLength(ObjSelecionados, i+1);
      ObjSelecionados[i] := FieldByName('Objective').AsString;
      if ObjSelecionados[i] = 'Cost' then
      begin
        Custos[0] := strtofloat(FieldByName('Regular_Production').AsString);
        Custos[1] := strtofloat(FieldByName('Overtime_Production').AsString);
        Custos[2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
        Custos[3] := strtofloat(FieldByName('Stock').AsString);
        Custos[4] := strtofloat(FieldByName('Hiring').AsString);
        Custos[5] := strtofloat(FieldByName('Resignation').AsString);
      end
      else
      begin
       SetLength(ParametrosFuncaoObjetivo, j+1, 6);
       ParametrosFuncaoObjetivo[j,0] := strtofloat(FieldByName('Regular_Production').AsString);
       ParametrosFuncaoObjetivo[j,1] := strtofloat(FieldByName('Overtime_Production').AsString);
       ParametrosFuncaoObjetivo[j,2] := strtofloat(FieldByName('Subcontrating_Production').AsString);
       ParametrosFuncaoObjetivo[j,3] := strtofloat(FieldByName('Stock').AsString);
       ParametrosFuncaoObjetivo[j,4] := strtofloat(FieldByName('Hiring').AsString);
       ParametrosFuncaoObjetivo[j,5] := strtofloat(FieldByName('Resignation').AsString);
       inc(j);
      end;
      inc(i);
      Next;
      end;
      Close;

      rgninput.Visible := True;
      rgnretomada.Visible := False;

      if Periodo = 12 then
      rgrpPeriodo.ItemIndex := 0;
      if Periodo = 6 then
      rgrpPeriodo.ItemIndex := 1;
      if Periodo = 4 then
      rgrpPeriodo.ItemIndex := 2;

      edtnomeproblema.Text := nameproblem[lstbxProjects.ItemIndex];

      if (Limites[0] = 0) or (Custos[0] = 0) then
      begin
        cbxproducaoregular.Checked := False;
        cbxproducaoregularClick(Self);
      end;

      if (Limites[1] = 0) or (Custos[1] = 0) then
      begin
        cbxHorasExtras.Checked := False;
        cbxHorasExtrasClick(Self);
      end;

      if (Limites[2] = 0) or (Custos[2] = 0) then
      begin
        cbxSubcontratacao.Checked := False;
        cbxSubcontratacaoClick(Self);
      end;

       if (Limites[3] = 0) or (Custos[3] = 0) then
      begin
        cbxEstoque.Checked := False;
        cbxEstoqueClick(Self);
      end;

      if (Custos[4] = 0) then
      begin
        cbxcontrat.Checked := False;
        cbxcontratClick(Self);
      end;

      if (Custos[5] = 0) then
      begin
        cbxdem.Checked := False;
        cbxdemclick(Self);
      end;

      edtprodregular.Text := FloatToStr(limites[0]);
      edthorasextras.Text := FloatToStr(limites[1]);
      edtsubcontrat.Text := FloatToStr(limites[2]);
      edtestoq.Text := FloatToStr(limites[3]);
      edtfunc.Text := FloatToStr(limites[4]);
      edtinifunc.Text := IntToStr(NinicialFunc);
      btnsalvaretapa1Click(Self);

      lstbxObjetivos.Items.Clear;
      lstbxObjetivos.Items.Add('Selected Objectives:');
      for j := 0 to length(ObjSelecionados) - 1 do
      begin
        lstbxObjetivos.Items.Add(ObjSelecionados[j]);
        for i := 0 to cmbbxObj.Items.Count - 1 do
        begin
          cmbbxObj.ItemIndex := i;
          if cmbbxobj.Text = objSelecionados[j] then
          begin
            cmbbxobj.Items.Delete(cmbbxobj.ItemIndex);
          end;
        end;
      end;
      btnavancaretapa1.Enabled := True;
      datainicio := datetostr(now);

  end;
end;

procedure TIWForm2.btnavancaretapa1Click(Sender: TObject);
var
i : integer;
begin
rgninput.Visible := False;
rgndemanda.Visible := True;
grddemanda.TotalRows := Periodo;
grddemanda.rowcount := Periodo;
Problema := edtnomeproblema.Text;
Chrtdemanda.Series[0].Clear;
Chrtdemanda.Series[1].Clear;
Chrtdemanda.Series[2].Clear;
Chrtdemanda.Series[0].Visible := False;
Chrtdemanda.Series[1].Visible := False;
Chrtdemanda.Series[2].Visible := True;
for i := 0 to Periodo -1  do
begin
chrtDemanda.Series[2].AddXY(i+1, Limites[0]*limites[4] + Limites[1]*limites[4] + Limites[2]);
grddemanda.Cells[0,i] := inttostr(i+1);
end;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));
SetLength(ParametrosFuncaoObjetivo, Length(ObjSelecionados)-1, 6);
cmbbxObjetivos.Items.Clear;
for I := 0 to Length(ObjSelecionados)-1 do
cmbbxobjetivos.Items.Add(objselecionados[i]);
cmbbxobjetivos.ItemIndex := 0;

if retomada = True then
begin
  for i := 0 to Periodo - 1 do
  grddemanda.DynEdits[1,i] := FloatToStr(Demanda[i]);
  btnsalvaretapa2click(self);
end;

end;

procedure TIWForm2.btnavancaretapa2Click(Sender: TObject);
var
i, aux : integer;
begin
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := True;
 SetLength(OrdemSel, cmbbxObjetivos.Items.Count);
if retomada = True then
begin
  for i := 0 to cmbbxObjetivos.Items.Count - 1 do
  begin
    ordemsel[0] := i;
    cmbbxObjetivos.ItemIndex := i;
    if i = 0 then
    begin
      lstbxFuncoesObjetivo.Items.Add('Min. ' + cmbbxObjetivos.Text + ' = ' + floattostr(custos[0]) +
      '*x1 + '+ floattostr(custos[1]) +'*x2 + '+ floattostr(custos[2])+ '*x3 + '+ floattostr(custos[3])+ '*x4 + '
      + floattostr(custos[4])+ '*x5 + ' + floattostr(custos[5])+ '*x6');
    end
    else
    begin
    aux := i;
      lstbxFuncoesObjetivo.Items.Add('Max. ' + cmbbxObjetivos.Text + ' = ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 0]) +
      '*x1 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 1]) +'*x2 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 2])+ '*x3 '
      + floattostr(ParametrosFuncaoObjetivo[aux-1, 3]) +'*x4 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 4]) +'*x5 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 5]) +'*x6');
    end;

 end;
 cmbbxObjetivos.Items.Clear;
 cmbbxobjetivos.Enabled := False;
 grdfuncaoobjetivo.Enabled := False;
 btnsalvaretapa3.Enabled := False;
 btnredefiniretapa3.Enabled := True;
 btnexecutar.Enabled := True;
end;
end;

procedure TIWForm2.btnBackRU2Click(Sender: TObject);
begin
iwrgnCaixaDialogoRU2.visible := false;
end;

procedure TIWForm2.btnBackRUClick(Sender: TObject);
begin
iwrgnCaixaDialogoRU.Visible := false;
end;

procedure TIWForm2.btncontinueRU2Click(Sender: TObject);
begin
WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm2.btncontinueRUClick(Sender: TObject);
begin
WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm2.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  quantregis: integer;

begin
  lblwarningemail.Visible := false;
  lblwarningpassword.Visible := false;
  primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM user WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end;
  if quantregis = 0 then
  begin
    lblwarningemail.Visible := true;
    IWButton1.Visible := true;
  end
  else
  begin
    email := edtemail.Text;
    with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      validation := FieldByName('validacao').AsString;
      close;

      SQL.Text := 'SELECT password FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      senha := FieldByName('password').AsString;
      close;
    end;
    if senha = edtSenha.Text then
    begin
      with UserSession.ZQuery1 do
      begin
        SQL.Text := 'SELECT id FROM user WHERE email="' + (edtemail.Text) + '"';
        ExecSQL;
        open;
        ID_user := FieldByName('id').AsString;
        close;

        SQL.Text := 'SELECT analyst FROM user WHERE email="' + (edtemail.Text)
          + '"';
        ExecSQL;
        open;
        analyst := FieldByName('analyst').AsInteger;
        close;

        if analyst = 0 then
        begin
          email := edtemail.Text;
          senha := edtSenha.Text;
        end;

      end;

      // * Caso o usuário esteja trocando a senha de cadastro *//
      {if validation = 'password' then
      begin
        WebApplication.ShowMessage('Temporary password has been found. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with changing your password.');
        edtSenha.Text := '';
        //Exit;
      end;}

      // * Caso seja o primeiro acesso *//
      if validation <> 'validado' then
      begin
        WebApplication.ShowMessage('User has not been validated. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with user validation.');
        edtSenha.Clear;
        Exit;
      end;

      rgnlogin.Visible := false;
      rgnprob.Visible := true;
    end
    else
    begin
      lblwarningpassword.Visible := true;
      IWButton1.Visible := true;
    end;

  end;

 { // ** Salvando informação de data e horário
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 46;
  SalvarBD;

  // end;
  }
end;

procedure TIWForm2.btnexecutarClick(Sender: TObject);
var
    i, j, k, z1, z2, z: integer;
    finished: boolean;

    function Roleta: Real;
    begin
      // ** Roda obtendo um valor aleatório [0;100] para tomada de decisão **//
      Result := Random * 100;
    end;


    procedure PrimeiraPopulacao;
    var
      str1, str2: string;
      i, j, k, z1, z2, z: integer;
      NFuncAtual, EstoqueAtual, ProducaoTotalPeriodo: Real;
      AuxC, AuxD, AuxE, AuxEstoque, AuxV, AuxEE: Real;
      MaxProd: Array [0 .. 2] of Real;
      flag: boolean;
    begin
      // ** Gera a primeira população aleatoriamente **//
      { ** Cada indivíduo da população é uma matriz estruturada da seguinte forma:
        C1 : Número de ítens em produção regular
        C2: Número de ítens Prod em Horas Extra
        C3: Número de ítens de prod subcontratadas
        C4: Número de prod que ficaram em estoque no período anterior
        C5: Número de contratados no período atual
        C6: Número de demitidos no período atual
        C7: Número de funcionários remanecentes no período atual

        Cada indivíduo é avaliado com respeito às funções objetivo e possui uma linha
        na matriz PopulationPerformance que é estruturada da seguinte forma:

        C1: Performance no objetivo custo
        C2: Performance no objetivo credibilidade
        C3: Performance no objetivo Qualidade
        C4: Performance no objetivo Flexibilidade
        C5: Performance no objetivo rapidez
        C6: Demanda não atendida avaliada a cada período *** }

      // ** Define uma semente aleatória com base na hora **//
      str1 := TimeToStr(Now);
      str2 := '';

      for i := 1 to Length(str1) - 1 do
        if str1[i] <> ':' then
          str2 := str2 + str1[i];

      RandSeed := 54730930; // strToInt(str2);

      // ** Define o tamanho da população como sendo 200 indivíduos **//
      Setlength(Populacao, 500, Periodo, 7);
      Setlength(Population, 500, Periodo, 7);

      // ** Inicialmente os limites máximos de produção são :**//
      NFuncAtual := NinicialFunc;
      EstoqueAtual := 0;
      ProducaoTotalPeriodo := 0;
      MaxProd[0] := DemandaMedia; // **Regular --> Atribui a média da demanda como sendo o máximo
      MaxProd[1] := Limites[1] * Limites[4]; // **HoraExtra
      MaxProd[2] := Limites[2]; // **Subcontratação

      for i := N1 to Length(Populacao) - 1 do
      begin
       NFuncAtual := NinicialFunc;
       EstoqueAtual := 0;
       ProducaoTotalPeriodo := 0;
        // ** Períodos **//
        j := 0;
        while j <= (Length(Populacao[i]) - 1) do
        begin
          flag := false;
          if EstoqueAtual >= Demanda[j] then
            flag := true;

           if demanda[j] > (Limites[0]*limites[4] + Limites[1]*limites[4] + Limites[2] + EstoqueAtual)  then
           begin
             Populacao[i,j,0] := Limites[0]*limites[4];
             Populacao[i,j,1] := Limites[1]*limites[4];
             Populacao[i,j,2] := Limites[2];
             Populacao[i,j,3] := 0;
             Populacao[i,j,4] := limites[4] - NFuncAtual;
             Populacao[i,j,5] := 0;
             Populacao[i,j,6] := limites[4];
             continue;
           end;
          // ** Evita gerar soluções inviáveis na primeira população **//
          while (ProducaoTotalPeriodo + EstoqueAtual < Demanda[j]) or
            (Populacao[i, j, 3] > Limites[3]) or (flag) do
          begin


            ProducaoTotalPeriodo := 0;
            flag := false;
            for k := 0 to Length(Populacao[i, j]) - 1 do
            // ** Gerando plano aleatório para o período **//
            begin
              // ** Produção **//
              if k <= 2 then
              begin
                if k <> 1 then
                  AuxE := MaxProd[k]
                else
                begin
                if limites[0] <> 0 then
                  AuxV := (Populacao[i, j, 0] / Limites[0]);
                  if AuxV < NFuncAtual then
                    AuxV := NFuncAtual;

                  // ** a máxima produção em HE vai depender da quantidade de pessoal em prod regular **//
                  // ** Ninguém contrata pessoas apenas para trabalhar em hora extra **//
                  AuxE := AuxV * Limites[1];
                end;

                Populacao[i, j, k] := Round(Random * AuxE);

                // ** Introduz uma "Mutação" das soluções em 25% dos casos **//
                if k = 0 then
                begin
                  // ** Força a produção da máxima capacidade regular **//
                  if Roleta < 25 then
                    Populacao[i, j, k] := NFuncAtual * Limites[0];

                  // ** Usar ao menos 60% da capacidade do sistema **//
                  if Populacao[i, j, k] < Limites[0] * NFuncAtual * 0.6 then
                    Populacao[i, j, k] := Round(Limites[0] * NFuncAtual * 0.6);
                end;

                // ** Contabiliza a produção **//
                ProducaoTotalPeriodo := ProducaoTotalPeriodo + Populacao
                  [i, j, k];

                // ** Continua para a próxima iteração **//
                continue;
              end;

              // **Estoque **//
              if k = 3 then
              begin
                AuxC := (ProducaoTotalPeriodo + EstoqueAtual) - Demanda[j];
                // ** A produção não atendeu a demada **//
                if AuxC < 0 then
                  Break;
                // ** Se já for o último período, não devemos ter estoque remanescente **//
                if j = Length(Populacao[i]) - 1 then
                  if AuxC <> 0 then
                  begin

                    // ** força uma "mutação" da solução **//
                    // ** Se o estoque já suprir a demanda nada será produzido **//
                    if EstoqueAtual >= Demanda[j] then
                    begin
                      Populacao[i, j, 0] := 0;
                      Populacao[i, j, 1] := 0;
                      Populacao[i, j, 2] := 0;
                      ProducaoTotalPeriodo := 0;
                      AuxC := EstoqueAtual - Demanda[j];

                    end
                    else if AuxC > 0 then
                    begin
                      ProducaoTotalPeriodo := 0;
                      // ** Caso contrário, deveremos produzir só o necessário para atender a demanda **//
                      // **Melhor dos casos **//
                      AuxD := Demanda[j] - EstoqueAtual;
                      AuxE := AuxD;
                      for z := 0 to 3 - 1 do
                      begin
                        // ** Se a Qtd a ser produzida já foi obtida, a produção nesta modalidade é nula **//
                        if AuxE = 0 then
                        begin
                          Populacao[i, j, z] := 0;
                          continue;
                        end;

                        // ** Verificando o valor a ser multiplicado pelo número aleatório **//
                        if MaxProd[z] > AuxE then
                          AuxEstoque := AuxE
                        else
                          AuxEstoque := MaxProd[z];

                        Populacao[i, j, z] := Round(Random * AuxEstoque);

                        // ** Desconta a quantidade já produzida **//
                        AuxE := AuxE - Populacao[i, j, z];

                        ProducaoTotalPeriodo := ProducaoTotalPeriodo + Populacao
                          [i, j, z];
                      end; // ZZ
                      // ** Caso ainda haja demanda **//
                      if AuxE > 0 then
                      begin
                        if Populacao[i, j, 0] + AuxE < NFuncAtual * Limites[0]
                          then
                          Populacao[i, j, 0] := Populacao[i, j, 0] + AuxE
                        else if Populacao[i, j, 1] + AuxE < NFuncAtual * Limites
                          [1] then
                          Populacao[i, j, 1] := Populacao[i, j, 1] + AuxE
                        else if Populacao[i, j, 2] + AuxE < NFuncAtual * Limites
                          [3] then
                          Populacao[i, j, 2] := Populacao[i, j, 2] + AuxE;
                        ProducaoTotalPeriodo := ProducaoTotalPeriodo + AuxE;
                      end;

                      AuxC := (ProducaoTotalPeriodo + EstoqueAtual)
                        - Demanda[j];
                    end; // Aux>0
                  end; // Aux <> 0
              end;

              // **Contratações**//
              if k = 4 then
              begin
                // ** Identifica a maior necessidade de funcionários **//
                if Populacao[i, j, 0] / Limites[0] >= Populacao[i, j, 1]
                  / Limites[1] then
                  AuxD := Populacao[i, j, 0] / Limites[0]
                else
                  AuxD := Populacao[i, j, 1] / Limites[1];

                // ** Armazena a informação para a próxima iteração **//
                AuxE := AuxD;
                AuxC := 0;
                // ** Contabiliza a necessidade de contratação **//
                while (AuxD > NFuncAtual) do
                begin
                  AuxD := AuxD - 1;
                  AuxC := AuxC + 1;
                end;
              end;

              // ** Demissões **//
              if k = 5 then
              begin
                // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
                AuxC := -1;
                // ** Contabiliza a necessidade de contratação **//
                while (AuxE <= NFuncAtual) do
                begin
                  AuxE := AuxE + 1;
                  AuxC := AuxC + 1;
                end;

                if AuxC < 0 then
                  AuxC := 0;

                // ** Se houver excesso de Mão de Obra **//
                if AuxC > 0 then
                begin
                  // ** Define aleatoriamente se haverá demissão **//
                  if (Roleta > 50) and (j <> Length(Populacao[i]) - 1) then
                    // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
                    AuxD := AuxC
                  else
                    AuxD := 0;

                  AuxC := AuxD
                end;
              end;

              // ** Total de funcionários remanescentes **//
              if k = 6 then
                AuxC := NFuncAtual + Populacao[i, j, 4] - Populacao[i, j, 5];

              Populacao[i, j, k] := AuxC;
            end; // ** k
          end; // **With

          // ** Caso a solução tenha sido adequada para o período, atualiza as var **//
          NFuncAtual := Populacao[i, j, 6];
          EstoqueAtual := Populacao[i, j, 3];
          ProducaoTotalPeriodo := 0;

          if (j = Length(Populacao[i]) - 1) and (EstoqueAtual <> 0) then
          begin
            dec(j);
            NFuncAtual := Populacao[i, j-1, 6];
            EstoqueAtual := Populacao[i, j-1, 3];
            ProducaoTotalPeriodo := 0;
          end;
          inc(j)
        end; // ** j

      end; // i

      // ** Passando a populacao para o vetor que armazena o conjunto total de indivíduos a serem avaliados **//
      for i := 0 to Length(Populacao) - 1 do
        for j := 0 to Length(Populacao[i]) - 1 do
          for k := 0 to Length(Populacao[i, j]) - 1 do
            Population[i, j, k] := Populacao[i, j, k];

    end;

    procedure avaliarPopulacao;
    var
      i, j, k, z: integer;
      AuxCalc, AuxV, Prod: Real;
    begin

      Setlength(PopulationPerformance, Length(Population), length(objSelecionados)+1);

      for i := 0 to Length(Population) - 1 do
      begin
        PopulationPerformance[i, length(objSelecionados)] := 0;
        AuxV := 0;
        for j := 0 to length(objSelecionados)-1 do
        begin
          PopulationPerformance[i, j] := 0;
          AuxCalc := 0;

          for z := 0 to Length(Population[i]) - 1 do
          begin
            Prod := 0;
            // ** Custo **//
            if j = 0 then
            begin
              for k := 0 to Length(Population[i, z]) - 2 do
              begin
                PopulationPerformance[i, j] := PopulationPerformance[i, j]
                  + Population[i, z, k] * Custos[k];

                // ** Armazena a produção no período **//
                if k < 3 then
                  Prod := Prod + Population[i, z, k];

                // ** Avaliando atendimento a restrição de capacidade **//
                if k < 2 then
                  if Population[i, z, k] > Limites[k] * 10 then
                    AuxV := AuxV + Population[i, z, k] - Limites[k] * 10;
                if k = 2 then
                  if Population[i, z, k] > Limites[k] then
                    AuxV := AuxV + Population[i, z, k] - Limites[k];
              end;
              if z > 0 then
                // ** Adiciona o estoque remanescente no período anterior **//
                Prod := Prod + Population[i, z - 1, 3];
              AuxCalc := 1;
              // ** Avaliando o atendimento às restrições - Demanda **//
              if Demanda[z] > Prod then
                AuxV := AuxV + Demanda[z] - Prod;
              // ** Avaliando o atendo=imento a restrição de limite de estoque **//
              if Population[i, z, 3] > Limites[3] then
                AuxV := AuxV + Population[i, z, 3] - Limites[3];
            end
            else
            begin
              for k := 0 to 5 do
              begin
                PopulationPerformance[i, j] := PopulationPerformance[i, j]
                  + Population[i, z, k] * ParametrosFuncaoObjetivo[j - 1, k];
                if ParametrosFuncaoObjetivo[j - 1, k] <> 0 then
                AuxCalc := AuxCalc + Population[i, z, k];
              end;

            end; // ** else
          end; // ** z

          if AuxCalc = 0 then
            AuxCalc := 1;

          // ** Valor médio do objetivo **//
          PopulationPerformance[i, j] := PopulationPerformance[i, j] / AuxCalc;
        end; // ** J
        // ** Violação total da restrição **//
        PopulationPerformance[i, length(objSelecionados)] := AuxV;
      end;

    end;

    procedure Crossover(X1, X2, z1, z2: integer);
    var
      i, j, k: integer;
      Aleatorio, Beta, Nc, AuxC, AuxD, AuxE1, AuxE2: Real;
      Producao, FunCAtual, Demissao: Array [0 .. 1] of Real;
    begin
      Nc := 2;

      Aleatorio := Random;

      if Aleatorio <= 50 then
        Beta := 2 * Power((Aleatorio), (1 / (Nc + 1)))
      else
        Beta := (1 / (2 * Power((1 - Aleatorio), (1 / (Nc + 1)))));

      // ** Funcionários no início **//
      FunCAtual[0] := NinicialFunc;
      FunCAtual[1] := NinicialFunc;

      for i := 0 to Length(NovaGeracao[0]) - 1 do
      begin
        Producao[0] := 0;
        Producao[1] := 0;

        // ** Valores da produ~ção **//
        for j := 0 to 3 - 1 do
        begin

          // ** Filho 1 **//
          NovaGeracao[z1, i, j] := ABS
            (Round(0.5 * ((1 + Beta) * Populacao[X1, i, j] + (1 - Beta)
                  * Populacao[X2, i, j])));
          // ** Filho 22 **//
          NovaGeracao[z2, i, j] := ABS
            (Round(0.5 * ((1 - Beta) * Populacao[X1, i, j] + (1 + Beta)
                  * Populacao[X2, i, j])));

          // ** Coletando dados **//
          Producao[0] := Producao[0] + NovaGeracao[z1, i, j];
          Producao[1] := Producao[1] + NovaGeracao[z2, i, j];
        end;

        // ** Calculando o estoque do período **//
        if i > 0 then
          AuxC := NovaGeracao[z1, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[0] < 0) then
          NovaGeracao[z1, i, 3] := AuxC + Producao[0] - Demanda[i]
        else
          NovaGeracao[z1, i, 3] := 0;

        if i > 0 then
          AuxC := NovaGeracao[z2, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[1] < 0) then
          NovaGeracao[z2, i, 3] := AuxC + Producao[1] - Demanda[i]
        else
          NovaGeracao[z2, i, 3] := 0;
        /// ////////////////////////////////////////////////////

        // ** Calculando contratações **//
        if NovaGeracao[z1, i, 0] / Limites[0] >= NovaGeracao[z1, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z1, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z1, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE1 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[0]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z1, i, 4] := AuxC;
        FunCAtual[0] := FunCAtual[0] + AuxC;

        // ** Calculando contratações **//
        if NovaGeracao[z2, i, 0] / Limites[0] >= NovaGeracao[z2, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z2, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z2, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE2 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[1]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z2, i, 4] := AuxC;
        FunCAtual[1] := FunCAtual[1] + AuxC;
        /// ////////////////////////////////////////////////////////////////////

        // ** Demissões **//
        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE1 <= FunCAtual[0]) do
        begin
          AuxE1 := AuxE1 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z1, i, 5] := AuxC;
        FunCAtual[0] := FunCAtual[0] - AuxC;

        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE2 <= FunCAtual[0]) do
        begin
          AuxE2 := AuxE2 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z2, i, 5] := AuxC;
        FunCAtual[1] := FunCAtual[1] - AuxC;

        // ** Total de funcionários remanescentes **//
        NovaGeracao[z1, i, 6] := FunCAtual[0];
        NovaGeracao[z2, i, 6] := FunCAtual[1];

      end; // ** i

    end;

    procedure Mutation(X1, X2, z1, z2: integer);
    var
      i, j, k: integer;
      Aleatorio, Beta, Nc, AuxC, AuxD, AuxE1, AuxE2: Real;
      Producao, FunCAtual, Demissao: Array [0 .. 1] of Real;
    begin

      // ** Funcionários no início **//
      FunCAtual[0] := NinicialFunc;
      FunCAtual[1] := NinicialFunc;

      for i := 0 to Length(NovaGeracao[0]) - 1 do
      begin
        Producao[0] := 0;
        Producao[1] := 0;

        // ** Valores da produ~ção **//
        for j := 0 to 3 - 1 do
        begin
          // ** Mutação 1 **//
          if Roleta > 50 then
            Aleatorio := Random
          else
            Aleatorio := 1;

          NovaGeracao[z1, i, j] := Round
            (Populacao[X1, i, j] * 0.5 + Populacao[X1, i, j] * Aleatorio);

          // ** Mutação 2 **//
          NovaGeracao[z2, i, j] := Round
            (Populacao[X2, i, j] * 0.5 + Aleatorio * Populacao[X2, i, j]);

          // ** Coletando dados **//
          Producao[0] := Producao[0] + NovaGeracao[z1, i, j];
          Producao[1] := Producao[1] + NovaGeracao[z2, i, j];
        end;

        // ** Calculando o estoque do período **//
        if i > 0 then
          AuxC := NovaGeracao[z1, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[0] < 0) then
          NovaGeracao[z1, i, 3] := AuxC + Producao[0] - Demanda[i]
        else
          NovaGeracao[z1, i, 3] := 0;

        if i > 0 then
          AuxC := NovaGeracao[z2, i - 1, 3]
        else
          AuxC := 0;

        if (Demanda[i] - AuxC - Producao[1] < 0) then
          NovaGeracao[z2, i, 3] := AuxC + Producao[1] - Demanda[i]
        else
          NovaGeracao[z2, i, 3] := 0;
        /// ////////////////////////////////////////////////////

        // ** Calculando contratações **//
        if NovaGeracao[z1, i, 0] / Limites[0] >= NovaGeracao[z1, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z1, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z1, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE1 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[0]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z1, i, 4] := AuxC;
        FunCAtual[0] := FunCAtual[0] + AuxC;

        // ** Calculando contratações **//
        if NovaGeracao[z2, i, 0] / Limites[0] >= NovaGeracao[z2, i, 1] / Limites
          [1] then
          AuxD := NovaGeracao[z2, i, 0] / Limites[0]
        else
          AuxD := NovaGeracao[z2, i, 1] / Limites[1];

        // ** Armazena a informação para a próxima iteração **//
        AuxE2 := AuxD;
        AuxC := 0;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxD > FunCAtual[1]) do
        begin
          AuxD := AuxD - 1;
          AuxC := AuxC + 1;
        end;
        NovaGeracao[z2, i, 4] := AuxC;
        FunCAtual[1] := FunCAtual[1] + AuxC;
        /// ////////////////////////////////////////////////////////////////////

        // ** Demissões **//
        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE1 <= FunCAtual[0]) do
        begin
          AuxE1 := AuxE1 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z1, i, 5] := AuxC;
        FunCAtual[0] := FunCAtual[0] - AuxC;

        // ** Iniciando em -1 garantimos que mesmo após a demissão ainda haverá MO suficiente **//
        AuxC := -1;
        // ** Contabiliza a necessidade de contratação **//
        while (AuxE2 <= FunCAtual[0]) do
        begin
          AuxE2 := AuxE2 + 1;
          AuxC := AuxC + 1;
        end;

        if AuxC < 0 then
          AuxC := 0;

        // ** Se houver excesso de Mão de Obra **//
        if AuxC > 0 then
        begin
          // ** Define aleatoriamente se haverá demissão **//
          if (Roleta > 50) and (j <> Length(NovaGeracao[i]) - 1) then
            // ** Caso vá haver demissão, seleciona aleatoriamente a qtd de funcionários demitidos **//
            AuxD := AuxC
          else
            AuxD := 0;

          AuxC := AuxD
        end;
        NovaGeracao[z2, i, 5] := AuxC;
        FunCAtual[1] := FunCAtual[1] - AuxC;

        // ** Total de funcionários remanescentes **//
        NovaGeracao[z1, i, 6] := FunCAtual[0];
        NovaGeracao[z2, i, 6] := FunCAtual[1];

      end; // ** i

    end;

    function RankingSorting: boolean;
    var
      i, j, k, z, n: integer;
      Ndominancias: Array of integer; // ** Número de projetos que domina cada projeto
      NSup: Array [0 .. 1] of integer;
    begin

      Setlength(PopulationRanking, 0, 0);
      Setlength(PopulationRanking, 1);
      Setlength(Ndominancias, Length(Population));

      // ** Inicializa a variável **//
      for i := 0 to Length(Population) - 1 do
      begin
        Ndominancias[i] := 0;
      end;

      // ** Comparação Par a Par **//
      for i := 0 to Length(Population) - 1 do
      Begin
        for j := i + 1 to Length(Population) - 1 do
        begin
          // ** Regras de dominância restrita **//
          if (PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              < PopulationPerformance[j, Length(PopulationPerformance[i]) - 1])
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[j]);
            continue;
          end;
          if (PopulationPerformance[i, Length(PopulationPerformance[i]) - 1]
              > PopulationPerformance[j, Length(PopulationPerformance[i]) - 1])
            then
          begin
            { ** Se I ferir menos as restrições do que J, já se tem uma dominância
              Este caso já inclui uma ser viável e a outra não, além de as duas serem inviáveis
              porém uma melhor que a outra ** }
            Inc(Ndominancias[i]);
            continue;
          end;
          // ** Inicializa contadores **//
          NSup[0] := 0;
          NSup[1] := 0;

          // ** Caso uma dominancia não tenha sido estabelecida ainda, vai para os payoffs **//
          for k := 0 to Length(PopulationPerformance[i]) - 2 do
            // ** Custo - Minimização **
            if k = 0 then
            begin
              if PopulationPerformance[i, k] > PopulationPerformance[j, k] then
                Inc(NSup[1]);
              if PopulationPerformance[i, k] < PopulationPerformance[j, k] then
                Inc(NSup[0]);
            end
            else
            begin
              // ** Obj de maximização **//
              if PopulationPerformance[i, k] > PopulationPerformance[j, k] then
                Inc(NSup[0]);
              if PopulationPerformance[i, k] < PopulationPerformance[j, k] then
                Inc(NSup[1]);
            end;

          // ** I Domina J **//
          if (NSup[0] > 0) and (NSup[1] = 0) then
            Inc(Ndominancias[j]);

          // ** J Domina I **//
          if (NSup[1] > 0) and (NSup[0] = 0) then
            Inc(Ndominancias[i]);

          // ** Caso nenhuma das regras tenham sido atendidas, os projetos são incomparáveis **//
        end;
        // ** Caso I não seja dominado por ninguém, ele pertence ao R1 **//
        if Ndominancias[i] = 0 then
        begin
          Setlength(PopulationRanking[0], Length(PopulationRanking[0]) + 1);
          // ** inclui na primeira posição **//
          PopulationRanking[0, Length(PopulationRanking[0]) - 1] := i;
        end;
      End;
      z := 1;

      n := 0;
      while n < Length(Population) do
      begin
        n := 0;
        for j := 0 to Length(Population) - 1 do
        begin
          // ** Subtrai 1 no número de Dominâncias para a construção do ranking**//
          Dec(Ndominancias[j]);
          if Ndominancias[j] <= 0 then
            Inc(n);
        end;
        z := 0;

        for i := 0 to Length(Population) - 1 do
        begin
          // ** Se após reduzir em um, o num de dom se tornar zero, a sol pertence ao atual nível do Ranking **//
          if Ndominancias[i] = 0 then
          begin
            Inc(z);

            // ** Cria o nível atual do ranking **//
            if z = 1 then
              Setlength(PopulationRanking, Length(PopulationRanking) + 1);

            // ** Adiciona mais um slot ao nível atual **//
            Setlength(PopulationRanking[Length(PopulationRanking) - 1], Length
                (PopulationRanking[Length(PopulationRanking) - 1]) + 1);
            // ** Adiciona a solução ao atual nível **//
            PopulationRanking[Length(PopulationRanking) - 1, Length
              (PopulationRanking[Length(PopulationRanking) - 1]) - 1] := i;
          end;
        end; // **i

      end; // ** while z

      // ** Se todas os indivíduos forem não dominados **//
      if Length(PopulationRanking) = 1 then
        // ** Avisa que chegamos ao final do processo **//
        Result := true
      else
        Result := false;

    end;

    procedure CrowdingDistance;
    var
      i, j, Katual, k, n: integer;
      OrderM: Array of integer;
      ConsA, ConsB, Mg, ConsRange: Real;
      Desempenho: array of Real;
      Ndominancias: Array of integer; // ** Número de projetos que domina cada projeto
    begin

      Setlength(PopulationDistance, Length(Population));
      Setlength(Desempenho, Length(Population));
      Setlength(OrderM, Length(Population));

      // ** Inicializando **//
      for i := 0 to Length(Population) - 1 do
        PopulationDistance[i] := 0;

      // ** Para cada objetivo **//
      for j := 0 to Length(PopulationPerformance[0]) - 2 do
      begin

        // ** Passa o desempenho para o vet auxiliar **//
        for i := 0 to Length(PopulationPerformance) - 1 do
          Desempenho[i] := PopulationPerformance[i, j];

        Setlength(Ndominancias, Length(Population));

        // ** Inicializa a variável **//
        for i := 0 to Length(Population) - 1 do
        begin
          Ndominancias[i] := 0;
          OrderM[i] := -1;
        end; // **i

        // ** Ordena os indivíduos do menor p/ o maior**//
        for i := 0 to Length(Population) - 1 do
        begin
          for k := i to Length(Population) - 1 do
          begin
            if Desempenho[i] < Desempenho[k] then
              Inc(Ndominancias[k]);
            if Desempenho[i] > Desempenho[k] then
              Inc(Ndominancias[i]);
          end; // **k

        end; // **i

        Katual := 0;
        n := 0;
        ConsA := -1;
        ConsB := -1;

        // ** Ordena com base no número de dominâncias recebidas **//
        while OrderM[Length(Population) - 1] = -1 do
        begin
          for k := 0 to Length(Population) - 1 do
            // ** Já não existem indivíduos com desempenho menor que K;
            if Ndominancias[k] = n then
            begin
              OrderM[Katual] := k;
              Inc(Katual);
              // ** Seleciona o menor e a melhor performance **//
              if n = 0 then
                ConsA := Desempenho[k];
              if Katual = Length(Population) then
                ConsB := Desempenho[k];
            end; // **Ndom

          if ConsB = -1 then
            if OrderM[Katual] = -1 then
              Inc(n);
        end; // **While

        if ConsA = ConsB then
        begin
          Mg := 0;
          ConsRange := 1
        end
        else
        begin
          Mg := Power(10, 5);
          ConsRange := ConsB - ConsA;
        end;

        // ** Define o Crowding Distance **//
        for i := 1 to Length(Population) - 2 do
        begin
          // ** O primeiro e o último indivíduo da população **//
          if (i = 1) then
            PopulationDistance[OrderM[i - 1]] := PopulationDistance
              [OrderM[i - 1]] + Mg;
          if (i = Length(Population) - 2) then
            PopulationDistance[OrderM[i + 1]] := PopulationDistance
              [OrderM[i + 1]] + Mg;

          PopulationDistance[OrderM[i]] := PopulationDistance[OrderM[i]] +
            (Desempenho[OrderM[i + 1]] - Desempenho[OrderM[i - 1]]) / ConsRange;

          if PopulationDistance[OrderM[i]] < 0 then
            PopulationDistance[OrderM[i]] := ABS(PopulationDistance[OrderM[i]]);

        end;

      end; // ** j

    end;

    procedure SelecaoNatural;
    var
      i, j, Isel, k, z, ICandidato: integer;
      UltimoNivel: boolean;
    begin
      Isel := 0;

      Setlength(PerformancePopulacao, Length(Populacao), Length
          (PopulationPerformance[0]));
      Setlength(DistanciaPopulacional, Length(Populacao));
      Setlength(OrdemPopulacao, 0, 0);

      // ** Percorre os níveis do ranking **//
      for i := 0 to Length(PopulationRanking) - 1 do
      begin
        // ** Acrescenta um nível no ranking final **//
        Setlength(OrdemPopulacao, i + 1);
        if Isel + Length(PopulationRanking[i]) > Length(Populacao) then
          UltimoNivel := true
        else
          UltimoNivel := false;

        for j := 0 to Length(PopulationRanking[i]) - 1 do
        begin
          ICandidato := PopulationRanking[i, j];
          // ** Todo o nível não cabe na população **//
          if UltimoNivel then
          begin
            for k := 0 to Length(PopulationRanking[i]) - 1 do
              if PopulationDistance[ICandidato] < PopulationDistance
                [PopulationRanking[i, k]] then
                // ** Seleciona a maior distância Populacional **//
                ICandidato := PopulationRanking[i, k];
          end;

          // ** Armazena a solução **//
          for k := 0 to Length(Population[ICandidato]) - 1 do
            for z := 0 to Length(Population[ICandidato, k]) - 1 do
              Populacao[Isel, k, z] := Population[ICandidato, k, z];

          // ** Armazena a performance **//
          for k := 0 to Length(PerformancePopulacao[0]) - 1 do
            PerformancePopulacao[Isel, k] := PopulationPerformance
              [ICandidato, k];

          // ** Armazena a Crowding Distance **//
          DistanciaPopulacional[Isel] := PopulationDistance[ICandidato];

          // ** Acrescenta um slot no nível atual **//
          Setlength(OrdemPopulacao[i], Length(OrdemPopulacao[i]) + 1);

          // ** Amazena o Ranking remanescente **//
          OrdemPopulacao[i, Length(OrdemPopulacao[i]) - 1] := Isel;

          // ** A partir desse momento, altera o Crowding D para evitar que se repita alguma solução**//
          if UltimoNivel then
            PopulationDistance[ICandidato] := -100;

          Inc(Isel);

          // ** Finaliza o processo **//
          if Isel = Length(Populacao) then
            exit;
        end; // **j

      end; // **i

    end;

    Procedure SelecaoGenitores;
    var
      Rposition, i, j: integer;
      X: array [0 .. 1] of integer;
      RDAtual: Real;
    begin
      G1 := 0;
      G2 := 0;

      // ** Os genitores devem ser distintos um do outro **//
      while G1 = G2 do
      begin
        for i := 0 to 2 - 1 do
        begin
          // ** Seleciona a posição do ranking aleatoriamente **//
          Rposition := Round((Length(OrdemPopulacao) - 1) * Random);

          // ** Como é bastante difícil [dependendo do num de niveis] selecionar o primeiro e o último níveis, damos uma mutada **//
          if Roleta < 30 then
            // ** +30% de chance de selecionar a primeira posição **//
            Rposition := 0
          else if Roleta > 95 then
            // ** +5% de chance de selecionar a ultima posição **//
            Rposition := Length(OrdemPopulacao) - 1;

          RDAtual := -8000;

          // **Seleciona aleatoriamente em 10% das vezes **//
          if Roleta < 10 then
            X[i] := Round((Length(OrdemPopulacao[Rposition]) - 1) * Random)
          else
            // ** Busca o indivíduo com maior crowding distance no ranking **//
            for j := 0 to Length(OrdemPopulacao[Rposition]) - 1 do
              if RDAtual < DistanciaPopulacional[OrdemPopulacao[Rposition, j]]
                then
              begin
                RDAtual := DistanciaPopulacional[OrdemPopulacao[Rposition, j]];
                X[i] := OrdemPopulacao[Rposition, j];
              end;
        end;
        /// ** i

        G1 := X[0];
        G2 := X[1];
      end; // **While

    end;

  begin
    // ** Gera a primeira geração **//
    PrimeiraPopulacao;

    finished := false;

    // ** Realiza o processo evolutivo **//
    while not finished do
    begin
      // ** Avalia o desempenho da população **//
      avaliarPopulacao;
      // ** Ranqueia as soluções **//
      finished := RankingSorting;
      // **Calcula o Crowding distance **//
      CrowdingDistance;
      // ** Realiza a seleção dos indivíduos que sobreviverão **//
      SelecaoNatural;

      // **Gerar uma nova populacao **//
      Setlength(NovaGeracao, Length(Populacao), Length(Populacao[0]), Length
          (Populacao[0, 0]));

      z1 := 0;
      z2 := 1;

      // ** Criar uma nova população de tamanho padrão **//
      While z2 < Length(Populacao) do
      begin
        // ** Seleciona genitores **//
        SelecaoGenitores;

        // ** 70% de chance de crossover, 30% de mutação **//
        if Roleta < 70 then
        begin
          Crossover(G1, G2, z1, z2);

          Inc(z1);
          Inc(z2);
        end
        else
        begin
          Mutation(G1, G2, z1, z2);

          Inc(z1);
          Inc(z2);
        end;
      end; // ** z

      // **Junta a nova população com a anterior **//
      Setlength(Population, 2 * Length(Populacao), Length(Populacao[0]), Length
          (Populacao[0, 0]));

      for i := 0 to 2 * Length(Populacao) - 1 do
        for j := 0 to Length(Populacao[0]) - 1 do
          for k := 0 to Length(Populacao[0, 0]) - 1 do
          begin
            if i < Length(Populacao) then
              Population[i, j, k] := Populacao[i, j, k]
            else
              Population[i, j, k] := NovaGeracao[i - Length(Populacao), j, k];
          end;

      Inc(Nger);

      if Nger = 5000 then
        finished := true;

    end; // ** Finished
  Exportar;
  lnkplanos.Visible := True;
  SalvarBD;
  WebApplication.ShowMessage('The problem has been solved successfully!');

  end;

procedure TIWForm2.Exportar;
var
  ExcelApplication: variant;
  Sheet1, Sheet2, Sheet3: variant;
  i, j, z, k, t, counter1, counter2: integer;
  Matriz: Array of Array of Real;
begin

  counter2 := -1;
  grdmatrizconseq.TotalRows := 9 + Length(Populacao);
  grdplanos.TotalRows := 3;
  for I := 0 to Length(ObjSelecionados) - 1 do
  begin
  grdmatrizconseq.columns.Add;
  grdmatrizconseq.Cells[i + 1,0] := ObjSelecionados[i];
  if i > 0 then
  grdmatrizconseq.Cells[i + 1,1] := '1'
  else
  grdmatrizconseq.Cells[i + 1,1] := '0'
  end;
  grdmatrizconseq.Cells[0,0] := 'Criteria:';
  grdmatrizconseq.Cells[0,1] := 'Preference Direction:';
  grdmatrizconseq.Cells[0,6] := 'Number of Levels:';
  grdmatrizconseq.Cells[0,7] := 'Alternatives:';
  grdmatrizconseq.Cells[1,7] := 'Consequence Matrix:';

  for I := 0 to 8 do
  begin
  grdplanos.columns.Add;
  end;
  //** Adicionar apenas alternativas não repetidas **//
  for i := 0 to Length(Populacao) - 1 do
  begin
  grdmatrizconseq.Cells[0,8+i] := 'Plan ' + inttostr(i+1);
    for t := 0 to Length(Matriz) - 1 do
    begin
    counter1 := 0;
      for j := 0 to Length(PerformancePopulacao[i]) - 1 do

        if PerformancePopulacao[i, j] = Matriz[t, j] then
          Inc(counter1);
    end;

    if counter1 = Length(PerformancePopulacao[i]) then
      continue
    else
    begin
      Setlength(Matriz, Length(Matriz) + 1, Length(PerformancePopulacao[i])-1);
      for j := 0 to Length(PerformancePopulacao[i]) - 2 do
      begin
        Matriz[Length(Matriz) - 1, j] := PerformancePopulacao[i, j];
        grdmatrizconseq.Cells[j + 1,Length(Matriz)+7] :=  floatToStr(PerformancePopulacao[i, j]);

      end;



    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
   grdplanos.Cells[0,Counter2+1] := 'Plan ' + inttostr(i+1) + ':';
    for k := 0 to Length(Populacao[i]) - 1 do
    begin
    Inc(Counter2);
    grdplanos.TotalRows := 2+ Counter2;
    //grdplanos.Cells[0,Counter2+1] := inttostr(z+1);
      for z := 0 to Length(Populacao[i, k]) - 1 do
       //
       grdplanos.Cells[Z + 1,Counter2+1] := floatToStr(Populacao[i, k, z]);
    end;
      end; //** else
  end;
   grdplanos.Cells[0,0] := 'Plan';
   grdplanos.Cells[1,0] := 'Regular Production';
   grdplanos.Cells[2,0] := 'Overtime Production';
   grdplanos.Cells[3,0] := 'Subcontrating Production';
   grdplanos.Cells[4,0] := 'Stock';
   grdplanos.Cells[5,0] := 'Hiring';
   grdplanos.Cells[6,0] := 'Resignation';
   grdplanos.Cells[7,0] := 'Total Employees';



  TIWAdvWebGridExcelIO1.AdvWebGrid:= grdmatrizconseq;
 TIWAdvWebGridExcelIO1.XLSExport('Consequence_Matrix_' + problema + '.xls', WebApplication);
 Inherited;

  end;

Procedure  TIWForm2.SalvarBD;
var
i, j : integer;
Id_Problema, prefdir : string;
begin
  with UserSession.ZQuery1 do
  begin
  sql.Text := 'INSERT INTO problem (User, Name, Creation_Date, Conclusion_Date) VALUES ("' + Id_User + '", "' + Problema + '", "' + datainicio + '", "' + datetostr(now) + '")';
  ExecSQL;

  SQL.Text := 'SELECT Id_Problem FROM problem WHERE User ="' + (Id_User)
        + '"';
      ExecSQL;
      open;
      Last;
      Id_Problema := FieldByName('Id_Problem').AsString;
      close;

  sql.Text := 'INSERT INTO restrictions (Id_Problem, Total_Period, Regular_Production, Overtime_Production, Subcontrating_Production, Stock, Ini_Employees, Max_Employees) VALUES ("' + Id_Problema + '" ,"' + inttostr(Periodo) + '","' + floattostr(limites[0]) + '", "' + floattostr(limites[1]) + '", "' + floattostr(limites[2]) + '", "' + floattostr(limites[3]) + '", "' + inttostr(NInicialFunc) + '", "' + floattostr(limites[4]) + '" )';
  ExecSQL;

  for i := 0 to Periodo - 1 do
  begin
   sql.Text := 'INSERT INTO demand (Id_Problem, Total_Period, Period, Demand) VALUES ("' + Id_Problema + '" ,"' + inttostr(Periodo) + '","' + inttostr(i+1) + '","' + floattostr(demanda[i]) + '")';
  ExecSQL;
  end;

  for i := 0 to length(ObjSelecionados) - 1 do
  begin
  if ObjSelecionados[ordemsel[i]] = 'Cost'  then
  begin
    sql.Text := 'INSERT INTO obj_functions (Id_Problem, Objective, Pref_dir) VALUES ("' + Id_Problema + '", "' + ObjSelecionados[ordemsel[i]] + '","' + '0' + '")';
    ExecSQL;

      // salvando no BD
      SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(custos[0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(custos[1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(custos[2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(custos[3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(custos[4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(custos[5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end
  else
  begin
    sql.Text := 'INSERT INTO obj_functions (Id_Problem, Objective, Pref_dir) VALUES ("' + Id_Problema + '", "' + ObjSelecionados[ordemsel[i]] + '","' + '1' + '")';
    ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Regular_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 0])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Overtime_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 1])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

    SQL.Text := 'UPDATE obj_functions SET Subcontrating_Production ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 2])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

          SQL.Text := 'UPDATE obj_functions SET Stock ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 3])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Hiring ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 4])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;

      SQL.Text := 'UPDATE obj_functions SET Resignation ="' + floattostr(ParametrosFuncaoObjetivo[ordemsel[i]-1, 5])
        + '" WHERE Id_Problem = "' + Id_Problema
        + '" AND Objective = "' + ObjSelecionados[ordemsel[i]] + '";';
      ExecSQL;
  end;
  end;






  end;
end;


procedure TIWForm2.btnImportarClick(Sender: TObject);
var
i, j, auxcount : integer;
user_ ,datafinal: string;
begin
lnkplanos.Visible := False;
  with UserSession.ZQuery1 do
  begin
    // inicialmente contar quantos problemas o mesmo usuario já cadastrou
    SQL.Text := 'SELECT COUNT(name) FROM problem WHERE user = "' +
      (ID_user) + '";';
    ExecSQL;
    open;
    auxcount := FieldByName('COUNT(name)').AsInteger;
    close;

    { após saber quantos problemas o mesmo usuario cadastrou é possivel fazer um
    laço de repetição variando de zero ate essa quantidade-1 para pegar os nomes dos problemas }
    // VERIFICAR COMO FAZER ESSE PROCEDIMENTO

    i := 0;
    j := 0;
    SQL.Text := 'SELECT * FROM problem';
    ExecSQL;
    open;
    while i < auxcount do
    begin
      user_ := FieldByName('user').AsString;
      datafinal := FieldByName('conclusion_date').AsString;

      if ID_user = user_ then
      begin
        Setlength(nameproblem, j + 1);
        Setlength(idproblemasuser,j + 1);
        Setlength(datainicioproblema, j + 1);
        Setlength(datafinalproblema, j + 1);

        nameproblem[j] :=FieldByName('name').AsString;
        idproblemasuser[j] := FieldByName('id_problem').AsString;
        datainicioproblema[j] := FieldByName('creation_date').AsString;
        datafinalproblema[j] := FieldByName('conclusion_date').AsString;
        if datafinalproblema[j] = '' then
        datafinalproblema[j] := '02/01/1900';

        INC(j);
        // end;
        INC(i);
      end;
      next;
    end;
    close;
  end;

  if length(nameproblem) = 0 then
  begin
    WebApplication.ShowMessage('There is no problem to be continued');
    Exit;
  end;

  rgnprob.Visible := false;
  rgnretomada.Visible := true;
  btnauxchoose.Enabled := true;
  Retomada := True;

  // colocar o vetor do nome no listbox
  lstbxProjects.Items.Clear;
  for i := 0 to length(nameproblem) - 1 do
  begin
    lstbxProjects.Items.Add(nameproblem[i] + ' - (' + datainicioproblema[i] + ')');
  end;
end;

procedure TIWForm2.btnInputManualClick(Sender: TObject);
begin
rgninput.Visible := True;
rgnprob.Visible := False;
Setlength(ObjSelecionados, 1);
ObjSelecionados[0] := 'Cost';
datainicio := datetostr(now);
retomada := False;
lnkplanos.Visible := False;
end;




procedure TIWForm2.btnLoginClick(Sender: TObject);
begin
rgnintro.Visible := false;
rgnlogin.Visible := true;
end;

procedure TIWForm2.btnredefiniretapa1Click(Sender: TObject);
begin
  edtprodregular.Enabled := True;
  edthorasextras.Enabled := True;
  edtsubcontrat.Enabled := True;
  edtestoq.Enabled := True;
  edtfunc.Enabled := True;
  edtinifunc.Enabled := True;

  edtprodregular.Clear;
  edthorasextras.Clear;
  edtsubcontrat.Clear;
  edtestoq.Clear;
  edtfunc.Clear;
  edtinifunc.Clear;

  cbxproducaoregular.Enabled := True;
  cbxHorasExtras.Enabled := True;
  cbxSubcontratacao.Enabled := True;
  cbxEstoque.Enabled := True;
  cbxcontrat.enabled := True;
  cbxdem.Enabled := True;
  edtnomeproblema.Text := '';
  rgrpPeriodo.ItemIndex := -1;
  lstbxObjetivos.Items.Clear;
  lstbxObjetivos.Items.Add('Selected Objectives:');
  lstbxObjetivos.Items.Add('Cost');
  cmbbxobj.Items.Clear;
  cmbbxobj.Items.Add('Quality');
  cmbbxobj.Items.Add('Reliability');
  cmbbxobj.Items.Add('Speed');
  cmbbxobj.Items.Add('Flexibility');
  cmbbxobj.ItemIndex := -1;

  Finalize(ObjSelecionados);
  Setlength(ObjSelecionados, 1);
  ObjSelecionados[0] := 'Cost';
  rgnadcobjetivo.Visible := False;
  btnavancaretapa1.Enabled := False;
  btnsalvaretapa1.enabled := True;
end;

procedure TIWForm2.btnredefiniretapa2Click(Sender: TObject);
begin
Chrtdemanda.Series[0].Clear;
Chrtdemanda.Series[1].Clear;
Chrtdemanda.Series[0].Visible := False;
Chrtdemanda.Series[1].Visible := False;
Chrtdemanda.Series[2].Visible := True;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));
Finalize(Demanda);
btnsalvaretapa2.Enabled := True;
btnredefiniretapa2.Enabled := False;
btnavancaretapa2.Enabled := False;

end;

procedure TIWForm2.btnredefiniretapa3Click(Sender: TObject);
begin
 cmbbxobjetivos.Enabled := True;
 grdfuncaoobjetivo.Enabled := True;
 btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 if lstbxFuncoesObjetivo.Items.Count > 0 then
 lstbxFuncoesObjetivo.Items.Delete(lstbxFuncoesObjetivo.Items.Count-1);
end;

procedure TIWForm2.btnRegisterClick(Sender: TObject);
begin
 iwrgnCaixaDialogoRU.Visible := true;
end;

procedure TIWForm2.btnreiniciaretapa3Click(Sender: TObject);
var
i : integer;
begin
 Finalize(custos);
 Finalize(ParametrosFuncaoObjetivo);
 Finalize(ordemsel);
 auxiliar := 0;
 cmbbxObjetivos.Items.Clear;
 for I := 0 to Length(ObjSelecionados)-1 do
 cmbbxobjetivos.Items.Add(objselecionados[i]);
 cmbbxobjetivos.ItemIndex := 0;
 cmbbxobjetivos.Enabled := True;
 //grdfuncaoobjetivo.Enabled := True;
 btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 btnexecutar.Enabled := False;
 lstbxFuncoesObjetivo.Clear;
end;

procedure TIWForm2.btnsalvaretapa1Click(Sender: TObject);
var
number : double;
begin
  if (TryStrToFloat(edtprodregular.Text, number) = True) and
   (TryStrToFloat(edthorasextras.Text, number) = True) and
   (TryStrToFloat(edtsubcontrat.Text, number) = True) and
   (TryStrToFloat(edtestoq.Text, number) = True) and
   (TryStrToFloat(edtfunc.Text, number) = True) and
   (TryStrToFloat(edtinifunc.Text, number) = True) then
  begin
  Setlength(Limites, 5);
  Limites[0] := Strtofloat(edtprodregular.Text); // ** Max prod regular
  Limites[1] := Strtofloat(edthorasextras.Text); // ** Max prod h extra
  Limites[2] := Strtofloat(edtsubcontrat.Text); // ** Maxima prod subcontratada
  Limites[3] := Strtofloat(edtestoq.Text); // **Máximo estoque por período
  Limites[4] := Strtofloat(edtfunc.Text); // ** Número máximo de funcionários
  NinicialFunc := Strtoint(edtinifunc.Text); // ** Número inicial de funcionários
  edtprodregular.Enabled := False;
  edthorasextras.Enabled := False;
  edtsubcontrat.Enabled := False;
  edtestoq.Enabled := False;
  edtfunc.Enabled := False;
  edtinifunc.Enabled := False;

  btnsalvaretapa1.enabled := False;
  end
  else
  WebApplication.ShowMessage('Fill in the information correctly!');

end;



procedure TIWForm2.btnsalvaretapa2Click(Sender: TObject);
var
i : integer;
number, aux : double;
nome : string;
begin
SetLength(Demanda, Periodo);
for i := 0 to periodo - 1 do
begin
  nome := grddemanda.DynEdits[1,i];
  if (TryStrToFloat(nome, number) = True) then
  begin
  demanda[i] := StrToFloat(grddemanda.DynEdits[1,i]);
  aux := aux + demanda[i];
  end
  else
  begin
  WebApplication.ShowMessage('Preencha corretamente os dados!');
  exit;
  end;
end;
aux := aux / periodo;
DemandaMedia := aux;
Chrtdemanda.Series[0].Visible := True;
Chrtdemanda.Series[1].Visible := True;
for i := 0 to Periodo -1  do
begin
chrtDemanda.Series[0].AddXY(i+1, demanda[i]);
chrtDemanda.Series[1].AddXY(i+1, aux);
end;
imgdemanda.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (chrtDemanda.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgdemanda.Width,
          imgdemanda.Height)));

btnsalvaretapa2.Enabled := False;
btnredefiniretapa2.Enabled := true;
btnavancaretapa2.Enabled := True;

end;



procedure TIWForm2.btnsalvaretapa3Click(Sender: TObject);
var
i, aux : integer;
begin
SetLength(Custos, 6);
SetLength(ParametrosFuncaoObjetivo, length(ObjSelecionados)-1, 6);

SetLength(ordemsel, length(ObjSelecionados));


aux := 0;

if cmbbxobjetivos.Text = 'Cost'  then
begin
 for I := 0 to 5 do
 custos[i] := StrToFloat(grdfuncaoobjetivo.DynEdits[i,0]);
 lstbxFuncoesObjetivo.Items.Add('Min. ' + cmbbxObjetivos.Text + ' = ' + floattostr(custos[0]) +
 '*x1 + '+ floattostr(custos[1]) +'*x2 + '+ floattostr(custos[2])+ '*x3 + '+ floattostr(custos[3])+ '*x4 + '
 + floattostr(custos[4])+ '*x5 + ' + floattostr(custos[5])+ '*x6');
end
else
begin
  for I := 0 to length(objselecionados) - 1 do
  begin
    if objselecionados[i] = cmbbxobjetivos.Text then
    aux := i;
  end;
 for I := 0 to 5 do
 ParametrosFuncaoObjetivo[aux-1, i] := StrToFloat(grdfuncaoobjetivo.DynEdits[i,0]);
 lstbxFuncoesObjetivo.Items.Add('Max. ' + cmbbxObjetivos.Text + ' = ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 0]) +
 '*x1 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 1]) +'*x2 + '+ floattostr(ParametrosFuncaoObjetivo[aux-1, 2])+ '*x3 '
 + floattostr(ParametrosFuncaoObjetivo[aux-1, 3]) +'*x4 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 4]) +'*x5 + ' + floattostr(ParametrosFuncaoObjetivo[aux-1, 5]) +'*x6');
 end;


 ordemsel[auxiliar] := aux;
 inc(auxiliar);

 cmbbxobjetivos.Items.Delete(cmbbxobjetivos.ItemIndex);
 cmbbxobjetivos.Enabled := True;
 grdfuncaoobjetivo.Enabled := True;
 btnredefiniretapa3.Enabled := False;
 btnsalvaretapa3.Enabled := True;
 if cmbbxobjetivos.Items.Count = 0 then
 begin
 cmbbxobjetivos.Enabled := False;
 grdfuncaoobjetivo.Enabled := False;
 btnsalvaretapa3.Enabled := False;
 btnredefiniretapa3.Enabled := True;
 btnexecutar.Enabled := True;
 end;
end;



procedure TIWForm2.cbxcontratClick(Sender: TObject);
begin
if cbxcontrat.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,4] := '0';
grdfuncaoobjetivo.Columns.Items[4].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[4].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxdemClick(Sender: TObject);
begin
if cbxdem.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,5] := '0';
grdfuncaoobjetivo.Columns.Items[5].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[5].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxEstoqueClick(Sender: TObject);
begin
if cbxestoque.Checked = False then
begin
edtestoq.Text := '0';
edtestoq.Enabled := False;
end
else
edtestoq.Enabled := True;

if cbxestoque.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,3] := '0';
grdfuncaoobjetivo.Columns.Items[3].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[3].ColumnType := ctDynEdit;

end;



procedure TIWForm2.cbxHorasExtrasClick(Sender: TObject);
begin
if cbxhorasextras.Checked = False then
begin
edthorasextras.Text := '0';
edthorasextras.Enabled := False;
end
else
edthorasextras.Enabled := True;

if cbxhorasextras.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,1] := '0';
grdfuncaoobjetivo.Columns.Items[1].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[1].ColumnType := ctDynEdit;

end;


procedure TIWForm2.cbxproducaoregularClick(Sender: TObject);
begin
if cbxproducaoregular.Checked = False then
begin
edtprodregular.Text := '0';
edtprodregular.Enabled := False;
end
else
edtprodregular.Enabled := True;

if cbxproducaoregular.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,0] := '0';
grdfuncaoobjetivo.Columns.Items[0].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[0].ColumnType := ctDynEdit;
end;

procedure TIWForm2.cbxSubcontratacaoClick(Sender: TObject);
begin
if cbxsubcontratacao.Checked = False then
begin
edtsubcontrat.Text := '0';
edtsubcontrat.Enabled := False;
end
else
edtsubcontrat.Enabled := True;

if cbxsubcontratacao.Checked = False then
begin
grdfuncaoobjetivo.DynEdits[0,2] := '0';
grdfuncaoobjetivo.Columns.Items[2].ColumnType := ctNormal;
end
else
grdfuncaoobjetivo.Columns.Items[2].ColumnType := ctDynEdit;
end;








procedure TIWForm2.edtprodregularAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtprodregular.Text[Length(edtprodregular.Text)] <> '0') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '1') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '2') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '3') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '4') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '5') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '6') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '7') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '8') and
    (edtprodregular.Text[Length(edtprodregular.Text)] <> '9') then
    edtprodregular.Text := Copy(edtprodregular.Text, 1, length(edtprodregular.Text) - 1)
    else
    edtprodregular.Text := edtprodregular.Text;
end;




procedure TIWForm2.edtsubcontratAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '0') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '1') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '2') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '3') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '4') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '5') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '6') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '7') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '8') and
    (edtsubcontrat.Text[Length(edtsubcontrat.Text)] <> '9') then
    edtsubcontrat.Text := Copy(edtsubcontrat.Text, 1, length(edtsubcontrat.Text) - 1)
    else
    edtsubcontrat.Text := edtsubcontrat.Text;
end;



procedure TIWForm2.IWButton1Click(Sender: TObject);
begin
lstbxobjetivos.Items.Add(cmbbxobj.Text);
SetLength(ObjSelecionados, Length(ObjSelecionados)+1);
ObjSelecionados[Length(ObjSelecionados)-1] := cmbbxobj.Text;
cmbbxobj.Items.Delete(cmbbxobj.ItemIndex);
btnavancaretapa1.Enabled := True;
end;

procedure TIWForm2.IWButton2Click(Sender: TObject);
begin
lstbxobjetivos.Items.Add(edtnomeobjetivo.Text);
SetLength(ObjSelecionados, Length(ObjSelecionados)+1);
ObjSelecionados[Length(ObjSelecionados)-1] := edtnomeobjetivo.Text;
edtnomeobjetivo.Text := '';
btnavancaretapa1.Enabled := True;
end;

procedure TIWForm2.IWButton4Click(Sender: TObject);
begin
  rgnlogin.Visible := false;
  rgnintro.Visible := true;
  lblwarningemail.Visible := false;
  lblwarningpassword.Visible := false;

  /// /////////////////////////////////////////////////////////////////////////////
  /// ////////////* Limpando edit's de login, validacao *//////////////////////////
  edtemail.Clear;
  edtSenha.Clear;
  /// /////////////////////////////////////////////////////////////////////////////
end;

procedure TIWForm2.lnkplanosClick(Sender: TObject);
begin
TIWAdvWebGridExcelIO2.AdvWebGrid := grdplanos;
  TIWAdvWebGridExcelIO2.XLSExport('Plans_' + problema + '.xls', WebApplication);
  Inherited;
end;

procedure TIWForm2.IWLink1Click(Sender: TObject);
begin
rgnadcobjetivo.Visible := True;
end;

procedure TIWForm2.IWLink2Click(Sender: TObject);
begin
btnredefiniretapa3Click(Self);
btnredefiniretapa2Click(Self);
btnredefiniretapa1Click(Self);
rgninput.Visible := False;
rgnadcobjetivo.Visible := False;
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := False;
rgnintro.Visible := False;
rgnlogin.Visible := False;
rgnretomada.Visible := False;
rgnprob.Visible := True;
end;

procedure TIWForm2.IWLink3Click(Sender: TObject);
begin
rgnadcobjetivo.Visible := False;
end;



procedure TIWForm2.IWLink4Click(Sender: TObject);
begin
rgnprob.Visible := True;
rgnretomada.Visible := False;
end;

procedure TIWForm2.IWLink5Click(Sender: TObject);
begin
rgnprob.visible := False;
rgnlogin.Visible := True;
edtemail.Clear;
edtSenha.Clear;
end;

procedure TIWForm2.IWLink6Click(Sender: TObject);
begin
btnredefiniretapa3Click(Self);
btnredefiniretapa2Click(Self);
btnredefiniretapa1Click(Self);
rgninput.Visible := False;
rgnadcobjetivo.Visible := False;
rgndemanda.Visible := False;
rgnfuncaoobjetivo.Visible := False;
rgnintro.Visible := False;
rgnlogin.Visible := False;
rgnretomada.Visible := False;
rgnprob.Visible := True;
end;

procedure TIWForm2.IWLink7Click(Sender: TObject);
begin
rgndemanda.Visible := False;
rgninput.Visible := True;
btnredefiniretapa2Click(Self);
end;

procedure TIWForm2.IWLink8Click(Sender: TObject);
begin
rgnfuncaoobjetivo.Visible := False;
rgndemanda.Visible := True;
btnredefiniretapa3Click(Self);
end;

procedure TIWForm2.lnkesqueceusenhaClick(Sender: TObject);
begin
iwrgnCaixaDialogoRU2.Visible := true;
end;

procedure TIWForm2.rgndemandaCreate(Sender: TObject);
begin
rgndemanda.align := AlClient;
end;

procedure TIWForm2.rgnfuncaoobjetivoCreate(Sender: TObject);
begin
rgnfuncaoobjetivo.align := AlClient;
end;

procedure TIWForm2.rgninputCreate(Sender: TObject);
begin
rgninput.align := AlClient;
end;

procedure TIWForm2.rgnprobCreate(Sender: TObject);
begin
rgnprob.Align := AlClient;
end;

procedure TIWForm2.rgnretomadaCreate(Sender: TObject);
begin
rgnretomada.Align := AlClient;
end;

procedure TIWForm2.rgnintroCreate(Sender: TObject);
begin
rgnintro.align := AlClient;
end;

procedure TIWForm2.rgnloginCreate(Sender: TObject);
begin
 rgnlogin.Align := AlClient;
end;

procedure TIWForm2.rgrpPeriodoClick(Sender: TObject);
begin
if rgrpPeriodo.ItemIndex = 0 then
Periodo := 12;
if rgrpPeriodo.ItemIndex = 1 then
Periodo := 6;
if rgrpPeriodo.ItemIndex = 2 then
Periodo := 4;


end;



procedure TIWForm2.edtestoqAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtestoq.Text[Length(edtestoq.Text)] <> '0') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '1') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '2') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '3') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '4') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '5') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '6') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '7') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '8') and
    (edtestoq.Text[Length(edtestoq.Text)] <> '9') then
    edtestoq.Text := Copy(edtestoq.Text, 1, length(edtestoq.Text) - 1)
    else
    edtestoq.Text := edtestoq.Text;
end;

procedure TIWForm2.edtfuncAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtfunc.Text[Length(edtfunc.Text)] <> '0') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '1') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '2') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '3') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '4') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '5') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '6') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '7') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '8') and
    (edtfunc.Text[Length(edtfunc.Text)] <> '9') then
    edtfunc.Text := Copy(edtfunc.Text, 1, length(edtfunc.Text) - 1)
    else
    edtfunc.Text := edtfunc.Text;
end;

procedure TIWForm2.edthorasextrasAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edthorasextras.Text[Length(edthorasextras.Text)] <> '0') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '1') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '2') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '3') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '4') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '5') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '6') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '7') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '8') and
    (edthorasextras.Text[Length(edthorasextras.Text)] <> '9') then
    edthorasextras.Text := Copy(edthorasextras.Text, 1, length(edthorasextras.Text) - 1)
    else
    edthorasextras.Text := edthorasextras.Text;
end;

procedure TIWForm2.edtinifuncAsyncKeyPress(Sender: TObject;
  EventParams: TStringList);
begin
if (edtinifunc.Text[Length(edtinifunc.Text)] <> '0') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '1') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '2') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '3') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '4') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '5') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '6') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '7') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '8') and
    (edtinifunc.Text[Length(edtinifunc.Text)] <> '9') then
    edtinifunc.Text := Copy(edtinifunc.Text, 1, length(edtinifunc.Text) - 1)
    else
    edtinifunc.Text := edtinifunc.Text;
end;
initialization
  TIWForm2.SetAsMainForm;

end.
