unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons,ExtCtrls,ComObj, RzPanel, RzPrgres;

type
  TForm3 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Edit1: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    OpenDialog1: TOpenDialog;
    RzGroupBox1: TRzGroupBox;
    RzProgressBar1: TRzProgressBar;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    ExcelApp1:Variant;
    ExcelApp2:Variant;
    fahuoListMaxRow:Integer;
    fahuoListMaxCol:Integer;
    dinghuorenCol:TStringList;
  end;

var
  Form3: TForm3;
  fahuoListRow:Integer;
  dinghuoListRow:Integer;
  fahuoListKuanHaoCol:Integer;
  dinghuoListKuanHaoCol:Integer;
  fahuoListYanSeCol:Integer;
  dinghuoListYanSeCol:Integer;
  fahuoListChiMaCol:Integer;
  dinghuoListChiMaCol:Integer;
  fahuolistPeihuoliangCol:Integer;
  dinghuoListShuLiangCol:Integer;

implementation

uses Unit1, Unit2;

{$R *.dfm}

procedure TForm3.BitBtn1Click(Sender: TObject);
var
i,j,k:Integer;
begin
dinghuorenCol:=TStringList.Create;
Opendialog1.Filter:='(EXCEL�ļ�.xls)|*.xls';//�á�|���ֿ�
Opendialog1.InitialDir:='C:\Users\Administrator\Desktop';
if OpenDialog1.Execute then
begin
  Form3.BitBtn2.Enabled:=true;
  Edit1.Text:=OpenDialog1.FileName;
end;
if Edit1.Text<>'' then
begin
ExcelApp1:=CreateOleObject('Excel.Application');
ExcelApp1.Caption:='Ӧ�ó������ Microsoft Excel';
ExcelApp1.workBooks.Open(Edit1.Text); //���Ѵ��ڹ�����
ExcelApp1.Visible:=true;
ExcelApp1.WorkSheets[1].Activate;
//excel1OpenFlag := true;
fahuoListMaxRow:= ExcelApp1.ActiveSheet.UsedRange.Rows.Count;
fahuoListMaxCol:= ExcelApp1.ActiveSheet.UsedRange.Columns.Count;
for k:=1 to  fahuoListMaxCol do
begin
  if ExcelApp1.ActiveSheet.Cells[1,k].Value='���' then
      fahuoListKuanHaoCol:=k;
  if ExcelApp1.ActiveSheet.Cells[1,k].Value='��ɫ' then
      fahuoListYanSeCol:=k;
  if ExcelApp1.ActiveSheet.Cells[1,k].Value='����' then
     fahuoListChiMaCol:=k;
  if ExcelApp1.ActiveSheet.Cells[1,k].Value='���������' then
     fahuolistPeihuoliangCol:=k;
end;
for i:=1 to  fahuoListMaxCol do
begin
  for j:= 0 to Form1.dinghuorenlist.Count-1 do
  begin
      if Pos(Form1.dinghuorenList[j]+'��',ExcelApp1.ActiveSheet.Cells[1,i].Value) > 0 then
      dinghuorenCol.Add(Form1.dinghuorenList[j] + '=' + inttostr(i));
  end;
end;
end;
end;

procedure TForm3.BitBtn2Click(Sender: TObject);
var
i,j,k,m:Integer;
dinghuoListMaxRow:Integer;
dinghuoListMaxCol:Integer;
yipeihuoliangCol:Integer;
begin
  for  i:=0 to Form1.dinghuoListPath.Count-1 do
  begin
    ExcelApp2:=CreateOleObject('Excel.Application');
    ExcelApp2.Caption:='Ӧ�ó������ Microsoft Excel';
    ExcelApp2.workBooks.Open(Form1.dinghuoListPath[i]); //���Ѵ��ڹ�����
    ExcelApp2.Visible:=true;
    ExcelApp2.WorkSheets[1].Activate;
    dinghuoListMaxRow:=ExcelApp2.ActiveSheet.UsedRange.Rows.Count;
    dinghuoListMaxCol:= ExcelApp2.ActiveSheet.UsedRange.Columns.Count;
    for m:= 1 to dinghuoListMaxCol do
    begin
       if ExcelApp2.ActiveSheet.Cells[1,m].Value='���' then
        dinghuoListKuanHaoCol:=m;
      if ExcelApp2.ActiveSheet.Cells[1,m].Value='��ɫ' then
        dinghuoListYanSeCol:=m;
      if ExcelApp2.ActiveSheet.Cells[1,m].Value='����' then
        dinghuoListChiMaCol:=m;
      if ExcelApp2.ActiveSheet.Cells[1,m].Value='����' then
        dinghuoListShuLiangCol:=m ;
    end;
    if Pos('�����',ExcelApp2.ActiveSheet.Cells[1,dinghuoListMaxCol].Value) > 0 then
    begin
      yipeihuoliangCol:=dinghuoListMaxCol;
    end else
    begin
      ExcelApp2.ActiveSheet.Cells[1,dinghuoListMaxCol+1].Value:= '�������';
      yipeihuoliangCol:=dinghuoListMaxCol+1;
    end;
    Form3.RzProgressBar1.TotalParts:= Form1.fahuoListMaxRow-3;
    for k:=2 to Form1.fahuoListMaxRow-1 do
    begin
    Form3.RzProgressBar1.PartsComplete:=k-2;
      for j:=2 to   dinghuoListMaxRow-1 do
      begin
        ExcelApp2.ActiveSheet.Cells[j,yipeihuoliangCol].Value:= '0';
        if (Trim(ExcelApp2.ActiveSheet.Cells[j,dinghuoListKuanHaoCol].Value) = Trim(ExcelApp1.ActiveSheet.Cells[k,fahuoListKuanHaoCol].Value)) and
        (Trim(ExcelApp2.ActiveSheet.Cells[j,dinghuoListYanSeCol].Value) = Trim(ExcelApp1.ActiveSheet.Cells[k,fahuoListYanSeCol].Value))
        and (Trim(ExcelApp2.ActiveSheet.Cells[j,dinghuoListChiMaCol].Value) = Trim(ExcelApp1.ActiveSheet.Cells[k,fahuoListChiMaCol].Value)) then
        begin
           //if Trim(ExcelApp2.ActiveSheet.Cells[j,yipeihuoliangCol].Value) = '' then
          // begin
              //ExcelApp2.ActiveSheet.Cells[j,yipeihuoliangCol].Value:= Trim(ExcelApp1.ActiveSheet.Cells[k,strtoint(dinghuorenCol.ValueFromIndex[i])].Value);
              //ExcelApp2.ActiveSheet.Cells[j,dinghuoListShuLiangCol].Value:=inttostr(strtoint(Trim(ExcelApp2.ActiveSheet.Cells[j,dinghuoListShuLiangCol].Value)) - strtoint(Trim(ExcelApp2.ActiveSheet.Cells[j,yipeihuoliangCol].Value)));
              //break;
           //end else
           //begin
              ExcelApp2.ActiveSheet.Cells[j,yipeihuoliangCol].Value:= inttostr(strtoint(Trim(ExcelApp1.Cells[k,strtoint(dinghuorenCol.ValueFromIndex[i])].Value)) + strtoint(Trim(ExcelApp2.Cells[j,yipeihuoliangCol].Value)));
              ExcelApp2.ActiveSheet.Cells[j,dinghuoListShuLiangCol].Value:=inttostr(strtoint(Trim(ExcelApp2.Cells[j,dinghuoListShuLiangCol].Value)) - strtoint(Trim(ExcelApp2.Cells[j,yipeihuoliangCol].Value)));
              break;
          // end;
        end;
      end;
    end;
   // ExcelApp2.saveAS(ExtractFileName(Form1.dinghuoListPath[i]+'(' + inttostr(i) + ')'));    //����
    ExcelApp2.ActiveWorkBook.save;  
    ExcelApp2.WorkBooks.Close; //�رչ�����
    ExcelApp2.Quit; //�˳� Excel
    ExcelApp2:=Unassigned;//�ͷ�excel����
  end;
  ExcelApp1.ActiveWorkBook.save;    //����
  ExcelApp1.WorkBooks.Close; //�رչ�����
  ExcelApp1.Quit; //�˳� Excel
  ExcelApp1:=Unassigned;//�ͷ�excel����
end;

procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Close;
end;

end.
