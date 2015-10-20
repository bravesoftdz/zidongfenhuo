unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XPMan, ComCtrls, Buttons, RzPrgres;

type
  TForm2 = class(TForm)
    XPManifest1: TXPManifest;
    Label1: TLabel;
    BitBtn1: TBitBtn;
    RzProgressBar1: TRzProgressBar;
    SaveDialog1: TSaveDialog;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    fahuoResult:String;
  end;

var
  Form2: TForm2;

implementation

uses Unit1, Unit3;

{$R *.dfm}

procedure TForm2.BitBtn1Click(Sender: TObject);
begin
SaveDialog1.DefaultExt:='xls';
SaveDialog1.Filter:='Excel文件(*.xls)|*.xls';
if SaveDialog1.Execute then
begin
   try
   Form1.ExcelApp3.ActiveWorkBook.save;    //保存
   fahuoResult:= SaveDialog1.FileName;   //保存最后结果的文件地址
   Form1.ExcelApp3.ActiveWorkBook.SaveAs(SaveDialog1.FileName);
   Form1.ExcelApp3.WorkBooks.Close; //关闭工作簿
   Form1.ExcelApp3.Quit; //退出 Excel
   Form1.ExcelApp3:=Unassigned;//释放excel进程
   ShowMessage('保存成功，点击下一步继续！');
   Form2.BitBtn1.Enabled:=false;
   Form2.BitBtn2.Enabled:=true;
   except
     ShowMessage('保存失败！');
   end;
end;
end;

procedure TForm2.BitBtn2Click(Sender: TObject);
begin
  Form2.Hide;
  Form3.Show;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Close;
end;

end.
