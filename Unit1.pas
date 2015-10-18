unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ExtCtrls,ComObj, XPMan, ComCtrls, Buttons, Menus,
  RzButton, RzPanel;

type
  TForm1 = class(TForm)
    Button3: TButton;
    Edit2: TEdit;
    Edit3: TEdit;
    OpenDialog1: TOpenDialog;
    XPManifest1: TXPManifest;
    Button6: TButton;
    Edit1: TEdit;
    BitBtn1: TBitBtn;
    ListBox1: TListBox;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    StatusBar1: TStatusBar;
    N5: TMenuItem;
    N6: TMenuItem;
    procedure Button3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button6Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    procedure startAutoReorganize;
    { Private declarations }
  public
    { Public declarations }
    ExcelApp1:Variant;
    ExcelApp2:Variant;
    ExcelApp3:Variant;
    ExcelApp4:Variant;
    dinghuorenList:TStringList;
    dinghuoListPath:TStringList; //�����嵥�ļ�·��
    fahuoListMaxRow:Integer;
    fahuoListMaxCol:Integer;
    dinghuoListMaxRow:Integer;
    dinghuoListMaxCol:Integer;
  end;

var
  Form1: TForm1;
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



  excel1OpenFlag:Boolean;
  excel2OpenFlag:Boolean;

  dinghuoren:String;      //����������



implementation

uses Unit2;

{$R *.dfm}




procedure TForm1.Button3Click(Sender: TObject);
var
k,j,i,m:Integer;
begin
dinghuoListPath := TStringList.Create;
Form2.RzProgressBar1.TotalParts:= fahuoListMaxRow-3;
if listbox1.Items.Count<>0 then
begin
  Form2.Show;
  Form2.Label1.Left := Round((Form2.Width - Form2.Label1.Width) / 2) ;
  m:=0;
  j:=0;
  i:=0;
  for m:=0 to listbox1.Items.Count-1 do
  begin
    dinghuoListPath.Add(listbox1.Items[m]);
    ExcelApp2:=CreateOleObject('Excel.Application');
   // ExcelApp2.visible:=true;
    ExcelApp2.Caption:='Ӧ�ó������ Microsoft Excel';
    ExcelApp2.workBooks.Open(listbox1.Items[m]); //���Ѵ��ڹ�����
    //Edit1.Text:=ExcelApp2.WorkSheets[1].Cells.Find('po');
    Form2.Label1.Caption:= '���ڴ�����'+inttostr(m+1)+'�������嵥����'+inttostr(listbox1.Items.Count)+'�������Ժ�...';
    dinghuoren := Copy(ExtractFileName(listbox1.Items[m]), 0, pos('��',ExtractFileName(listbox1.Items[m]))-1);
    excel2OpenFlag := true;
    dinghuoListMaxRow:=ExcelApp2.WorkSheets[1].UsedRange.Rows.Count;
    dinghuoListMaxCol:= ExcelApp2.WorkSheets[1].UsedRange.Columns.Count;
    for i:=1 to  dinghuoListMaxCol do
    begin
      if ExcelApp2.Cells[1,i].Value='���' then
        dinghuoListKuanHaoCol:=i;
      if ExcelApp2.Cells[1,i].Value='��ɫ' then
        dinghuoListYanSeCol:=i;
      if ExcelApp2.Cells[1,i].Value='����' then
        dinghuoListChiMaCol:=i;
      if ExcelApp2.Cells[1,i].Value='����' then
        dinghuoListShuLiangCol:=i ;
    end;
    ExcelApp1.Cells[1,fahuoListMaxCol+m+1].Value:= dinghuoren+'������';
    dinghuorenList.Add(dinghuoren);
    for k:=2 to fahuoListMaxRow-1 do
    begin
      Form2.RzProgressBar1.PartsComplete:=k-2;
      for j:=2 to dinghuoListMaxRow-1 do
      begin
        if (Trim(ExcelApp2.Cells[j,dinghuoListKuanHaoCol].Value) = Trim(ExcelApp1.Cells[k,fahuoListKuanHaoCol].Value)) and
        (Trim(ExcelApp2.Cells[j,dinghuoListYanSeCol].Value) = Trim(ExcelApp1.Cells[k,fahuoListYanSeCol].Value))
        and (Trim(ExcelApp2.Cells[j,dinghuoListChiMaCol].Value) = Trim(ExcelApp1.Cells[k,fahuoListChiMaCol].Value)) then
        begin
          ExcelApp1.Cells[k,fahuoListMaxCol+m+1].Value:= ExcelApp2.Cells[j,dinghuoListShuLiangCol].Value;
          break;
        end else
        begin
          if j=dinghuoListMaxRow-1 then
            ExcelApp1.Cells[k,fahuoListMaxCol+m+1].Value:= '0';
        end;
      end;
    end;
    ExcelApp2.WorkBooks.Close; //�رչ���t
    ExcelApp2.Quit; //�˳� Excel
    ExcelApp2:=Unassigned;//�ͷ�excel����
    excel2OpenFlag:=false;
  end;
end;
    ExcelApp1.ActiveWorkBook.save;    //����
    ExcelApp1.WorkBooks.Close; //�رչ�����
    ExcelApp1.Quit; //�˳� Excel
    excel1OpenFlag:=false;
    ExcelApp1:=Unassigned;//�ͷ�excel����
    Button6.Click;
end;




procedure TForm1.FormShow(Sender: TObject);
begin
Edit2.Enabled:=false;
Edit3.Enabled:=false;
dinghuorenList:=TStringList.Create;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if excel1OpenFlag = true then
begin
  ExcelApp1.WorkBooks.Close; //�رչ���t
  ExcelApp1.Quit; //�˳� Excel
  ExcelApp1:=Unassigned;//�ͷ�excel����
end;
if excel2OpenFlag = true then
begin
  ExcelApp2.WorkBooks.Close; //�رչ���t
  ExcelApp2.Quit; //�˳� Excel
  ExcelApp2:=Unassigned;//�ͷ�excel����
end;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
i:Integer;
begin
Edit1.Text:=inttostr(dinghuorenList.Count);
Opendialog1.Filter:='(EXCEL�ļ�.xls)|*.xls';//�á�|���ֿ�
Opendialog1.InitialDir:='C:\Users\Administrator\Desktop';
if Edit2.Text<>'' then
begin
ExcelApp3:=CreateOleObject('Excel.Application');
ExcelApp3.Caption:='Ӧ�ó������ Microsoft Excel';
ExcelApp3.visible:=true;
ExcelApp3.workBooks.Open(Edit2.Text); //���Ѵ��ڹ�����
fahuoListMaxRow:= ExcelApp3.WorkSheets[1].UsedRange.Rows.Count;
fahuoListMaxCol:= ExcelApp3.WorkSheets[1].UsedRange.Columns.Count;
for i:=(fahuoListMaxCol+1) to fahuoListMaxCol+dinghuorenList.Count do
begin
  ExcelApp3.Cells[1,i].Value:= dinghuorenList[i-fahuoListMaxCol-1]+'�����';
end;
startAutoReorganize();
//ExcelApp3.ActiveWorkBook.save;    //����
//ExcelApp3.WorkBooks.Close; //�رչ�����
//xcelApp3.Quit; //�˳� Excel
//excel1OpenFlag:=false;
//ExcelApp3:=Unassigned;//�ͷ�excel����
end;
end;

procedure TForm1.startAutoReorganize;
var
j,k,l,m,dinghuoSum,peihuoSum:Integer;
dinghuozhanbiList:TStringList;
yifenpeiSum:Integer;
yifenpeiNum:Integer;
breakFlag:boolean;
begin
Form2.RzProgressBar1.TotalParts:= fahuoListMaxRow-3;
Form2.Label1.Caption:='�����Զ��ֻ������Ժ�...' ;
Form2.Label1.Left := Round((Form2.Width - Form2.Label1.Width) / 2) ;
for k:=2 to fahuoListMaxRow-1 do
begin
  Form2.RzProgressBar1.PartsComplete:=k-2;
  dinghuozhanbiList:= TStringList.Create;
  dinghuozhanbiList.Sorted:=true;
  dinghuozhanbiList.Clear;
  dinghuoSum:=0;
  yifenpeiSum:=0;
  yifenpeiNum:=0;
  breakFlag:=false;
  peihuoSum:= strtoint(Trim(ExcelApp3.Cells[k,fahuolistPeihuoliangCol].Value));
  //ExcelApp3.ActiveSheet.Rows[k].Interior.Color:=clMoneyGreen;
  ExcelApp3.ActiveSheet.Rows[k].Borders[8].Weight := 2;
  ExcelApp3.ActiveSheet.Rows[k].Borders[9].Weight := 2;
  ExcelApp3.ActiveSheet.Rows[k].Borders[11].Weight := 2;
  for j:= fahuoListMaxCol-dinghuorenList.Count+1 to fahuoListMaxCol do    //ͳ�����ж�����
  begin
    dinghuoSum:=dinghuoSum+strtoint( Trim(ExcelApp3.Cells[k,j].Value));
  end;
  for j:= fahuoListMaxCol-dinghuorenList.Count+1 to fahuoListMaxCol do
  begin
    if  ExcelApp3.Cells[k,j].Value ='0' then
    begin
      ExcelApp3.Cells[k,j+dinghuorenList.Count].Value:='0';
    end else if  peihuoSum>=dinghuoSum then
    begin
      if peihuoSum > dinghuoSum then
      begin
        ExcelApp3.ActiveSheet.Rows[k].Font.Color := clRed;
        ExcelApp3.ActiveSheet.Rows[k].Font.Bold := True;
        ExcelApp3.Cells[k,j+dinghuorenList.Count].Value:=ExcelApp3.Cells[k,j].Value;
      end else
      begin
         ExcelApp3.Cells[k,j+dinghuorenList.Count].Value:=ExcelApp3.Cells[k,j].Value;
      end;
    end else
    begin
      ExcelApp3.ActiveSheet.Rows[k].Font.Color := clBlue;
      ExcelApp3.ActiveSheet.Rows[k].Font.Bold := True;
      dinghuozhanbiList.Add(floattostr(strtoint(Trim(ExcelApp3.Cells[k,j].Value))/dinghuoSum)+'='+inttostr(j));
    end;
    if j=fahuoListMaxCol then
    begin
      dinghuozhanbiList.Sort;
       if  dinghuozhanbiList.Count=1 then
       begin
          ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[0]))+dinghuorenList.Count].Value:=inttoStr(Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[0])));
       end else if dinghuozhanbiList.Count > 1 then
       begin
          //if peihuoSum*strtofloat(dinghuozhanbiList.Names[dinghuozhanbiList.Count-1]) < 0.95  then
          //begin
              //for l:=0 to dinghuozhanbiList.Count-2 do
              //begin
                //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[l])+dinghuorenList.Count].Value:=inttoStr(Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[l])));
                //yifenpeiSum:=yifenpeiSum+ Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[l]));
                //if (peihuoSum-yifenpeiSum-Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[l+1])))<0 then
                //begin
                  //yifenpeiNum:=l;
                  //breakFlag:=true;
                  //break;
                //end;
              //end;
              //if breakFlag=true then
              //begin
                //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[yifenpeiNum+1])+dinghuorenList.Count].Value:= inttostr(peihuoSum-yifenpeiSum);
                //for m:=yifenpeiNum+2 to dinghuozhanbiList.Count-1 do
                //begin
                  //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[m])+dinghuorenList.Count].Value:='0';
                //end;
              //end else
              //begin
                //if peihuoSum-yifenpeiSum > strtoint(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])].Value)  then
                //begin
                   //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])+dinghuorenList.Count].Value:= ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])].Value;
                   //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])+dinghuorenList.Count].Value:=inttostr(strtoint(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])+dinghuorenList.Count].Value)+peihuoSum-yifenpeiSum-strtoint(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])].Value));
                //end else
                //begin
                //ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])+dinghuorenList.Count].Value:= inttostr(peihuoSum-yifenpeiSum);
                //end;
              //end;
          //end else
          //begin
            for l:=0 to dinghuozhanbiList.Count-2 do
            begin
              ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-l-1]))+dinghuorenList.Count].Value:=inttoStr(Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[dinghuozhanbiList.Count-l-1])+0.0001));
              yifenpeiSum:=yifenpeiSum+ Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[dinghuozhanbiList.Count-l-1])+0.0001);
              if (peihuoSum-yifenpeiSum-Round(peihuoSum*strtofloat(dinghuozhanbiList.Names[dinghuozhanbiList.Count-l-2])+0.0001))<0 then
              begin
                yifenpeiNum:=l;
                breakFlag:=true;
                break;
              end;
            end;
            if breakFlag=true then
            begin
              ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-yifenpeiNum-2]))+dinghuorenList.Count].Value:= inttostr(peihuoSum-yifenpeiSum);
              for m:=0 to dinghuozhanbiList.Count-yifenpeiNum-3 do
              begin
                ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[m]))+dinghuorenList.Count].Value:='0';
              end;
            end else
            begin
              if peihuoSum-yifenpeiSum > strtoint(Trim(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])].Value))  then
                begin
                   ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1]))+dinghuorenList.Count].Value:=inttostr(strtoint(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[dinghuozhanbiList.Count-1])+dinghuorenList.Count].Value)+peihuoSum-yifenpeiSum-strtoint(ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])].Value));
                   ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])+dinghuorenList.Count].Value:= ExcelApp3.Cells[k,strtoint(dinghuozhanbiList.ValueFromIndex[0])].Value;
                end else
                begin
                  ExcelApp3.Cells[k,strtoint(Trim(dinghuozhanbiList.ValueFromIndex[0]))+dinghuorenList.Count].Value:= inttostr(peihuoSum-yifenpeiSum);
                end;
            end;
          end;
       end;
    end;
  end;
BitBtn3.Enabled:=true;
BitBtn2.Enabled:=true;
Form2.Label1.Caption:='�Զ��ֻ���ɣ��뱣��ֻ����'  ;
Form2.Label1.Left := Round((Form2.Width - Form2.Label1.Width) / 2) ;
Form2.BitBtn1.Enabled:=true;
Edit2.Text:='';
Listbox1.Items.Clear;
Form2.BitBtn1.Enabled:=true;
end;







procedure TForm1.BitBtn3Click(Sender: TObject);
var
i:Integer;
begin
Opendialog1.Filter:='(EXCEL�ļ�.xls)|*.xls';//�á�|���ֿ�
Opendialog1.InitialDir:='C:\Users\Administrator\Desktop';
if OpenDialog1.Execute then
begin
  Edit2.Text:=OpenDialog1.FileName;
end;
if Edit2.Text<>'' then
begin
ExcelApp1:=CreateOleObject('Excel.Application');
ExcelApp1.Caption:='Ӧ�ó������ Microsoft Excel';
ExcelApp1.workBooks.Open(Edit2.Text); //���Ѵ��ڹ�����
//ExcelApp1.Visible:=true;
excel1OpenFlag := true;
fahuoListMaxRow:= ExcelApp1.WorkSheets[1].UsedRange.Rows.Count;
fahuoListMaxCol:= ExcelApp1.WorkSheets[1].UsedRange.Columns.Count;
for i:=1 to  fahuoListMaxCol do
begin
  if ExcelApp1.Cells[1,i].Value='���' then
      fahuoListKuanHaoCol:=i;
  if ExcelApp1.Cells[1,i].Value='��ɫ' then
      fahuoListYanSeCol:=i;
  if ExcelApp1.Cells[1,i].Value='����' then
     fahuoListChiMaCol:=i;
  if ExcelApp1.Cells[1,i].Value='���������' then
     fahuolistPeihuoliangCol:=i;
end;
end;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
var
j:Integer;
begin
Opendialog1.Filter:='(EXCEL�ļ�,��һ�δ򿪶���ļ�.xls)|*.xls';//�á�|���ֿ�
Opendialog1.InitialDir:='C:\Users\Administrator\Desktop';
if OpenDialog1.Execute then
begin
  for j:=0 to Opendialog1.Files.Count-1 do
  begin
      listbox1.Items.Add(OpenDialog1.Files.Strings[j]);
  end;
  //Edit2.Text:=OpenDialog1.FileName;
end;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
if (Edit2.Text='') or (listbox1.Items.Count=0) then
begin
   ShowMessage('�����嵥�򶩻��嵥Ϊ��');
end else
begin
  Form1.Hide;
  BitBtn3.Enabled:=false;
  BitBtn2.Enabled:=false;
  Button3.Click;
end;

end;

end.