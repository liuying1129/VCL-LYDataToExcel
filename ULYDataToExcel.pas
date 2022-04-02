{**********************************************************}
{                                                          }
{  数据集导出Excel组件:TLYDataToExcel Version 05.04.30     }
{                                                          }
{  作者：刘鹰                                              }
{                                                          }
{                                                          }
{  新功能：1.                                              }
{          2.                                              }
{          3.                                              }
{          4.                                              }
{                                                          }
{  功能:                                                   }
{  调用方法：                                              }
{begin                                                     }
{  LYDataToExcel1.DataSet:= adoquery2;                     }
{  LYDataToExcel1.ExcelTitel:='标题';                      }
{  LYDataToExcel1.Execute;                                 }
{end;                                                      }
{                                                          }
{                                                          }
{  他是一个免费软件,如果你修改了他,希望我能有幸看到你的杰作}
{                                                          }
{  我的 Email: Liuying1129@163.com                         }
{                                                          }
{  版权所有,欲用于商业用途,请与我联系!!!                   }
{                                                          }
{                                                          }
{**********************************************************}

unit ULYDataToExcel;

interface

uses
  Windows, SysUtils, Classes, Graphics, Controls, Forms,
  StdCtrls, Buttons, ExtCtrls, DB, Variants, ComCtrls, ComObj, math{max函数} ;

type
  TfrmLYDataToExcel = class(TForm)
    ProgressBar1: TProgressBar;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Panel4: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton1: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton3: TSpeedButton;
    Label1: TLabel;
    Panel2: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    TabSheet2: TTabSheet;
    GroupBox1: TGroupBox;
    CheckBox_Pages: TCheckBox;
    CheckBox_PageCount: TCheckBox;
    CheckBox_Date: TCheckBox;
    CheckBox_user: TCheckBox;
    GroupBox3: TGroupBox;
    CheckBox_EdgesLines: TCheckBox;
    CheckBox_InLines: TCheckBox;
    TreeView1: TTreeView;
    TreeView2: TTreeView;
    Label3: TLabel;
    ComboBox1: TComboBox;
    RadioGroup1: TRadioGroup;
    Title_Edit: TLabeledEdit;
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure TreeView1DblClick(Sender: TObject);
    procedure TreeView2DblClick(Sender: TObject);
  private
    { Private declarations }
    prindataset : Tdataset;
    SopCaption  : string;
    procedure up_Move(MySource,MyTarget:TTreeView; MyIndex: Integer);
  public
    { Public declarations }
  end;

type
  TLYDataToExcel = class(TComponent)
  private
    { Private declarations }
    FDataSet:tdataset;
    FExcelTitel:STRING;
    ffrmLYDataToExcel: TfrmLYDataToExcel;
    procedure FSetExcelTitel(const value:string);
    procedure FSetDataSet(const value:tdataset);
  protected
    { Protected declarations }
  public
    { Public declarations }
    constructor create(aowner:tcomponent);override;
    destructor destroy;override;
    function Execute:boolean;
  published
    { Published declarations }
    property DataSet: tdataset read FDataSet write FSetDataSet;
    property ExcelTitel:string read FExcelTitel write FSetExcelTitel;
  end;

procedure Register;

implementation

var
  varDatas : Variant;
{$R *.DFM}

procedure Register;
begin
  RegisterComponents('Eagle_Ly', [TLYDataToExcel]);
end;

procedure TfrmLYDataToExcel.up_Move(MySource,MyTarget:TTreeView; MyIndex: Integer);
var
  i: Integer;
  MyNode: TTreeNode ;
begin
  if MySource.Items.Count=0 then Exit;
  if MySource.Selected=nil then MySource.Selected:=MySource.TopItem;
  if MyIndex=-1 then
  begin
    for i:=0 to MySource.Items.Count-1 do
    begin
       MyNode := MySource.Items[i];
       with MyTarget.Items.Add(nil,MyNode.Text) do
         StateIndex:=MyNode.StateIndex;
    end;
    MySource.Items.Clear;
  end
  else
  begin
    MyNode:=MySource.Selected;
    with MyTarget.Items.Add(nil,MyNode.Text) do
      StateIndex:=MyNode.StateIndex;

    MySource.Items[MySource.Selected.AbsoluteIndex].Delete;
  end;
end;

procedure TfrmLYDataToExcel.FormShow(Sender: TObject);
var
  i,intNum: Integer;
begin
  Title_Edit.Text := SopCaption;
  ComboBox1.ItemIndex:=1;
  RadioGroup1.ItemIndex:=0;
  PageControl1.ActivePageIndex:=0;
  TreeView1.SetFocus;

  if (not PrinDataSet.Active) or (PrinDataSet.RecordCount=0) then Exit;
  
  intNum:=0;
  for i:=0 to PrinDataSet.FieldCount-1  do
  begin
    if (PrinDataSet.Fields[i].Visible) then
    begin
      with TreeView1.Items.Add(nil,PrinDataSet.Fields[i].displaylabel) do
        StateIndex:=i;
      inc(intNum);
    end;
  end;
  varDatas:= varArrayCreate([1,500,1,intNum],varVariant); //这里创建动态数组
end;

procedure TfrmLYDataToExcel.SpeedButton1Click(Sender: TObject);
begin
  up_Move(Treeview1,Treeview2,0);
end;

procedure TfrmLYDataToExcel.SpeedButton2Click(Sender: TObject);
begin
  up_Move(Treeview1,Treeview2,-1);
end;

procedure TfrmLYDataToExcel.SpeedButton3Click(Sender: TObject);
begin
  up_Move(Treeview2,Treeview1,0);
end;

procedure TfrmLYDataToExcel.SpeedButton4Click(Sender: TObject);
begin
  up_Move(Treeview2,Treeview1,-1);
end;

procedure TfrmLYDataToExcel.BitBtn2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmLYDataToExcel.BitBtn1Click(Sender: TObject);
const
  TextRow = 2;
  TextCol = 2;
  xlHAlignLeft =-4131;
  xlHAlignRight =-4152;
  xlHAlignCenter =-4108;
  xlInsideHorizontal = 12;
  xlInsideVertical = 11;
  xlEdgeBottom = 9;
  xlEdgeLeft = 7;
  xlEdgeRight = 10;
  xlEdgeTop = 8;
  xlThin =2;
  xlThick =4;
  xlContinuous = 1;
  xlPaperA3 = $00000008;
  xlPaperA4 = $00000009;
  xlLandscape = $00000002;
  xlPortrait = $00000001;
var
  VExcelApp: Variant;
  VExcelWorkBook: Variant;
  VExcelWorkSheet1: Variant;
  I, VCellRow,ExcelColWidth: Integer;
  FieldNumber: Integer;
  S: String;
  FieldType1: TFieldType;
  CurrentRecordBookMark: TBookMark;

  intRow: Integer;
begin
  if TreeView2.Items.Count=0 then
  begin
    raise Exception.Create('请选择要打印的字段!');
    Exit;
  end;

  try
    VExcelApp := CreateOleObject('Excel.Application');
  except
    on E:Exception do
    begin
      raise Exception.Create('Execl异常:'+E.Message);
      exit;
    end;
  end;

    if (not PrinDataSet.Active) or (PrinDataSet.RecordCount=0) then Exit;
    
    ProgressBar1.Min := 0;
    ProgressBar1.Max := PrinDataSet.RecordCount+1;
    ProgressBar1.Step := 1;
    //ProgressBar1.Visible:=True;
    ProgressBar1.StepIt;

    PrinDataSet.DisableControls;
    CurrentRecordBookMark := PrinDataSet.GetBookmark;
    VExcelApp.Caption:='Microsoft Excel('+SopCaption+')';
    VExcelApp.Visible := False;
    VExcelApp.SheetsInNewWorkbook := 1;
    VExcelWorkBook := VExcelApp.WorkBooks.Add;
    VExcelWorkSheet1 := VExcelWorkBook.Sheets[1];

    VExcelWorkSheet1.Cells[1,1].Value := '序号';

    VExcelWorkSheet1.Columns[1].ColumnWidth := 6;
    FieldNumber:=0;
    for I := 0 to Treeview2.Items.Count -1 do
    begin
      FieldType1 := PrinDataSet.Fields[TreeView2.Items[I].StateIndex].DataType;

      //===================设置列宽===========================================//
      ExcelColWidth:=max(PrinDataSet.Fields[TreeView2.Items[I].StateIndex].DisplayWidth,
                         Length(PrinDataSet.Fields[TreeView2.Items[I].StateIndex].DisplayLabel));
      if ExcelColWidth>255 then ExcelColWidth:=255;
      VExcelWorkSheet1.Columns[TextCol+FieldNumber].ColumnWidth := ExcelColWidth;//当ExcelColWidth大于255时，设置将出错
      //======================================================================//

      VExcelWorkSheet1.Cells[1,TextCol+FieldNumber].Value := PrinDataSet.Fields[TreeView2.Items[I].StateIndex].DisplayLabel;
      case FieldType1 of
        ftBoolean    ,
        ftMemo       ,
        ftFmtMemo    ,
        ftWideString ,
        ftString     ,
        ftTime       ,
        ftDate       ,
        ftDateTime   : begin
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].NumberFormat := '@';
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].HorizontalAlignment := xlHAlignLeft;
                       end;
        ftFloat      : begin
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].NumberFormat := '0.00';
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].HorizontalAlignment := xlHAlignRight;
                       end;
        ftCurrency   : begin
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].NumberFormat := '￥#,##0.00;[红色]￥-#,##0.00';
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].HorizontalAlignment := xlHAlignRight;
                       end;
        ftSmallint   ,
        ftWord       ,
        ftLargeint   ,
        ftAutoInc    ,
        ftInteger    : begin
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].NumberFormat := '0_ ;[红色]-0';
                         VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow, TextCol+FieldNumber], VExcelWorkSheet1.Cells[TextRow+PrinDataSet.RecordCount-1, TextCol+FieldNumber]].HorizontalAlignment := xlHAlignRight;
                       end;
      end;
      FieldNumber:=FieldNumber+1;
    end;

    varDatas:= varArrayCreate([1,500,1,FieldNumber+1],varVariant); //这里创建动态数组
    intRow:=1;
    PrinDataSet.First;
    VCellRow := TextRow;
    while not PrinDataSet.Eof do
    begin
      varDatas[intRow,1]:=VCellRow-TextRow+1;
      FieldNumber:=0;
      for I := 0 to TreeView2.Items.Count-1 do
      begin
        with PrinDataSet.Fields[TreeView2.Items[I].StateIndex] do
        begin
          if DataType = ftBoolean then
          begin
            S := AsString;
            if S<>'' then if AsBoolean then S := '是' else S := '否';
          end
          else
          if Lookup then
          begin
            if LookupDataset.Locate(LookupKeyFields,PrinDataSet.FieldByName(KeyFields).AsString,[]) then
              S := LookupDataset.FieldByName(LookupResultField).AsString
            else S := '';
          end
          else S := AsString;
        end;
        varDatas[intRow,TextCol+FieldNumber]:=S;
        FieldNumber:=FieldNumber+1;
      end;
      PrinDataSet.Next;
      if (intRow=500) or (PrinDataSet.Eof) then
      begin
        VExcelWorkSheet1.Range[VExcelWorkSheet1.cells.Item[VCellRow-intRow+1,1],
                               VExcelWorkSheet1.cells.Item[VCellRow,FieldNumber+1]].Value:=varDatas;
        intRow:=0;
      end;
      intRow:=intRow+1;
      VCellRow := VCellRow+1;
      ProgressBar1.Stepit;
    end;

    VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[1, 1], VExcelWorkSheet1.Cells[1, FieldNumber+1]].HorizontalAlignment := xlHAlignCenter;
    VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[1, 1], VExcelWorkSheet1.Cells[1, FieldNumber+1]].Borders[xlEdgeRight].Weight := xlThick;

    VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Font.Size:= 9;
    VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Font.Name:= '宋体';

    if CheckBox_InLines.Checked then
       begin
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders.LineStyle := xlContinuous;
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders.Weight := xlThin;
       end;

    if CheckBox_EdgesLines.Checked then
       begin
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders[xlEdgeBottom].Weight := xlThick;
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders[xlEdgeLeft].Weight := xlThick;
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders[xlEdgeRight].Weight := xlThick;
        VExcelWorkSheet1.Range[VExcelWorkSheet1.Cells[TextRow-1, 1], VExcelWorkSheet1.Cells[VCellRow-1, TextCol+FieldNumber-1]].Borders[xlEdgeTop].Weight := xlThick;
       end;

    VExcelWorkSheet1.PageSetup.PrintTitleRows := VExcelWorkSheet1.Rows[1].Address;
    VExcelWorkSheet1.PageSetup.PrintTitleColumns := VExcelWorkSheet1.Columns[1].Address;
    VExcelWorkSheet1.PageSetup.CenterHeader := '&18&"黑体"'+title_edit.Text;

    if ComboBox1.ItemIndex=0 then VExcelWorkSheet1.PageSetup.PaperSize:=xlPaperA3
    else VExcelWorkSheet1.PageSetup.PaperSize:=xlPaperA4;

    if RadioGroup1.ItemIndex=0 then VExcelWorkSheet1.PageSetup.Orientation:=xlPortrait
    else VExcelWorkSheet1.PageSetup.Orientation:=xlLandscape;

    if CheckBox_Date.Checked then
      VExcelWorkSheet1.PageSetup.LeftFooter := '打印日期:&D';

    if (CheckBox_Pages.Checked) and (CheckBox_PageCount.checked) then
      VExcelWorkSheet1.PageSetup.CenterFooter := '第&P页  共&N页'
    else
    if (CheckBox_Pages.Checked) and (not CheckBox_PageCount.checked) then
      VExcelWorkSheet1.PageSetup.CenterFooter := '第&P页'
    else
    if (not CheckBox_Pages.Checked) and (CheckBox_PageCount.checked) then
      VExcelWorkSheet1.PageSetup.CenterFooter := '共&N页';

    if CheckBox_user.Checked then
       VExcelWorkSheet1.PageSetup.rightFooter := '制表人：';//+Operator.Name;

    //ProgressBar1.Visible := False;

    PrinDataSet.GotoBookmark(CurrentRecordBookmark);
    PrinDataSet.EnableControls;
    PrinDataSet.FreeBookMark(CurrentRecordBookmark);
    VExcelApp.Visible := True;
end;

{ TLYDataToExcel }

constructor TLYDataToExcel.create(aowner: tcomponent);
begin
  inherited Create(AOwner);
end;

destructor TLYDataToExcel.destroy;
begin
  inherited Destroy;
end;

function TLYDataToExcel.Execute: boolean;
begin
  if FDataSet=nil then
  begin
    raise Exception.Create('没有设置DataSet属性!');
    result:=false;
    exit;
  end;
  if not FDataSet.Active then
  begin
    raise Exception.Create('数据集没有打开!');
    result:=false;
    exit;
  end;
  if FDataSet.RecordCount=0 then
  begin
    raise Exception.Create('数据集记录行数为零!');
    result:=false;
    exit;
  end;
  ffrmLYDataToExcel:=tfrmLYDataToExcel.Create(nil);
  ffrmLYDataToExcel.prindataset:=FDataSet;
  ffrmLYDataToExcel.SopCaption:=FExcelTitel;
  ffrmLYDataToExcel.ShowModal;
  result:=true;
  ffrmLYDataToExcel.Free;
end;

procedure TLYDataToExcel.FSetDataSet(const value: tdataset);
begin
  //if value=FInField then exit;
  FDataSet:=value;
end;

procedure TLYDataToExcel.FSetExcelTitel(const value: string);
begin
  if value=FExcelTitel then exit;
  FExcelTitel:=value;
end;

procedure TfrmLYDataToExcel.TreeView1DblClick(Sender: TObject);
begin
  SpeedButton1Click(nil);
end;

procedure TfrmLYDataToExcel.TreeView2DblClick(Sender: TObject);
begin
  SpeedButton3Click(nil);
end;

end.
