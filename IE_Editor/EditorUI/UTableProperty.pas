unit UTableProperty;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask,
  Vcl.Buttons, JvExStdCtrls, JvCombobox, JvColorCombo, JvExMask, JvSpin;

type
  TFmTableProperty = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    SpinEdit1: TJvSpinEdit;
    SpinEdit2: TJvSpinEdit;
    SpinEdit3: TJvSpinEdit;
    SpinEdit4: TJvSpinEdit;
    Label5: TLabel;
    CheckBox1: TCheckBox;
    ColorComboBox1: TJvColorComboBox;
    Label6: TLabel;
    Label7: TLabel;
    SpinEdit5: TJvSpinEdit;
    SpinEdit6: TJvSpinEdit;
    RzBitBtn1: TBitBtn;
    RzBitBtn2: TBitBtn;
    Label8: TLabel;
    Edit1: TEdit;
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure GetParams(var Rows,Cols,Width,BorderWidth,Padding,SPacing: Integer;
       var BgColor: TColor; var BiaoTi: WideString);
  end;

var
  FmTableProperty: TFmTableProperty;

implementation

{$R *.dfm}

{ TFmTableProperty }

procedure TFmTableProperty.GetParams(var Rows, Cols, Width, BorderWidth,
  Padding, SPacing: Integer; var BgColor: TColor; var BiaoTi: WideString);
begin
  Rows := Trunc(SpinEdit1.Value);
  Cols := Trunc(SpinEdit2.Value);;
  Width := Trunc(SpinEdit3.Value);;
  BorderWidth := Trunc(SpinEdit4.Value);
  Padding := Trunc(SpinEdit5.Value);
  Spacing := Trunc(SpinEdit6.Value);
  if CheckBox1.Checked then BgColor := clWhite
  else BgColor := ColorComboBox1.ColorValue; //   .SelectedColor;

  BiaoTi := Edit1.Text;
end;

procedure TFmTableProperty.CheckBox1Click(Sender: TObject);
begin
  ColorComboBox1.Enabled := not CheckBox1.Checked; 
end;

end.
