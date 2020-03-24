object FmTableProperty: TFmTableProperty
  Left = 348
  Top = 326
  BorderStyle = bsDialog
  Caption = #34920#26684#23646#24615#35774#32622
  ClientHeight = 366
  ClientWidth = 327
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 161
    Height = 229
    Caption = #34920#26684
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 48
      Width = 24
      Height = 13
      Caption = #34892#25968
    end
    object Label2: TLabel
      Left = 16
      Top = 95
      Width = 24
      Height = 13
      Caption = #21015#25968
    end
    object Label3: TLabel
      Left = 16
      Top = 139
      Width = 24
      Height = 13
      Caption = #23485#24230
    end
    object Label4: TLabel
      Left = 16
      Top = 183
      Width = 24
      Height = 13
      Caption = #36793#23485
    end
    object Label8: TLabel
      Left = 16
      Top = 17
      Width = 24
      Height = 13
      Caption = #26631#39064
      Visible = False
    end
    object SpinEdit1: TJvSpinEdit
      Left = 16
      Top = 64
      Width = 121
      Height = 21
      Value = 4.000000000000000000
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 0
    end
    object SpinEdit2: TJvSpinEdit
      Left = 16
      Top = 111
      Width = 121
      Height = 21
      Value = 4.000000000000000000
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 1
    end
    object SpinEdit3: TJvSpinEdit
      Tag = 100
      Left = 16
      Top = 155
      Width = 121
      Height = 21
      Value = 100.000000000000000000
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 2
    end
    object SpinEdit4: TJvSpinEdit
      Tag = 1
      Left = 16
      Top = 199
      Width = 121
      Height = 21
      Value = 1.000000000000000000
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 3
    end
    object Edit1: TEdit
      Left = 16
      Top = 32
      Width = 121
      Height = 21
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 4
      Visible = False
    end
  end
  object GroupBox2: TGroupBox
    Left = 176
    Top = 8
    Width = 145
    Height = 229
    Caption = #21333#20803#26684
    TabOrder = 1
    object Label5: TLabel
      Left = 16
      Top = 72
      Width = 36
      Height = 13
      Caption = #32972#26223#33394
    end
    object CheckBox1: TCheckBox
      Left = 16
      Top = 32
      Width = 89
      Height = 17
      Caption = #36879#26126
      Checked = True
      State = cbChecked
      TabOrder = 0
      OnClick = CheckBox1Click
    end
    object ColorComboBox1: TJvColorComboBox
      Left = 16
      Top = 88
      Width = 113
      Height = 20
      ColorDialogText = 'Custom...'
      DroppedDownWidth = 113
      NewColorText = 'Custom'
      Enabled = False
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 1
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 242
    Width = 313
    Height = 73
    Caption = #36793#36317#21644#38388#36317
    TabOrder = 2
    object Label6: TLabel
      Left = 24
      Top = 44
      Width = 24
      Height = 13
      Caption = #36793#36317
    end
    object Label7: TLabel
      Left = 168
      Top = 44
      Width = 24
      Height = 13
      Caption = #38388#36317
    end
    object SpinEdit5: TJvSpinEdit
      Left = 48
      Top = 36
      Width = 73
      Height = 21
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 0
    end
    object SpinEdit6: TJvSpinEdit
      Left = 200
      Top = 36
      Width = 73
      Height = 21
      ImeName = #35895#27468#25340#38899#36755#20837#27861' 2'
      TabOrder = 1
    end
  end
  object RzBitBtn1: TBitBtn
    Left = 152
    Top = 330
    Width = 75
    Height = 25
    Caption = #30830#23450
    Default = True
    ModalResult = 1
    TabOrder = 3
  end
  object RzBitBtn2: TBitBtn
    Left = 240
    Top = 330
    Width = 75
    Height = 25
    Cancel = True
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 4
  end
end
