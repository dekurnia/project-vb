unit Unit5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, DBGrids, ExtCtrls;

type
  TfrmItem = class(TForm)
    Label1: TLabel;
    Shape1: TShape;
    Panel1: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    MKet: TMemo;
    EStyle: TEdit;
    EItem: TEdit;
    EStock: TEdit;
    DBGrid1: TDBGrid;
    BNew: TBitBtn;
    BSimpan: TBitBtn;
    BEdit: TBitBtn;
    BUpdate: TBitBtn;
    BHapus: TBitBtn;
    BRefresh: TBitBtn;
    BKeluar: TBitBtn;
    procedure tampil;
    procedure awal;
    procedure aktif;
    procedure nonaktif;
    procedure BSimpanClick(Sender: TObject);
    procedure BHapusClick(Sender: TObject);
    procedure BRefreshClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BKeluarClick(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure BNewClick(Sender: TObject);
    procedure BEditClick(Sender: TObject);
    procedure BUpdateClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmItem: TfrmItem;

implementation
Uses Unit1;
{$R *.dfm}

procedure TfrmItem.Awal;
begin
  EStyle.Clear;
  EItem.Clear;
  EStock.clear;
  MKet.Clear;
  BSimpan.Enabled:=False;
  BEdit.Enabled:=True;
  BHapus.Enabled:=True;
  BUpdate.Enabled:=False;
  BRefresh.Enabled:=True;
  BNew.Enabled:=True;
  nonaktif;
end;

procedure TfrmItem.aktif;
begin
  EStyle.Enabled := True;
  EItem.Enabled := True;
  EStock.Enabled := True;
  MKet.Enabled := True;
end;

procedure TfrmItem.nonaktif;
begin
  EStyle.Enabled := False;
  EItem.Enabled := False;
  EStock.Enabled := False;
  MKet.Enabled := False;
end;

procedure TfrmItem.tampil;
begin
    with frmMenu.QItem do
    begin
      SQL.Clear;
      SQL.Add('select * from tb_item');
      Open;
    end;
end;

procedure TfrmItem.BSimpanClick(Sender: TObject);
begin
  if (EStyle.Text='') or (EItem.Text='') then
    begin
      MessageDlg('Data tidak lengkap...',mtError,[mbOK],0);
      Exit;
    end
  else
  with frmMenu.QItem do
    begin
      SQL.Clear;
      SQL.Add('select * from tb_item where style="'+EStyle.Text+'"');
      Open;
      if frmMenu.QItem.Eof then
        begin
          if MessageDlg('Yakin ingin menyimpan data..?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
            begin

              SQL.Clear;
              SQL.Add('insert into tb_item (style,item,stock,ket) values ("'+EStyle.Text+'","'+EItem.Text+'","'+EStock.Text+'","'+MKet.Text+'")');
              frmMenu.QItem.ExecSQL;
              ShowMessage('Data berhasil di simpan...');
              EStyle.Clear;
              EStyle.SetFocus;
              Awal;
              tampil;
        end
      else
        begin
          MessageDlg('Item Sudah Ada',mtError,[mbOK],0);
          EStyle.Clear;
          EStyle.SetFocus;
        end;
    end;
    end;
end;

procedure TfrmItem.BHapusClick(Sender: TObject);
begin
  if (EStyle.Text='') then
    begin
      MessageDlg('Silahkan Pilih Data Terlebih Dahulu',mtWarning,[mbOK],0);
      Exit;
    end
  else
    begin
      frmMenu.QItem.SQL.Clear;
      frmMenu.QItem.SQL.Add('select * from tb_item where style="'+EStyle.Text+'"');
      frmMenu.QItem.Open;
      if not frmMenu.QItem.Eof then
        begin
          if MessageDlg('Yakin ingin menghapus Data Item..?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
            begin
              frmMenu.QItem.SQL.Clear;
              frmMenu.QItem.SQL.Add('delete from tb_item where style="'+EStyle.Text+'"');
              frmMenu.QItem.ExecSQL;
              Awal;
              tampil;
              nonaktif;
              end;
          end
        else
          begin
            MessageDlg('Item Tidak Ada',mtWarning,[mbOK],0);
            Awal;
            tampil;
            nonaktif;
          end;
   end;
end;

procedure TfrmItem.BRefreshClick(Sender: TObject);
begin
  awal;
  tampil;
  nonaktif;
end;

procedure TfrmItem.FormCreate(Sender: TObject);
begin
 awal;
 tampil;
end;

procedure TfrmItem.BKeluarClick(Sender: TObject);
begin
 frmItem.Close;
end;

procedure TfrmItem.DBGrid1CellClick(Column: TColumn);
begin
 EStyle.Text := DBGrid1.Fields[0].AsString;
 EItem.Text := DBGrid1.Fields[1].AsString;
 EStock.Text := DBGrid1.Fields[2].AsString;
 MKet.Text := DBGrid1.Fields[3].AsString;
end;

procedure TfrmItem.BNewClick(Sender: TObject);
begin
  aktif;
  Bsimpan.Enabled := True;
  BEdit.Enabled := False;
  BUpdate.Enabled := False;
  BHapus.Enabled := False;
  EStyle.SetFocus;
end;

procedure TfrmItem.BEditClick(Sender: TObject);
begin
  if (EStyle.Text='') or (EItem.Text = '') then
    begin
      MessageDlg('Data Masih Kosong !',mtWarning,[mbOK],0);
    end
  else
    begin
      aktif;
      BUpdate.Enabled := True;
      BNew.Enabled := False;
      BSimpan.Enabled := False;
      BHapus.Enabled := False;
    end;
  end;

procedure TfrmItem.BUpdateClick(Sender: TObject);
begin
  if (EStyle.Text='') or (EItem.Text = '') then
    begin
      MessageDlg('Data Masih Kosong !',mtWarning,[mbOK],0);
      EStyle.SetFocus;
    end
  else
    begin
      frmMenu.QItem.SQL.Clear;
      frmMenu.QItem.SQL.Add('select * from tb_item where style="'+EStyle.Text+'"');
      frmMenu.QItem.Open;
      if not frmMenu.QItem.Eof then
            begin
               frmMenu.QItem.SQL.Clear;
               frmMenu.QItem.SQL.Add('update tb_item set item="'+Eitem.Text+'", stock="'+EStock.Text+'", ket="'+Mket.Text+'" where style ="'+EStyle.Text+'"');
               frmMenu.QItem.ExecSQL;
               ShowMessage('Data Berhasil diubah...');
               Awal;
               tampil;
            end
      else
        begin
          MessageDlg('Buyer Tidak Terdaftar',mtError,[mbOK],0);
          Awal;
          tampil;
        end;
    end;
end;

end.
