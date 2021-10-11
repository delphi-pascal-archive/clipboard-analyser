unit f_prin;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DB, DBClient, Clipbrd,
  Buttons, ShellAPI;

type
  TfrmPrin = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    MemClipBrd: TMemo;
    Splitter1: TSplitter;
    DBGrid1: TDBGrid;
    ClientDSClipBrd: TClientDataSet;
    DSClientDSClipBrd: TDataSource;
    ClientDSClipBrdNUM: TIntegerField;
    ClientDSClipBrdFORMAT: TIntegerField;
    ClientDSClipBrdStrFormat: TStringField;
    SBAnalyzeClipbrd1: TSpeedButton;
    SBAnalyzeClipbrd2: TSpeedButton;
    ScrollBox1: TScrollBox;
    ImgClipBrd: TImage;
    Panel4: TPanel;
    SBGetClipBrd: TSpeedButton;
    procedure SBAnalyzeClipbrd1Click(Sender: TObject);
    procedure SBGetClipBrdClick(Sender: TObject);
    procedure SBAnalyzeClipbrd2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function GetclipboardStrFormat(Format: Word): String;
  end;

var
  frmPrin: TfrmPrin;

implementation

{$R *.dfm}

function TfrmPrin.GetclipboardStrFormat(Format: Word): String;
var buf: Array [0..99] of Char;
begin
// This part comes from http://www.delphipages.com/threads/thread.cfm?ID=172640&G=172339 //
  if Windows.GetclipboardFormatName(Format, buf, Pred(Sizeof(buf))) <> 0
  then
    RESULT := StrPas(buf)
  else
    case Format of
            1: RESULT := 'CF_TEXT';
            2: RESULT := 'CF_BITMAP';
            3: RESULT := 'CF_METAFILEPICT';
            4: RESULT := 'CF_SYLK';
            5: RESULT := 'CF_DIF';
            6: RESULT := 'CF_TIFF';
            7: RESULT := 'CF_OEMTEXT';
            8: RESULT := 'CF_DIB';
            9: RESULT := 'CF_PALETTE';
           10: RESULT := 'CF_PENDATA';
           11: RESULT := 'CF_RIFF';
           12: RESULT := 'CF_WAVE';
           13: RESULT := 'CF_UNICODETEXT';
           14: RESULT := 'CF_ENHMETAFILE';
           15: RESULT := 'CF_HDROP (Win 95)';
           16: RESULT := 'CF_LOCALE (Win 95)';
           17: RESULT := 'CF_MAX (Win 95)';
        $0080: RESULT := 'CF_OWNERDISPLAY';
        $0081: RESULT := 'CF_DSPTEXT';
        $0082: RESULT := 'CF_DSPBITMAP';
        $0083: RESULT := 'CF_DSPMETAFILEPICT';
        $008E: RESULT := 'CF_DSPENHMETAFILE';
 $0200..$02FF: RESULT := 'private format';
 $0300..$03FF: RESULT := 'GDI object';
        else
               RESULT := 'unknown format';
    end;
// This part comes from http://www.delphipages.com/threads/thread.cfm?ID=172640&G=172339 //
end;

procedure TfrmPrin.SBAnalyzeClipbrd1Click(Sender: TObject);
var
  Num, Format: Word;
  StrFormat: String;
begin
  if ClientDSClipBrd.Active
  then ClientDSClipBrd.Close;

  ClientDSClipBrd.CreateDataSet;

  with Clipboard do
  begin
    Open;
    Num := 0;
    try
      Format := EnumClipboardFormats(0);   // Browse all clipboard formats from the first one
      while Format <> 0 do
      begin
        StrFormat := GetclipboardStrFormat(Format);

        ClientDSClipBrd.Append;
        ClientDSClipBrdNUM.Value       := Num;
        ClientDSClipBrdFORMAT.Value    := Format;
        ClientDSClipBrdStrFormat.Value := StrFormat;
        ClientDSClipBrd.Post;

        inc(Num, 1);
        Format := EnumClipboardFormats(Format);  // Next object
      end;
    finally
      Close;
    end;
  end;
end;

procedure TfrmPrin.SBAnalyzeClipbrd2Click(Sender: TObject);
var
  f, Format: Word;
  StrFormat: String;
begin
  if ClientDSClipBrd.Active
  then ClientDSClipBrd.Close;

  ClientDSClipBrd.CreateDataSet;

  with Clipboard do
  begin
    Open;
    try
      for f := 0 to FormatCount-1 do
      begin
        Format := Formats[f];
        StrFormat := GetclipboardStrFormat(Format);

        ClientDSClipBrd.Append;
        ClientDSClipBrdNUM.Value       := f;
        ClientDSClipBrdFORMAT.Value    := Format;
        ClientDSClipBrdStrFormat.Value := StrFormat;
        ClientDSClipBrd.Post;
      end;
    finally
      Close;
    end;
  end;
end;

procedure TfrmPrin.SBGetClipBrdClick(Sender: TObject);
var
  Format: Word;

      procedure GetAsPictureFormat(aFormat: Word);
      var
        FHandle: THandle;
        FPalette: HPalette;
        aPicture: TPicture;
      begin
        FHandle  := GetClipboardData(aFormat);
        FPalette := GetClipboardData(CF_PALETTE);
        aPicture := TPicture.Create;
        aPicture.LoadFromClipboardFormat(aFormat, FHandle, FPalette);
        ImgClipBrd.Picture.Assign(aPicture);
        aPicture.Free;
      end;

      procedure GetAsText;
      var
        DataHandle: THandle;
        DataString: String;
        DataPtr: Pointer;
      begin
        DataHandle := GetClipboardData(Format);
        if DataHandle <> 0
        then begin
          DataPtr := GlobalLock(DataHandle);
          DataString := PChar(DataPtr);
          MemClipBrd.Lines.Text := DataString;
          GlobalUnLock(DataHandle);
          MemClipBrd.Lines.Text := DataString;
        end;
      end;

      procedure GetUnicodeText;
      var
        DataHandle: THandle;
        DataString: WideString;
        DataPtr: PWideChar;
      begin
        DataHandle := GetClipboardData(Format);
        if DataHandle <> 0
        then begin
          DataPtr := GlobalLock(DataHandle);
          DataString := DataPtr;
          MemClipBrd.Lines.Text := DataString;
          GlobalUnLock(DataHandle);
          MemClipBrd.Lines.Text := DataString;
        end;
      end;

      procedure GetWindowsFiles;
      var
        DataHandle: THandle;
        DataPtr: Pointer;
        Buffer : array[0..MAX_PATH] of Char;
        Count, i: Integer;
      begin
        DataHandle := GetClipboardData(CF_HDROP);

        if DataHandle <> 0
        then begin
          DataPtr := PChar(GlobalLock(DataHandle));
          Count := DragQueryFile(HDROP(DataPtr), $FFFFFFFF, nil, 0);

          for i:=0 to Count-1 do
          begin
            DragQueryFile(HDROP(DataPtr), i, @Buffer, SizeOf(Buffer));

            if strlen(PChar(@Buffer)) > 0
            then MemClipBrd.Lines.Add(PChar(@Buffer));
          end;
          GlobalUnlock(DataHandle);
        end;
      end;

begin
  MemClipBrd.Lines.Clear;
  ImgClipBrd.Picture := nil;

  if ClientDSClipBrd.Bof and ClientDSClipBrd.Eof
  then EXIT;

  with Clipboard do
  begin
    Open;         
    try
      Format := ClientDSClipBrdFORMAT.AsInteger; 

      if Format = CF_UNICODETEXT
      then GetUnicodeText
      else
        if Format = CF_HDROP         // Windows directories and filenames
        then GetWindowsFiles
        else
          if TPicture.SupportsClipboardFormat(Format)
          then GetAsPictureFormat(Format)    // try to retrieve as TPicture supported image format
          else GetAsText;
    finally
      Close;
    end;
  end;
end;

end.
