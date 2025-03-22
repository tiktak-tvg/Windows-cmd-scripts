[Setup]
AppName=My Program
AppVerName=My Program v 1.5
DefaultDirName={pf}\My Program
OutputDir=.
Compression=lzma/ultra
InternalCompressLevel=ultra
SolidCompression=yes

[Languages]
Name: rus; MessagesFile: compiler:Languages\Russian.isl

[Code]
var
  text: TLabel;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  i, n: Integer;
  str: string;
begin
  Result:= True;
  if CurPageID = wpSelectDir then
    begin
      str:= WizardForm.DirEdit.Text;
      for i:= 1 to Length(str) do if str[i] = '\' then n:= n + 1;
      if (n = 1) and (Pos(ExpandConstant('{sd}'), WizardForm.DirEdit.Text) > 0) then
        begin
          text.Caption:= 'Внимание, в корень системного диска, устанавливать нельзя.';
          Result:= False;
        end
      else text.Caption:= '';
    end;
end;

procedure InitializeWizard();
begin
  text:= TLabel.Create(WizardForm);
  with text do
    begin
      SetBounds(WizardForm.DirEdit.Left,120,300,100);
      AutoSize:= True;
      Font.Style:= [fsBold];
      Font.Color:= clRed;
      Parent:= WizardForm.SelectDirPage;
    end;
end;