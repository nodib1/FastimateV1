# FASTIMATE

Samostalna WPF aplikacija (.NET Framework 4.8) koja čita podatke iz otvorenog SolidWorks sklopa
i upisuje proračun vremena (Ponuda / Tehnologija) u Excel. Makroi se više ne koriste.

## Struktura

```
FASTIMATE.sln
FASTIMATE/
  *.xaml, *.xaml.cs            prozori i ekrani aplikacije
  TehnologijaProcessor.cs      proračun tehnologije (zamjena za Macro_Tehnologija.swp)
  PonudaProcessor.cs           proračun ponude
  *DbManager.cs, KFactorManager.cs   baze (C:\ProgramData\FASTIMATE)
  Resources/                   Excel šabloni (ugrađeni u exe)
  Pomoc_FASTIMATE.pdf          uputstvo (ugrađeno u exe)
  PdfSharp.dll                 biblioteka za PDF
  obfuscar.xml                 podešavanja zaštite koda (Release)
  Fastimate_Installer_Release.iss   Inno Setup skripta za instalaciju
```

## Build

- **Debug**: običan build u Visual Studiju, bez obfuscacije (za razvoj i debug).
- **Release**: nakon builda se automatski pokreće Obfuscar (NuGet paket) i u `bin\Release`
  ostaje zaštićeni `FASTIMATE.exe`. Mapa preimenovanja (za čitanje grešaka iz Release verzije)
  je u `FASTIMATE\obj\Obfuscated\Mapping.txt` – ne dijeliti je.

Potrebno na računaru za build: Visual Studio 2022 i instaliran SolidWorks
(interop DLL-ovi iz `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\`).

## Instalacija

Nakon Release builda otvoriti `FASTIMATE\Fastimate_Installer_Release.iss` u Inno Setupu i kompajlirati.
