# Provjera prije obfuskacije (pokreće se automatski u Release buildu).
#
# Kad se 'dynamic' vrijednost (npr. Excel sheet) preda NAŠOJ metodi, .NET tu metodu traži po imenu
# tek dok program radi. Obfuscar mijenja imena metoda, pa takav poziv u Release verziji pukne
# ("... does not contain a definition for ...").
# Ova skripta traži imena naših metoda zapisana kao tekst u exe-u i prekida build ako ih nađe.
# Rješenje u kodu: argument proslijediti kao (object), npr. buffer.Flush((object)sheet).

param([Parameter(Mandatory = $true)][string]$Exe)

$asm = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($Exe))
$types = try { $asm.GetTypes() } catch [Reflection.ReflectionTypeLoadException] { $_.Exception.Types | Where-Object { $_ } }
$flags = [Reflection.BindingFlags]'Public,NonPublic,Instance,Static,DeclaredOnly'

$methods = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($t in $types) {
    if ($t.Namespace -notlike 'FASTIMATE*') { continue }
    foreach ($m in $t.GetMethods($flags)) {
        if ($m.Name -match '^[<.]|^(get_|set_|add_|remove_)' -or $m.Name.Length -lt 4) { continue }
        [void]$methods.Add($m.Name)
    }
}

# Imena koja su legitimno tekst u kodu (npr. imena COM metoda Excela/SolidWorksa)
$ignore = @('Save', 'Close', 'Quit', 'Delete', 'Activate', 'Calculate', 'Value')

$bytes = [IO.File]::ReadAllBytes($Exe)
$texts = @([Text.Encoding]::Unicode.GetString($bytes), [Text.Encoding]::Unicode.GetString($bytes, 1, $bytes.Length - 1))

$found = @()
foreach ($n in $methods) {
    if ($ignore -contains $n) { continue }
    $re = '(?<![A-Za-z0-9_])' + [regex]::Escape($n) + '(?![A-Za-z0-9_])'
    foreach ($tx in $texts) { if ([regex]::IsMatch($tx, $re)) { $found += $n; break } }
}

if ($found.Count -gt 0) {
    Write-Output ("ProvjeraObfuskacije: error : Dinamicki pozivi vlastitih metoda (pukli bi nakon obfuskacije): " + (($found | Sort-Object) -join ', ') + ". Argument proslijediti kao (object).")
    exit 1
}
Write-Output "ProvjeraObfuskacije: OK (nema dinamickih poziva vlastitih metoda)"
exit 0
