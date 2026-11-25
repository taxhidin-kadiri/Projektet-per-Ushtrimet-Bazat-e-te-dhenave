# =====================================================================
# P01_Zgjidhja.ps1  -  "Bazat e të dhënave", Klasa 12, Projekti P01
# Biblioteka e shkollës: katalogu i librave dhe kalimi nga lista e
# rrafshët te tabelat e lidhura
#
# Qëllimi : krijon C:\BazatDB\P01\P01_Zgjidhja.accdb me tabelat, vetitë
#           e fushave, relacionet, të dhënat shembull dhe pyetësorët e
#           zgjidhur të P01. Mësimdhënësi e kopjon bazën me USB në 8
#           kompjuterët e kabinetit. Nxënësit nuk e ekzekutojnë skriptin.
# Temat   : N01, N04, N05, 1.2, N03 (.xlsx, fleta «Plani i orëve»,
#           Prioriteti = Thelbësore)
# Kërkesat: Windows PowerShell 5.1; Microsoft Access me ACE OLEDB
#           (16.0 ose 12.0) dhe DAO.DBEngine.120, me të njëjtin bitness
#           si PowerShell-i që e ekzekuton skriptin.
#           Pa Access.Application, pa VBE, pa SendKeys.
# Ruaje skedarin si "UTF-8 with BOM" që shkronjat ë, ç, Ë, Ç të mos prishen.
#
# Ekzekutimi, Office 64-bit:
#   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P01_Zgjidhja.ps1"
# Ekzekutimi, Office 32-bit:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P01_Zgjidhja.ps1"
#
# Rendi: dosja -> kopja rezervë -> baza -> tabelat -> vetitë e fushave ->
#        relacionet -> të dhënat (prindërit para fëmijëve) -> pyetësorët ->
#        verifikimi. Kodet e daljes: 0 = në rregull, 10-18 = gabim në
#        bllokun 0-8, 99 = gabim i papritur.
# =====================================================================

$ErrorActionPreference = 'Stop'

$Folder   = 'C:\BazatDB\P01'
$DbPath   = Join-Path $Folder 'P01_Zgjidhja.accdb'
$LockPath = Join-Path $Folder 'P01_Zgjidhja.laccdb'

$script:KodiDaljes    = 0
$script:Mosperputhje  = 0
$script:Paralajmerime = New-Object System.Collections.ArrayList
$Provider = $null
$Cat      = $null
$Conn     = $null
$DbE      = $null
$Db       = $null

# Konstantet DAO me vlerat numerike (PowerShell nuk i njeh emrat dbText etj.)
$dbBoolean        = 1
$dbByte           = 2
$dbInteger        = 3
$dbCurrency       = 5
$dbText           = 10
$dbMemo           = 12
$dbAttachment     = 101
$dbHyperlinkField = 32768
$dbFailOnError    = 128
$dbOpenSnapshot   = 4
$acCheckBox       = 106

# ---------------------------------------------------------------------
# Funksionet ndihmëse
# ---------------------------------------------------------------------

# Liron një objekt COM (nuk jep gabim nëse objekti është $null)
function Release-ComObjekt($Obj) {
    if ($null -ne $Obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($Obj)) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
    }
}

# Kthen mesazhin e gabimit më të brendshëm (aty është shkaku i vërtetë)
function Get-MesazhiGabimit($Gabimi) {
    $e = $null
    if ($Gabimi -is [System.Management.Automation.ErrorRecord]) { $e = $Gabimi.Exception }
    elseif ($Gabimi -is [System.Exception]) { $e = $Gabimi }
    if ($null -eq $e) { return [string]$Gabimi }
    while ($null -ne $e.InnerException) { $e = $e.InnerException }
    return $e.Message
}

# Kthen kodin HRESULT të gabimit më të brendshëm
function Get-HResult($Gabimi) {
    $e = $Gabimi.Exception
    while ($null -ne $e.InnerException) { $e = $e.InnerException }
    return $e.HResult
}

# Shfaq "GABIM në BLLOKU n" dhe e ndalon skriptin me kod jo-zero
function Ndalo([int]$Nr, [string]$Pershkrimi, $Gabimi) {
    Write-Host ''
    Write-Host ('GABIM në BLLOKU ' + $Nr + ': ' + $Pershkrimi) -ForegroundColor Red
    if ($null -ne $Gabimi) { Write-Host ('Hollësi: ' + (Get-MesazhiGabimit $Gabimi)) -ForegroundColor Red }
    $script:KodiDaljes = 10 + $Nr
    throw (New-Object System.OperationCanceledException ('Ndalur në bllokun ' + $Nr))
}

# Regjistron një paralajmërim (skripti vazhdon)
function Add-Paralajmerim([string]$Teksti) {
    [void]$script:Paralajmerime.Add($Teksti)
    Write-Host ('KUJDES: ' + $Teksti) -ForegroundColor Yellow
}

# Ekzekuton një deklaratë SQL me ADODB (vetëm CREATE TABLE, BLLOKU 3)
function Invoke-Sql([string]$Sql) { [void]$Conn.Execute($Sql) }

# Kontrollon me ADODB nëse ekziston tabela (TABLE_TYPE = 'TABLE')
function Test-AdoTable([string]$Emri) {
    $rs = $Conn.OpenSchema(20)          # 20 = adSchemaTables
    $gjet = $false
    try {
        while (-not $rs.EOF) {
            if ($rs.Fields.Item('TABLE_NAME').Value -eq $Emri -and $rs.Fields.Item('TABLE_TYPE').Value -eq 'TABLE') { $gjet = $true; break }
            $rs.MoveNext()
        }
    } finally { $rs.Close(); Release-ComObjekt $rs }
    return $gjet
}

# Krijon tabelën me DDL (ADODB). Baza është e re, prandaj tabela nuk duhet të ekzistojë.
function New-AccTable([string]$Emri, [string]$Ddl) {
    if (Test-AdoTable $Emri) { throw ('Tabela ' + $Emri + ' ekziston tashmë në bazën e re.') }
    Invoke-Sql $Ddl
    Write-Host ('U krijua tabela ' + $Emri)
}

# Kontrollon me DAO nëse ekziston tabela ('TABLE') ose pyetësori ('QUERY')
function Test-AccObject([string]$Emri, [string]$Lloji) {
    if ($Lloji -eq 'TABLE') { $koleksioni = $Db.TableDefs } else { $koleksioni = $Db.QueryDefs }
    $koleksioni.Refresh()
    $gjet = $false
    for ($i = 0; $i -lt $koleksioni.Count; $i++) {
        $o = $koleksioni.Item($i)
        if ($o.Name -eq $Emri) { $gjet = $true }
        Release-ComObjekt $o
        if ($gjet) { break }
    }
    return $gjet
}

# Kontrollon me DAO nëse tabela ka fushën e dhënë
function Test-Fusha([string]$Tabela, [string]$Fusha) {
    $tdf = $Db.TableDefs.Item($Tabela)
    $gjet = $false
    try {
        for ($i = 0; $i -lt $tdf.Fields.Count; $i++) {
            $f = $tdf.Fields.Item($i)
            if ($f.Name -eq $Fusha) { $gjet = $true }
            Release-ComObjekt $f
            if ($gjet) { break }
        }
    } finally { Release-ComObjekt $tdf }
    return $gjet
}

# Shton me DAO një fushë që DDL-ja nuk e krijon (Hyperlink, Attachment, Calculated)
function Add-DaoField([string]$Tabela, [string]$Fusha, [int]$Tipi, [string]$Shprehja, [int]$Atributet) {
    $tdf = $Db.TableDefs.Item($Tabela)
    $fld = $tdf.CreateField($Fusha, $Tipi)
    try {
        if ($Atributet -ne 0) { $fld.Attributes = $Atributet }
        if ($Shprehja -ne '') { $fld.Expression = $Shprehja }
        $tdf.Fields.Append($fld)
    } finally {
        Release-ComObjekt $fld
        Release-ComObjekt $tdf
    }
    Write-Host ('U shtua fusha ' + $Tabela + '.' + $Fusha)
}

# Vendos me DAO një veti të fushës; nëse vetia nuk ekziston, e krijon
function Set-FieldProp([string]$Tabela, [string]$Fusha, [string]$Emri, [int]$Tipi, $Vlera) {
    $tdf = $Db.TableDefs.Item($Tabela)
    $fld = $tdf.Fields.Item($Fusha)
    try {
        $ekziston = $true
        $p = $null
        try { $p = $fld.Properties.Item($Emri) } catch { $ekziston = $false }
        if ($ekziston) {
            $p.Value = $Vlera
        } else {
            $p = $fld.CreateProperty($Emri, $Tipi, $Vlera)
            $fld.Properties.Append($p)
        }
        Release-ComObjekt $p
    } finally {
        Release-ComObjekt $fld
        Release-ComObjekt $tdf
    }
    Write-Host ('U vendos vetia ' + $Tabela + '.' + $Fusha + ': ' + $Emri + ' = ' + $Vlera)
}

# Krijon me DAO një relacion me një fushë. Atributet 0 = Enforce Referential Integrity, pa Cascade.
function New-AccRelation([string]$Emri, [string]$Tabela, [string]$TabelaHuaj, [string]$Fusha, [string]$FushaHuaj, [int]$Atributet) {
    $rel = $Db.CreateRelation($Emri, $Tabela, $TabelaHuaj, $Atributet)
    $f = $rel.CreateField($Fusha)
    try {
        $f.ForeignName = $FushaHuaj
        $rel.Fields.Append($f)
        $Db.Relations.Append($rel)
    } finally {
        Release-ComObjekt $f
        Release-ComObjekt $rel
    }
    Write-Host ('U krijua relacioni ' + $Emri + ': ' + $Tabela + '.' + $Fusha + ' (1) -> ' + $TabelaHuaj + '.' + $FushaHuaj + ' (shumë)')
}

# Ekzekuton SQL me DAO; dbFailOnError e ndalon skriptin nëse rreshti refuzohet
function Invoke-DaoSql([string]$Sql) { $Db.Execute($Sql, $dbFailOnError) }

# Fut rreshtat e një tabele: çdo element i $Vlerat është lista VALUES e shkruar
# si tekst literal SQL (datat #yyyy-mm-dd#, dhjetorët me pikë, apostrofi '')
function Add-Rreshta([string]$Tabela, [string]$Kolonat, [string[]]$Vlerat) {
    foreach ($v in $Vlerat) {
        Invoke-DaoSql ('INSERT INTO ' + $Tabela + ' (' + $Kolonat + ') VALUES (' + $v + ')')
    }
    Write-Host ('U futën ' + $Vlerat.Count + ' rreshta në tabelën ' + $Tabela)
}

# Krijon me DAO një pyetësor të ruajtur (CreateQueryDef me emër e ruan menjëherë)
function New-AccQuery([string]$Emri, [string]$Sql) {
    if (Test-AccObject $Emri 'QUERY') { $Db.QueryDefs.Delete($Emri) }
    $qd = $Db.CreateQueryDef($Emri, $Sql)
    Release-ComObjekt $qd
    Write-Host ('U krijua pyetësori ' + $Emri)
}

# Kthen vlerën e kolonës së parë të rreshtit të parë (ose $null)
function Get-Scalar([string]$Sql) {
    $rs = $Db.OpenRecordset($Sql, $dbOpenSnapshot)
    try {
        if ($rs.EOF) { return $null }
        return $rs.Fields.Item(0).Value
    } finally {
        $rs.Close()
        Release-ComObjekt $rs
    }
}

# Kthen fushat e çelësit primar të tabelës
function Get-PrimaryKey([string]$Tabela) {
    $tdf = $Db.TableDefs.Item($Tabela)
    $rez = '-'
    try {
        for ($i = 0; $i -lt $tdf.Indexes.Count; $i++) {
            $idx = $tdf.Indexes.Item($i)
            if ($idx.Primary) {
                $emrat = @()
                try {
                    $ff = $idx.Fields
                    for ($j = 0; $j -lt $ff.Count; $j++) { $emrat += $ff.Item($j).Name }
                } catch {
                    $emrat = @((([string]$idx.Fields) -replace '[+;]', ' ').Trim())
                }
                $rez = ($emrat -join ',') + ' [' + $idx.Name + ']'
            }
            Release-ComObjekt $idx
        }
    } finally { Release-ComObjekt $tdf }
    return $rez
}

# Kthen çelësat e huaj të tabelës nga relacionet DAO
function Get-ForeignKeys([string]$Tabela) {
    $lista = @()
    for ($i = 0; $i -lt $Db.Relations.Count; $i++) {
        $r = $Db.Relations.Item($i)
        if ($r.ForeignTable -eq $Tabela) {
            $f = $r.Fields.Item(0)
            $lista += ($f.ForeignName + ' -> ' + $r.Table + '.' + $f.Name)
            Release-ComObjekt $f
        }
        Release-ComObjekt $r
    }
    if ($lista.Count -eq 0) { return '-' }
    return ($lista -join '; ')
}

# Kthen emrin e tipit të fushës siç e shfaq Access-i në Design View
function Get-LlojiFushes($Fusha) {
    $t = [int]$Fusha.Type
    $a = [int]$Fusha.Attributes
    $shprehja = ''
    try { $shprehja = [string]$Fusha.Expression } catch { $shprehja = '' }
    if ($shprehja -ne '') { return ('Calculated: ' + $shprehja) }
    if ($t -eq 4 -and ($a -band 16)) { return 'AutoNumber' }
    if ($t -eq 12 -and ($a -band $dbHyperlinkField)) { return 'Hyperlink' }
    switch ($t) {
        1   { return 'Yes/No' }
        2   { return 'Number (Byte)' }
        3   { return 'Number (Integer)' }
        4   { return 'Number (Long Integer)' }
        5   { return 'Currency' }
        6   { return 'Number (Single)' }
        7   { return 'Number (Double)' }
        8   { return 'Date/Time' }
        10  { return ('Short Text (' + $Fusha.Size + ')') }
        11  { return 'OLE Object' }
        12  { return 'Long Text' }
        101 { return 'Attachment' }
    }
    return ('Tipi ' + $t)
}

# Kthen vetinë Format të fushës, ose '-'
function Get-FormatFushes($Fusha) {
    try { return [string]$Fusha.Properties.Item('Format').Value } catch { return '-' }
}

# =====================================================================
try {

# ---------- BLLOKU 0: Kontrolli i mjedisit (PowerShell, kodimi, ACE OLEDB, DAO) ----------
try {
    Write-Host '=== P01_Zgjidhja.ps1 - Biblioteka e shkollës ==='
    Write-Host ('PowerShell: ' + $PSVersionTable.PSVersion.ToString() + ' (' + $PSVersionTable.PSEdition + ')')
    Write-Host ('Procesi 64-bit: ' + [Environment]::Is64BitProcess + ' | Windows 64-bit: ' + [Environment]::Is64BitOperatingSystem)
    Write-Host 'Test: ë ç Ë Ç'
    $TestE = 'ë'
    if ($TestE.Length -ne 1) {
        throw ("Shkronjat ë dhe ç janë prishur: skedari nuk është ruajtur si UTF-8 with BOM.`n" +
               "Zgjidhja: hape në Notepad -> File -> Save As -> Encoding: UTF-8 with BOM -> Save, pastaj ekzekutoje përsëri.")
    }
    if ($PSVersionTable.PSVersion.Major -ne 5) {
        Add-Paralajmerim 'Skripti është shkruar për Windows PowerShell 5.1 (powershell.exe), jo për PowerShell 7 (pwsh.exe).'
    }

    if ([Environment]::Is64BitProcess) {
        $Udhezim = 'Ky PowerShell është 64-bit. Nëse Office është 32-bit, ekzekuto skriptin me C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    } else {
        $Udhezim = 'Ky PowerShell është 32-bit. Nëse Office është 64-bit, ekzekuto skriptin me C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    }

    # ACE OLEDB: provohet hapja e një skedari që nuk ekziston. Nëse gabimi është
    # "Provider cannot be found" (0x800A0E7A), ofruesi mungon; çdo gabim tjetër
    # do të thotë që ofruesi u ngarkua. Nuk krijohet asnjë skedar.
    $GabimetOfruesi = @()
    $OfruesitOK = @()
    foreach ($p in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0')) {
        $t = $null
        try {
            $t = New-Object -ComObject ADODB.Connection
            $t.Open('Provider=' + $p + ';Data Source=' + (Join-Path $Folder '__prove_nuk_ekziston__.accdb') + ';')
            $t.Close()
            $OfruesitOK += $p
        } catch {
            $hr = Get-HResult $_
            if ($hr -eq -2146824582) { $GabimetOfruesi += ($p + ': nuk është i regjistruar') }        # 0x800A0E7A
            elseif ($hr -eq -2147221164) { $GabimetOfruesi += 'ADODB.Connection: Class not registered' }  # 0x80040154
            else { $OfruesitOK += $p }
        } finally {
            Release-ComObjekt $t
        }
    }
    if ($OfruesitOK.Count -eq 0) {
        throw ("ACE OLEDB nuk u gjet në këtë PowerShell (" + ($GabimetOfruesi -join '; ') + ").`n" +
               "Shkaku i mundshëm: Office dhe PowerShell kanë bitness të ndryshëm, ose izolimi i Office Click-to-Run nuk ua shfaq ofruesin programeve të jashtme.`n" +
               "Zgjidhja 1: " + $Udhezim + "`n" +
               "Zgjidhja 2: instalo 'Microsoft Access Database Engine Redistributable' me të njëjtin bitness si Office.")
    }
    Write-Host ('Ofruesit ACE OLEDB të gjetur: ' + ($OfruesitOK -join ', ')) -ForegroundColor Green

    try {
        $DbE = New-Object -ComObject DAO.DBEngine.120
    } catch {
        throw ("DAO.DBEngine.120 nuk u krijua: " + (Get-MesazhiGabimit $_) + "`n" +
               "Shkaku i mundshëm: Office dhe PowerShell kanë bitness të ndryshëm, ose izolimi i Office Click-to-Run nuk ua shfaq DAO-n programeve të jashtme.`n" +
               "Zgjidhja 1: " + $Udhezim + "`n" +
               "Zgjidhja 2: instalo 'Microsoft Access Database Engine Redistributable' me të njëjtin bitness si Office.")
    }
    Write-Host ('U gjet DAO.DBEngine.120, versioni ' + $DbE.Version) -ForegroundColor Green
} catch { Ndalo 0 'Kontrolli i mjedisit (PowerShell, kodimi, ACE OLEDB, DAO)' $_ }

# ---------- BLLOKU 1: Dosja, kontrolli i bllokimit dhe kopja rezervë ----------
try {
    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        Write-Host ('U krijua dosja ' + $Folder)
    } else {
        Write-Host ('Dosja ekziston: ' + $Folder)
    }
    if (Test-Path -LiteralPath $LockPath) {
        throw ('Baza është e hapur në Access (ekziston skedari ' + $LockPath + ").`n" +
               'Mbylle Access-in dhe ekzekutoje përsëri skriptin. Nëse Access është i mbyllur dhe skedari .laccdb mbetet, fshije me dorë.')
    }
    if (Test-Path -LiteralPath $DbPath) {
        $Vula = Get-Date -Format 'yyyyMMdd_HHmm'
        $EmriRezerve = 'P01_Zgjidhja_' + $Vula + '.bak.accdb'
        $nr = 2
        while (Test-Path -LiteralPath (Join-Path $Folder $EmriRezerve)) {
            $EmriRezerve = 'P01_Zgjidhja_' + $Vula + '_' + $nr + '.bak.accdb'
            $nr++
        }
        Rename-Item -LiteralPath $DbPath -NewName $EmriRezerve
        Write-Host ('Baza e vjetër u ruajt si kopje rezervë: ' + $EmriRezerve)
    }
} catch { Ndalo 1 'Dosja, kontrolli i bllokimit dhe kopja rezervë' $_ }

# ---------- BLLOKU 2: Krijimi i skedarit .accdb me ADOX ----------
# Provohen me radhë ofruesit e gjetur në BLLOKUN 0 (16.0, pastaj 12.0).
try {
    $GabimetKrijimi = @()
    foreach ($p in $OfruesitOK) {
        try {
            $Cat = New-Object -ComObject ADOX.Catalog
            [void]$Cat.Create('Provider=' + $p + ';Data Source=' + $DbPath + ';')
            try { $Cat.ActiveConnection.Close() } catch { }
            $Provider = $p
        } catch {
            $GabimetKrijimi += ($p + ': ' + (Get-MesazhiGabimit $_))
            # skedari nuk ekzistonte para këtij hapi (BLLOKU 1), prandaj një skedar i krijuar përgjysmë hiqet
            if (Test-Path -LiteralPath $DbPath) { Remove-Item -LiteralPath $DbPath -Force }
        } finally {
            Release-ComObjekt $Cat
            $Cat = $null
        }
        if ($null -ne $Provider) { break }
    }
    if ($null -eq $Provider) { throw ('Baza nuk u krijua me asnjë ofrues: ' + ($GabimetKrijimi -join ' | ')) }
    Write-Host ('U krijua baza: ' + $DbPath + ' (ofruesi ' + $Provider + ')')
} catch { Ndalo 2 'Krijimi i skedarit .accdb me ADOX' $_ }

# ---------- BLLOKU 3: Tabelat me ADODB (DDL) ----------
# Tipet në DDL: COUNTER = AutoNumber, TEXT(n) = Short Text, MEMO = Long Text,
# BYTE / SMALLINT / LONG / DOUBLE = Number (Byte / Integer / Long Integer / Double),
# CURRENCY = Currency, DATETIME = Date/Time, YESNO = Yes/No.
# Hyperlink, Attachment dhe Calculated shtohen me DAO në BLLOKUN 4.
try {
    $Conn = New-Object -ComObject ADODB.Connection
    $Conn.Open('Provider=' + $Provider + ';Data Source=' + $DbPath + ';')

    # 1.2 + N03: katalogu i librave (një tabelë me shumë tipe të dhënash)
    New-AccTable 'Librat' 'CREATE TABLE Librat (LibriID COUNTER CONSTRAINT PK_Librat PRIMARY KEY, Titulli TEXT(100), Autori TEXT(50), Zhanri TEXT(20), ISBN TEXT(13), VitiBotimit SMALLINT, NrFaqeve SMALLINT, NrKopjeve BYTE, Vleresimi DOUBLE, Cmimi CURRENCY, DataBlerjes DATETIME, Disponueshem YESNO, Pershkrimi MEMO)'

    # N04: lista e rrafshët e huazimeve (të dhënat e anëtarit dhe të librit përsëriten)
    New-AccTable 'Huazimet_Rrafshet' 'CREATE TABLE Huazimet_Rrafshet (HuazimiID COUNTER CONSTRAINT PK_Huazimet_Rrafshet PRIMARY KEY, Titulli TEXT(100), Autori TEXT(50), Emri TEXT(30), Mbiemri TEXT(30), Klasa TEXT(5), Telefoni TEXT(15), DataHuazimit DATETIME)'

    # N04: kalimi në tabela të lidhura (Anetaret + Huazimet; librat janë te Librat)
    New-AccTable 'Anetaret' 'CREATE TABLE Anetaret (AnetariID COUNTER CONSTRAINT PK_Anetaret PRIMARY KEY, Emri TEXT(30), Mbiemri TEXT(30), Klasa TEXT(5), Telefoni TEXT(15))'
    New-AccTable 'Huazimet' 'CREATE TABLE Huazimet (HuazimiID COUNTER CONSTRAINT PK_Huazimet PRIMARY KEY, AnetariID LONG, LibriID LONG, DataHuazimit DATETIME)'

    # N05: tabela krahasuese e teknologjive të bazave të të dhënave
    New-AccTable 'Teknologjite' 'CREATE TABLE Teknologjite (TeknologjiaID COUNTER CONSTRAINT PK_Teknologjite PRIMARY KEY, Emri TEXT(30), Prodhuesi TEXT(40), Lloji TEXT(15), Ruajtja TEXT(20), Licenca TEXT(20), PerdorimTipik TEXT(60))'

    $Conn.Close()
    Release-ComObjekt $Conn
    $Conn = $null
    Start-Sleep -Milliseconds 1000      # pritje e shkurtër që ndryshimet të shkruhen në disk
} catch { Ndalo 3 'Krijimi i tabelave me ADODB (DDL)' $_ }

# ---------- BLLOKU 4: Fushat speciale dhe vetitë e fushave (DAO) ----------
try {
    $Db = $DbE.OpenDatabase($DbPath, $false, $false)
    $Db.TableDefs.Refresh()

    # 1.2: Hyperlink = Long Text me atributin dbHyperlinkField; Attachment = tipi 101
    Add-DaoField 'Librat' 'FaqjaWeb' $dbMemo '' $dbHyperlinkField
    Add-DaoField 'Librat' 'Kopertina' $dbAttachment '' 0

    # 1.2: Calculated (rezultati Currency). Nëse formati i bazës nuk e lejon, skripti vazhdon me paralajmërim.
    try {
        Add-DaoField 'Librat' 'VleraTotale' $dbCurrency '[Cmimi]*[NrKopjeve]' 0
    } catch {
        Add-Paralajmerim ('Fusha Calculated Librat.VleraTotale nuk u krijua: ' + (Get-MesazhiGabimit $_) +
                          ' Shtoje me dorë: Librat -> Design View -> VleraTotale, Data Type = Calculated, [Cmimi]*[NrKopjeve], Result Type = Currency.')
    }
    $Db.TableDefs.Refresh()

    # Formati i shfaqjes (Kosova: dd.mm.yyyy; euro me 2 shifra dhjetore)
    Set-FieldProp 'Librat' 'Vleresimi' 'Format' $dbText 'Fixed'
    Set-FieldProp 'Librat' 'Vleresimi' 'DecimalPlaces' $dbByte 1
    Set-FieldProp 'Librat' 'Cmimi' 'Format' $dbText 'Euro'
    Set-FieldProp 'Librat' 'Cmimi' 'DecimalPlaces' $dbByte 2
    Set-FieldProp 'Librat' 'DataBlerjes' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Librat' 'Disponueshem' 'Format' $dbText 'Yes/No'
    Set-FieldProp 'Librat' 'Disponueshem' 'DisplayControl' $dbInteger $acCheckBox
    if (Test-Fusha 'Librat' 'VleraTotale') {
        Set-FieldProp 'Librat' 'VleraTotale' 'Format' $dbText 'Euro'
        Set-FieldProp 'Librat' 'VleraTotale' 'DecimalPlaces' $dbByte 2
    }
    Set-FieldProp 'Huazimet_Rrafshet' 'DataHuazimit' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Huazimet' 'DataHuazimit' 'Format' $dbText 'dd.mm.yyyy'
} catch { Ndalo 4 'Fushat speciale dhe vetitë e fushave (DAO)' $_ }

# ---------- BLLOKU 5: Relacionet (DAO) ----------
# N04: Enforce Referential Integrity, pa Cascade (Cascade është temë e P04).
try {
    New-AccRelation 'FK_Huazimet_Anetaret' 'Anetaret' 'Huazimet' 'AnetariID' 'AnetariID' 0
    New-AccRelation 'FK_Huazimet_Librat' 'Librat' 'Huazimet' 'LibriID' 'LibriID' 0
} catch { Ndalo 5 'Krijimi i relacioneve (DAO)' $_ }

# ---------- BLLOKU 6: Të dhënat shembull (prindërit para fëmijëve) ----------
# Të dhëna fiktive. Datat #yyyy-mm-dd#, numrat dhjetorë me pikë, apostrofi shkruhet ''.
try {
    # Librat: 20 libra, 8 zhanre (një zhanër për çdo kompjuter)
    Add-Rreshta 'Librat' 'LibriID, Titulli, Autori, Zhanri, ISBN, VitiBotimit, NrFaqeve, NrKopjeve, Vleresimi, Cmimi, DataBlerjes, Disponueshem, Pershkrimi, FaqjaWeb' @(
        "1, 'Rruga e gurtë', 'Agron Berisha', 'Roman', '0012345000001', 2012, 248, 4, 4.2, 9.50, #2025-09-08#, True, 'Roman për tre breza të një familjeje në Prizren.', '#https://www.example.com/libri/1#'",
        "2, 'Hija e kalasë', 'Teuta Kelmendi', 'Roman', '0012345000002', 2018, 312, 3, 4.6, 12.90, #2025-10-14#, True, 'Roman historik me ngjarje rreth një kalaje mesjetare.', '#https://www.example.com/libri/2#'",
        "3, 'Dimri në Rugovë', 'Ilir Gashi', 'Roman', '0012345000003', 2021, 276, 2, 4.8, 14.50, #2025-11-05#, False, 'Roman aventure në malet e Rugovës.', '#https://www.example.com/libri/3#'",
        "4, 'Këngë për lumin', 'Mimoza Hoxha', 'Poezi', '0012345000004', 2009, 96, 5, 3.9, 6.00, #2025-09-15#, True, 'Përmbledhje poezish për natyrën dhe vendlindjen.', '#https://www.example.com/libri/4#'",
        "5, 'Yjet mbi Sharr', 'Vjosa Rexhepi', 'Poezi', '0012345000005', 2016, 120, 3, 4.4, 7.50, #2025-12-03#, True, 'Poezi lirike për rininë dhe ëndrrat.', '#https://www.example.com/libri/5#'",
        "6, 'Fjalë të heshtura', 'Mimoza Hoxha', 'Poezi', '0012345000006', 2023, 88, 2, 4.1, 8.20, #2026-01-12#, False, 'Poezi të shkurtra për miqësinë.', '#https://www.example.com/libri/6#'",
        "7, 'Tregime nga fshati', 'Driton Shala', 'Tregime', '0012345000007', 2014, 180, 4, 4.0, 7.90, #2025-10-20#, True, 'Dhjetë tregime për jetën në fshat.', '#https://www.example.com/libri/7#'",
        "8, 'Udhëtari i natës', 'Arben Dushi', 'Tregime', '0012345000008', 2020, 204, 2, 4.3, 10.00, #2026-02-09#, True, 'Tregime misteri për lexues të rinj.', '#https://www.example.com/libri/8#'",
        "9, 'Kalatë e Dardanisë', 'Labinot Morina', 'Histori', '0012345000009', 2011, 340, 3, 4.5, 15.00, #2025-11-18#, True, 'Histori e kalave të vjetra me harta dhe fotografi.', '#https://www.example.com/libri/9#'",
        "10, 'Ilirët në Ballkan', 'Labinot Morina', 'Histori', '0012345000010', 2017, 296, 2, 4.7, 16.50, #2026-03-04#, False, 'Histori e fiseve ilire për shkollat e mesme.', '#https://www.example.com/libri/10#'",
        "11, 'Qytetet e lashta', 'Shpresa Lika', 'Histori', '0012345000011', 2022, 228, 1, 4.0, 13.20, #2026-01-21#, True, 'Histori e qyteteve antike në rajon.', '#https://www.example.com/libri/11#'",
        "12, 'Bota e atomeve', 'Ilir Gashi', 'Shkencë', '0012345000012', 2015, 210, 3, 4.4, 11.80, #2025-09-22#, True, 'Shkencë e shpjeguar thjesht: atomet dhe molekulat.', '#https://www.example.com/libri/12#'",
        "13, 'Energjia e diellit', 'Besnik Zeqiri', 'Shkencë', '0012345000013', 2024, 164, 2, 4.6, 12.40, #2026-04-07#, True, 'Si punojnë panelet diellore; s''ka nevojë për njohuri paraprake.', '#https://www.example.com/libri/13#'",
        "14, 'Ariu dhe bleta', 'Fjolla Bytyqi', 'Fëmijë', '0012345000014', 2013, 48, 6, 4.9, 4.50, #2025-12-10#, True, 'Përrallë me ilustrime për fëmijët e vegjël.', '#https://www.example.com/libri/14#'",
        "15, 'Lodrat e Ardit', 'Fjolla Bytyqi', 'Fëmijë', '0012345000015', 2019, 56, 5, 4.5, 5.20, #2026-02-16#, False, 'Tregim për fëmijë për ndarjen dhe miqësinë.', '#https://www.example.com/libri/15#'",
        "16, 'Pylli i çudirave', 'Vjosa Rexhepi', 'Fëmijë', '0012345000016', 2010, 64, 4, 4.2, 4.90, #2026-03-11#, True, 'Përrallë për kafshët e pyllit.', '#https://www.example.com/libri/16#'",
        "17, 'Rrjetet kompjuterike', 'Agron Berisha', 'Teknologji', '0012345000017', 2019, 388, 2, 4.3, 18.90, #2025-11-26#, True, 'Bazat e rrjeteve LAN, IP dhe Wi-Fi me ushtrime.', '#https://www.example.com/libri/17#'",
        "18, 'Bazat e të dhënave hap pas hapi', 'Teuta Kelmendi', 'Teknologji', '0012345000018', 2023, 256, 3, 4.8, 17.50, #2026-01-28#, True, 'Tabela, pyetësorë dhe SQL me shembuj në Access.', '#https://www.example.com/libri/18#'",
        "19, 'Malet dhe lumenjtë', 'Driton Shala', 'Gjeografi', '0012345000019', 2008, 144, 2, 3.8, 9.00, #2026-03-23#, False, 'Gjeografi fizike e Kosovës me harta.', '#https://www.example.com/libri/19#'",
        "20, 'Atlasi i komunave për shkolla', 'Valbona Musliu', 'Gjeografi', '0012345000020', 2025, 72, 4, 4.6, 11.00, #2026-04-20#, True, 'Atlas me harta të komunave dhe të lumenjve.', '#https://www.example.com/libri/20#'"
    )

    # Anetaret: 8 nxënës fiktivë (një për çdo kompjuter)
    Add-Rreshta 'Anetaret' 'AnetariID, Emri, Mbiemri, Klasa, Telefoni' @(
        "1, 'Arta', 'Krasniqi', '12-1', '044 100 201'",
        "2, 'Blerim', 'Hasani', '12-2', '045 100 202'",
        "3, 'Drenusha', 'Mustafa', '11-1', '049 100 203'",
        "4, 'Endrit', 'Kryeziu', '11-2', '044 100 204'",
        "5, 'Flutura', 'Rama', '10-1', '045 100 205'",
        "6, 'Gëzim', 'Sylejmani', '12-1', '049 100 206'",
        "7, 'Hana', 'Ahmeti', '10-2', '044 100 207'",
        "8, 'Ilirjan', 'Bajrami', '11-1', '045 100 208'"
    )

    # Teknologjite: 10 sisteme për krahasimin desktop / server / i integruar (N05)
    Add-Rreshta 'Teknologjite' 'TeknologjiaID, Emri, Prodhuesi, Lloji, Ruajtja, Licenca, PerdorimTipik' @(
        "1, 'Microsoft Access', 'Microsoft', 'Desktop', 'Skedar (.accdb)', 'Komerciale', 'Evidenca të vogla në zyrë ose shkollë'",
        "2, 'LibreOffice Base', 'The Document Foundation', 'Desktop', 'Skedar (.odb)', 'Open source', 'Evidenca personale ose të zyrës'",
        "3, 'SQLite', 'Projekti SQLite', 'I integruar', 'Skedar', 'Domen publik', 'Aplikacione celulari dhe shfletues'",
        "4, 'Microsoft SQL Server', 'Microsoft', 'Server', 'Shërbim', 'Komerciale', 'Sisteme të ndërmarrjeve'",
        "5, 'MySQL', 'Oracle', 'Server', 'Shërbim', 'Open source', 'Faqe dhe aplikacione web'",
        "6, 'Oracle Database', 'Oracle', 'Server', 'Shërbim', 'Komerciale', 'Banka dhe sisteme të mëdha'",
        "7, 'PostgreSQL', 'PostgreSQL Global Development Group', 'Server', 'Shërbim', 'Open source', 'Aplikacione web dhe harta (GIS)'",
        "8, 'MariaDB', 'MariaDB Foundation', 'Server', 'Shërbim', 'Open source', 'Faqe web, vazhdim i MySQL'",
        "9, 'IBM Db2', 'IBM', 'Server', 'Shërbim', 'Komerciale', 'Banka dhe sigurime'",
        "10, 'Firebird', 'Firebird Project', 'Server', 'Shërbim', 'Open source', 'Programe biznesi të vogla'"
    )

    # Huazimet_Rrafshet: 16 huazime në një tabelë të vetme (N04: të dhënat përsëriten)
    Add-Rreshta 'Huazimet_Rrafshet' 'HuazimiID, Titulli, Autori, Emri, Mbiemri, Klasa, Telefoni, DataHuazimit' @(
        "1, 'Rruga e gurtë', 'Agron Berisha', 'Arta', 'Krasniqi', '12-1', '044 100 201', #2026-09-07#",
        "2, 'Rrjetet kompjuterike', 'Agron Berisha', 'Blerim', 'Hasani', '12-2', '045 100 202', #2026-09-08#",
        "3, 'Yjet mbi Sharr', 'Vjosa Rexhepi', 'Drenusha', 'Mustafa', '11-1', '049 100 203', #2026-09-10#",
        "4, 'Bota e atomeve', 'Ilir Gashi', 'Endrit', 'Kryeziu', '11-2', '044 100 204', #2026-09-14#",
        "5, 'Kalatë e Dardanisë', 'Labinot Morina', 'Flutura', 'Rama', '10-1', '045 100 205', #2026-09-16#",
        "6, 'Bazat e të dhënave hap pas hapi', 'Teuta Kelmendi', 'Gëzim', 'Sylejmani', '12-1', '049 100 206', #2026-09-21#",
        "7, 'Atlasi i komunave për shkolla', 'Valbona Musliu', 'Hana', 'Ahmeti', '10-2', '044 100 207', #2026-09-23#",
        "8, 'Hija e kalasë', 'Teuta Kelmendi', 'Ilirjan', 'Bajrami', '11-1', '045 100 208', #2026-09-28#",
        "9, 'Bazat e të dhënave hap pas hapi', 'Teuta Kelmendi', 'Arta', 'Krasniqi', '12-1', '044 100 201', #2026-10-05#",
        "10, 'Rruga e gurtë', 'Agron Berisha', 'Blerim', 'Hasani', '12-2', '045 100 202', #2026-10-07#",
        "11, 'Tregime nga fshati', 'Driton Shala', 'Drenusha', 'Mustafa', '11-1', '049 100 203', #2026-10-12#",
        "12, 'Rrjetet kompjuterike', 'Agron Berisha', 'Endrit', 'Kryeziu', '11-2', '044 100 204', #2026-10-19#",
        "13, 'Pylli i çudirave', 'Vjosa Rexhepi', 'Flutura', 'Rama', '10-1', '045 100 205', #2026-10-21#",
        "14, 'Dimri në Rugovë', 'Ilir Gashi', 'Gëzim', 'Sylejmani', '12-1', '049 100 206', #2026-10-26#",
        "15, 'Energjia e diellit', 'Besnik Zeqiri', 'Hana', 'Ahmeti', '10-2', '044 100 207', #2026-11-02#",
        "16, 'Ilirët në Ballkan', 'Labinot Morina', 'Ilirjan', 'Bajrami', '11-1', '045 100 208', #2026-11-09#"
    )

    # Huazimet: të njëjtat 16 huazime, vetëm me çelësa (fëmija futet i fundit)
    Add-Rreshta 'Huazimet' 'HuazimiID, AnetariID, LibriID, DataHuazimit' @(
        "1, 1, 1, #2026-09-07#",
        "2, 2, 17, #2026-09-08#",
        "3, 3, 5, #2026-09-10#",
        "4, 4, 12, #2026-09-14#",
        "5, 5, 9, #2026-09-16#",
        "6, 6, 18, #2026-09-21#",
        "7, 7, 20, #2026-09-23#",
        "8, 8, 2, #2026-09-28#",
        "9, 1, 18, #2026-10-05#",
        "10, 2, 1, #2026-10-07#",
        "11, 3, 7, #2026-10-12#",
        "12, 4, 17, #2026-10-19#",
        "13, 5, 16, #2026-10-21#",
        "14, 6, 3, #2026-10-26#",
        "15, 7, 13, #2026-11-02#",
        "16, 8, 10, #2026-11-09#"
    )
} catch { Ndalo 6 'Futja e të dhënave' $_ }

# ---------- BLLOKU 7: Pyetësorët e ruajtur (DAO CreateQueryDef) ----------
try {
    # Kontrolli i filtrit kryesor të çdo kompjuteri (Niveli 2): Zhanri = zhanri i variantit
    New-AccQuery 'Q_Kontroll_K1' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Roman' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K2' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Poezi' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K3' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Tregime' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K4' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Histori' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K5' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Shkencë' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K6' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Fëmijë' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K7' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Teknologji' ORDER BY Librat.VitiBotimit DESC;"
    New-AccQuery 'Q_Kontroll_K8' "SELECT Librat.LibriID, Librat.Titulli, Librat.Autori, Librat.Zhanri, Librat.VitiBotimit, Librat.Cmimi FROM Librat WHERE Librat.Zhanri = 'Gjeografi' ORDER BY Librat.VitiBotimit DESC;"

    # N04: tabelat e lidhura rindërtojnë listën e rrafshët pa përsëritje në ruajtje
    New-AccQuery 'Q_Huazimet_Bashkuar' "SELECT Huazimet.HuazimiID, Librat.Titulli, Librat.Autori, Anetaret.Emri, Anetaret.Mbiemri, Anetaret.Klasa, Anetaret.Telefoni, Huazimet.DataHuazimit FROM (Huazimet INNER JOIN Anetaret ON Huazimet.AnetariID = Anetaret.AnetariID) INNER JOIN Librat ON Huazimet.LibriID = Librat.LibriID ORDER BY Huazimet.HuazimiID;"
} catch { Ndalo 7 'Krijimi i pyetësorëve' $_ }

# ---------- BLLOKU 8: Verifikimi ----------
try {
    Write-Host ''
    Write-Host '=== VERIFIKIMI ==='
    $Db.TableDefs.Refresh()
    $Db.QueryDefs.Refresh()

    # Vlerat e pritura (të njëjtat si në çelësin e përgjigjeve)
    $PritetTabela = [ordered]@{ 'Librat' = 20; 'Huazimet_Rrafshet' = 16; 'Anetaret' = 8; 'Huazimet' = 16; 'Teknologjite' = 10 }
    $PritetQ = [ordered]@{ 'Q_Kontroll_K1' = 3; 'Q_Kontroll_K2' = 3; 'Q_Kontroll_K3' = 2; 'Q_Kontroll_K4' = 3; 'Q_Kontroll_K5' = 2; 'Q_Kontroll_K6' = 3; 'Q_Kontroll_K7' = 2; 'Q_Kontroll_K8' = 2; 'Q_Huazimet_Bashkuar' = 16 }
    $ParamShembull = @{}      # P01 nuk ka pyetësorë me parametra
    $PritetVleraTotale = '602.50'
    $PritetDisponueshemJo = 5
    $PritetServer = 7
    $PritetSkedar = 3
    $Variantet = @(
        [pscustomobject]@{ K = 1; Zhanri = 'Roman'; A = 3; Shuma = '36.90'; KriteriB = 'VitiBotimit BETWEEN 2008 AND 2011'; B = 4; KriteriC = 'DataBlerjes BETWEEN #2025-09-01# AND #2025-09-30#'; C = 3; KriteriD = 'Autori LIKE ''*Berisha*'''; D = 2; KriteriE = 'Mbiemri = ''Krasniqi'''; E = 2 },
        [pscustomobject]@{ K = 2; Zhanri = 'Poezi'; A = 3; Shuma = '21.70'; KriteriB = 'VitiBotimit BETWEEN 2012 AND 2015'; B = 4; KriteriC = 'DataBlerjes BETWEEN #2025-10-01# AND #2025-10-31#'; C = 2; KriteriD = 'Autori LIKE ''*Hoxha*'''; D = 2; KriteriE = 'Mbiemri = ''Hasani'''; E = 2 },
        [pscustomobject]@{ K = 3; Zhanri = 'Tregime'; A = 2; Shuma = '17.90'; KriteriB = 'VitiBotimit BETWEEN 2016 AND 2018'; B = 3; KriteriC = 'DataBlerjes BETWEEN #2025-11-01# AND #2025-11-30#'; C = 3; KriteriD = 'Autori LIKE ''*Gashi*'''; D = 2; KriteriE = 'Mbiemri = ''Mustafa'''; E = 2 },
        [pscustomobject]@{ K = 4; Zhanri = 'Histori'; A = 3; Shuma = '44.70'; KriteriB = 'VitiBotimit BETWEEN 2019 AND 2020'; B = 3; KriteriC = 'DataBlerjes BETWEEN #2025-12-01# AND #2025-12-31#'; C = 2; KriteriD = 'Autori LIKE ''*Kelmendi*'''; D = 2; KriteriE = 'Mbiemri = ''Kryeziu'''; E = 2 },
        [pscustomobject]@{ K = 5; Zhanri = 'Shkencë'; A = 2; Shuma = '24.20'; KriteriB = 'VitiBotimit BETWEEN 2021 AND 2023'; B = 4; KriteriC = 'DataBlerjes BETWEEN #2026-01-01# AND #2026-01-31#'; C = 3; KriteriD = 'Autori LIKE ''*Shala*'''; D = 2; KriteriE = 'Mbiemri = ''Rama'''; E = 2 },
        [pscustomobject]@{ K = 6; Zhanri = 'Fëmijë'; A = 3; Shuma = '14.60'; KriteriB = 'VitiBotimit BETWEEN 2024 AND 2025'; B = 2; KriteriC = 'DataBlerjes BETWEEN #2026-02-01# AND #2026-02-28#'; C = 2; KriteriD = 'Autori LIKE ''*Rexhepi*'''; D = 2; KriteriE = 'Mbiemri = ''Sylejmani'''; E = 2 },
        [pscustomobject]@{ K = 7; Zhanri = 'Teknologji'; A = 2; Shuma = '36.40'; KriteriB = 'VitiBotimit BETWEEN 2010 AND 2013'; B = 4; KriteriC = 'DataBlerjes BETWEEN #2026-03-01# AND #2026-03-31#'; C = 3; KriteriD = 'Autori LIKE ''*Morina*'''; D = 2; KriteriE = 'Mbiemri = ''Ahmeti'''; E = 2 },
        [pscustomobject]@{ K = 8; Zhanri = 'Gjeografi'; A = 2; Shuma = '20.00'; KriteriB = 'VitiBotimit BETWEEN 2017 AND 2019'; B = 4; KriteriC = 'DataBlerjes BETWEEN #2026-04-01# AND #2026-04-30#'; C = 2; KriteriD = 'Autori LIKE ''*Bytyqi*'''; D = 2; KriteriE = 'Mbiemri = ''Bajrami'''; E = 2 }
    )

    # 8a. Tabelat: rreshtat, çelësi primar, çelësat e huaj
    $raportT = foreach ($t in $PritetTabela.Keys) {
        if (-not (Test-AccObject $t 'TABLE')) {
            $script:Mosperputhje++
            [pscustomobject]@{ Tabela = $t; Rreshta = '-'; PK = '-'; FK = '-'; Kontrolli = 'GABIM: mungon' }
            continue
        }
        $n = Get-Scalar ('SELECT Count(*) FROM [' + $t + ']')
        $ok = ($n -eq $PritetTabela[$t])
        if (-not $ok) { $script:Mosperputhje++ }
        [pscustomobject]@{
            Tabela    = $t
            Rreshta   = $n
            PK        = (Get-PrimaryKey $t)
            FK        = (Get-ForeignKeys $t)
            Kontrolli = $(if ($ok) { 'OK' } else { 'GABIM (pritej ' + $PritetTabela[$t] + ')' })
        }
    }
    Write-Host 'Tabelat:'
    $raportT | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 8b. Fushat e tabelës Librat (tema 1.2: tipet e të dhënave)
    $tdfL = $Db.TableDefs.Item('Librat')
    $raportF = for ($i = 0; $i -lt $tdfL.Fields.Count; $i++) {
        $f = $tdfL.Fields.Item($i)
        [pscustomobject]@{ Fusha = $f.Name; Tipi = (Get-LlojiFushes $f); Format = (Get-FormatFushes $f) }
        Release-ComObjekt $f
    }
    Release-ComObjekt $tdfL
    Write-Host 'Fushat e tabelës Librat:'
    $raportF | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 8c. Relacionet
    Write-Host 'Relacionet:'
    for ($i = 0; $i -lt $Db.Relations.Count; $i++) {
        $r = $Db.Relations.Item($i)
        if ($r.Name -like 'FK_*') {
            $cascade = 'pa Cascade'
            if (([int]$r.Attributes -band 256) -or ([int]$r.Attributes -band 4096)) { $cascade = 'me Cascade' }
            $rif = 'Enforce Referential Integrity'
            if ([int]$r.Attributes -band 2) { $rif = 'PA integritet referencial' }
            Write-Host ('  ' + $r.Name + ': ' + $r.Table + ' (1) -> ' + $r.ForeignTable + ' (shumë), ' + $rif + ', ' + $cascade)
        }
        Release-ComObjekt $r
    }
    foreach ($emriRel in @('FK_Huazimet_Anetaret', 'FK_Huazimet_Librat')) {
        $gjetRel = $false
        for ($i = 0; $i -lt $Db.Relations.Count; $i++) {
            $r = $Db.Relations.Item($i)
            if ($r.Name -eq $emriRel) { $gjetRel = $true }
            Release-ComObjekt $r
        }
        if (-not $gjetRel) { $script:Mosperputhje++; Write-Host ('  GABIM: mungon relacioni ' + $emriRel) -ForegroundColor Red }
    }

    # 8d. Pyetësorët: lloji dhe numri i rreshtave
    $LlojiQ = @{ 0 = 'Select'; 16 = 'Crosstab'; 32 = 'Delete'; 48 = 'Update'; 64 = 'Append'; 80 = 'Make-Table'; 96 = 'Data Definition' }
    $raportQ = for ($i = 0; $i -lt $Db.QueryDefs.Count; $i++) {
        $qd = $Db.QueryDefs.Item($i)
        $emri = [string]$qd.Name
        if ($emri.StartsWith('~')) { Release-ComObjekt $qd; continue }
        $tipi = [int]$qd.Type
        $lloji = [string]$LlojiQ[$tipi]
        if ($lloji -eq '') { $lloji = 'Tipi ' + $tipi }
        if ($tipi -eq 0 -and [string]$qd.SQL -match 'GROUP BY') { $lloji = 'Select (Totals)' }
        $nParam = [int]$qd.Parameters.Count
        if ($nParam -gt 0) { $lloji = $lloji + ' + Parameters' }
        $rreshta = '-'
        if ($tipi -eq 0 -or $tipi -eq 16) {
            $kaVlera = $true
            for ($j = 0; $j -lt $nParam; $j++) {
                $prm = $qd.Parameters.Item($j)
                $celesi = $emri + '|' + $prm.Name
                if ($ParamShembull.ContainsKey($celesi)) { $prm.Value = $ParamShembull[$celesi] } else { $kaVlera = $false }
                Release-ComObjekt $prm
            }
            if ($kaVlera) {
                $rs = $qd.OpenRecordset($dbOpenSnapshot)
                if (-not $rs.EOF) { $rs.MoveLast() }
                $rreshta = [int]$rs.RecordCount
                $rs.Close()
                Release-ComObjekt $rs
            } else { $rreshta = 'pa vlerë shembull' }
        } else { $rreshta = 'pyetësor veprimi: nuk ekzekutohet' }
        $kontrolli = '-'
        if ($PritetQ.Contains($emri)) {
            if ($rreshta -eq $PritetQ[$emri]) { $kontrolli = 'OK' } else { $kontrolli = 'GABIM (pritej ' + $PritetQ[$emri] + ')'; $script:Mosperputhje++ }
        }
        [pscustomobject]@{ Pyetesori = $emri; Lloji = $lloji; Rreshta = $rreshta; Kontrolli = $kontrolli }
        Release-ComObjekt $qd
    }
    Write-Host 'Pyetësorët e ruajtur:'
    $raportQ | Format-Table -AutoSize | Out-String -Width 250 | Write-Host
    foreach ($q in $PritetQ.Keys) {
        if (-not (Test-AccObject $q 'QUERY')) { $script:Mosperputhje++; Write-Host ('GABIM: mungon pyetësori ' + $q) -ForegroundColor Red }
    }

    # 8e. Rezultatet e 8 varianteve (Kompjuteri 1-8), krahasuar me çelësin e përgjigjeve
    $raportV = foreach ($v in $Variantet) {
        $a  = Get-Scalar ("SELECT Count(*) FROM Librat WHERE Zhanri = '" + $v.Zhanri + "'")
        $sh = Get-Scalar ("SELECT Sum(Cmimi) FROM Librat WHERE Zhanri = '" + $v.Zhanri + "'")
        $b  = Get-Scalar ('SELECT Count(*) FROM Librat WHERE ' + $v.KriteriB)
        $c  = Get-Scalar ('SELECT Count(*) FROM Librat WHERE ' + $v.KriteriC)
        $dd = Get-Scalar ('SELECT Count(*) FROM Librat WHERE ' + $v.KriteriD)
        $e  = Get-Scalar ('SELECT Count(*) FROM Huazimet_Rrafshet WHERE ' + $v.KriteriE)
        $ok = ($a -eq $v.A) -and ([decimal]$sh -eq [decimal]$v.Shuma) -and ($b -eq $v.B) -and ($c -eq $v.C) -and ($dd -eq $v.D) -and ($e -eq $v.E)
        if (-not $ok) { $script:Mosperputhje++ }
        [pscustomobject]@{
            K          = $v.K
            Zhanri     = $v.Zhanri
            A_Zhanri   = $a
            ShumaCmimi = ([decimal]$sh).ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
            B_Viti     = $b
            C_Data     = $c
            D_Autori   = $dd
            E_Anetari  = $e
            Kontrolli  = $(if ($ok) { 'OK' } else { 'GABIM' })
        }
    }
    Write-Host 'Variantet (A = Zhanri, B = VitiBotimit, C = DataBlerjes, D = Autori, E = anëtari te Huazimet_Rrafshet):'
    $raportV | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 8f. Kontrolle të veçanta
    $nPerputhje = Get-Scalar ('SELECT Count(*) FROM Huazimet_Rrafshet AS R INNER JOIN Q_Huazimet_Bashkuar AS Q ON (R.HuazimiID = Q.HuazimiID) AND (R.Titulli = Q.Titulli) AND (R.Autori = Q.Autori) AND (R.Emri = Q.Emri) AND (R.Mbiemri = Q.Mbiemri) AND (R.Klasa = Q.Klasa) AND (R.Telefoni = Q.Telefoni) AND (R.DataHuazimit = Q.DataHuazimit)')
    if ($nPerputhje -eq $PritetTabela['Huazimet_Rrafshet']) { Write-Host ('N04: Q_Huazimet_Bashkuar rindërton saktë listën e rrafshët (' + $nPerputhje + ' rreshta të njëjtë).') -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM N04: vetëm ' + $nPerputhje + ' rreshta përputhen me Huazimet_Rrafshet.') -ForegroundColor Red }

    $zhanri12 = [string](Get-Scalar 'SELECT Zhanri FROM Librat WHERE LibriID = 12')
    if ($zhanri12 -ceq 'Shkencë') { Write-Host 'Kodimi: shkronjat ë/ç u ruajtën saktë në bazë (LibriID 12 = Shkencë).' -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM kodimi: LibriID 12 ka Zhanri = ' + $zhanri12) -ForegroundColor Red }

    if (Test-Fusha 'Librat' 'VleraTotale') {
        $vt = Get-Scalar 'SELECT Sum(VleraTotale) FROM Librat'
        if ([decimal]$vt -eq [decimal]$PritetVleraTotale) { Write-Host ('1.2: fusha Calculated VleraTotale punon (shuma = ' + $PritetVleraTotale + ' €).') -ForegroundColor Green }
        else { $script:Mosperputhje++; Write-Host ('GABIM: shuma e VleraTotale = ' + $vt + ', pritej ' + $PritetVleraTotale) -ForegroundColor Red }
    }

    $nJo = Get-Scalar 'SELECT Count(*) FROM Librat WHERE Disponueshem = False'
    Write-Host ('Librat me Disponueshem = No: ' + $nJo + ' (pritej ' + $PritetDisponueshemJo + ')')
    if ($nJo -ne $PritetDisponueshemJo) { $script:Mosperputhje++ }

    $nServer = Get-Scalar "SELECT Count(*) FROM Teknologjite WHERE Lloji = 'Server'"
    $nSkedar = Get-Scalar "SELECT Count(*) FROM Teknologjite WHERE Ruajtja LIKE 'Skedar*'"
    Write-Host ('N05: Lloji = Server: ' + $nServer + ' (pritej ' + $PritetServer + '); Ruajtja fillon me Skedar: ' + $nSkedar + ' (pritej ' + $PritetSkedar + ')')
    if ($nServer -ne $PritetServer -or $nSkedar -ne $PritetSkedar) { $script:Mosperputhje++ }

    if ($script:Mosperputhje -gt 0) {
        throw ('Verifikimi gjeti ' + $script:Mosperputhje + ' mospërputhje. Shih rreshtat GABIM më sipër.')
    }

    $Db.Close()
    Release-ComObjekt $Db
    $Db = $null

    Write-Host ''
    Write-Host 'Verifikimi: të gjitha kontrollet OK.' -ForegroundColor Green
    if ($script:Paralajmerime.Count -gt 0) {
        Write-Host ('Paralajmërime: ' + $script:Paralajmerime.Count) -ForegroundColor Yellow
        foreach ($pp in $script:Paralajmerime) { Write-Host ('  - ' + $pp) -ForegroundColor Yellow }
    }
    Write-Host 'Përfundoi. Hap C:\BazatDB\P01\P01_Zgjidhja.accdb në Access.' -ForegroundColor Green
} catch { Ndalo 8 'Verifikimi' $_ }

# =====================================================================
}
catch [System.OperationCanceledException] {
    # Mesazhi "GABIM në BLLOKU n" është shfaqur tashmë nga funksioni Ndalo.
}
catch {
    Write-Host ('GABIM i papritur: ' + (Get-MesazhiGabimit $_)) -ForegroundColor Red
    $script:KodiDaljes = 99
}
finally {
    # Mbyll lidhjet dhe liron objektet COM, që të mos mbetet asnjë bllokim (.laccdb)
    if ($null -ne $Db)   { try { $Db.Close() } catch { } }
    if ($null -ne $Conn) { try { if ($Conn.State -ne 0) { $Conn.Close() } } catch { } }
    Release-ComObjekt $Db
    Release-ComObjekt $Conn
    Release-ComObjekt $Cat
    Release-ComObjekt $DbE
    $Db = $null; $Conn = $null; $Cat = $null; $DbE = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}
exit $script:KodiDaljes
