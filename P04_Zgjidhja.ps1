# =====================================================================
# P04_Zgjidhja.ps1  -  "Bazat e të dhënave", Klasa 12, Projekti P04
# Magazina: emërtimi, vetitë e fushave, çelësat dhe indekset, validimi,
# fushat Lookup dhe integriteti referencial (Cascade)
#
# Qëllimi : krijon C:\BazatDB\P04\P04_Zgjidhja.accdb me tabelat, vetitë
#           e fushave, indekset, rregullat e validimit, fushat Lookup,
#           relacionet (me dhe pa Cascade), të dhënat dhe pyetësorët e P04.
#           Testet e Cascade bëhen nga nxënësit vetëm te tabelat kopje
#           Kategorite_Kopje dhe Artikujt_Kopje.
#           Mësimdhënësi e kopjon bazën me USB në 8 kompjuterët.
#           Nxënësit nuk e ekzekutojnë skriptin.
# Temat   : N02, 1.1, N03, 1.6, 1.4, 1.5, 2.2+N10 (.xlsx, «Plani i orëve», Thelbësore)
# Kërkesat: Windows PowerShell 5.1; Microsoft Access me ACE OLEDB
#           (16.0 ose 12.0) dhe DAO.DBEngine.120, me të njëjtin bitness
#           si PowerShell-i që e ekzekuton skriptin.
#           Pa Access.Application, pa VBE, pa SendKeys.
# Ruaje skedarin si "UTF-8 with BOM" që shkronjat ë, ç, Ë, Ç të mos prishen.
#
# Ekzekutimi, Office 64-bit:
#   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P04_Zgjidhja.ps1"
# Ekzekutimi, Office 32-bit:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P04_Zgjidhja.ps1"
#
# Rendi: dosja -> kopja rezervë -> baza -> tabelat -> vetitë e fushave ->
#        relacionet -> të dhënat (prindërit para fëmijëve) -> pyetësorët ->
#        verifikimi. Kodet e daljes: 0 = në rregull, 10-18 = gabim në
#        bllokun 0-8, 99 = gabim i papritur.
# =====================================================================

$ErrorActionPreference = 'Stop'

$Folder   = 'C:\BazatDB\P04'
$DbPath   = Join-Path $Folder 'P04_Zgjidhja.accdb'
$LockPath = Join-Path $Folder 'P04_Zgjidhja.laccdb'

$script:KodiDaljes    = 0
$script:Mosperputhje  = 0
$script:Paralajmerime = New-Object System.Collections.ArrayList
$Provider = $null
$Cat      = $null
$Conn     = $null
$DbE      = $null
$Db       = $null

# Konstantet DAO me vlerat numerike (PowerShell nuk i njeh emrat dbText etj.)
$dbByte           = 2
$dbInteger        = 3
$dbText           = 10
$acCheckBox       = 106
$dbHyperlinkField = 32768
$dbFailOnError    = 128
$dbOpenSnapshot   = 4
$dbRelationUnique = 1
$dbBoolean        = 1
$dbRelUpdCascade  = 256
$dbRelDelCascade  = 4096
$acComboBox       = 111

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

# Ekzekuton një deklaratë SQL me ADODB (CREATE TABLE në BLLOKUN 3)
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

# Vendos me DAO një veti të tabelës (p.sh. SubdatasheetName); nëse nuk ekziston, e krijon
function Set-TableProp([string]$Tabela, [string]$Emri, [int]$Tipi, $Vlera) {
    $tdf = $Db.TableDefs.Item($Tabela)
    try {
        $ekziston = $true
        $p = $null
        try { $p = $tdf.Properties.Item($Emri) } catch { $ekziston = $false }
        if ($ekziston) {
            $p.Value = $Vlera
        } else {
            $p = $tdf.CreateProperty($Emri, $Tipi, $Vlera)
            $tdf.Properties.Append($p)
        }
        Release-ComObjekt $p
    } finally {
        Release-ComObjekt $tdf
    }
    Write-Host ('U vendos vetia e tabelës ' + $Tabela + ': ' + $Emri + ' = ' + $Vlera)
}

# Krijon me DAO një relacion me një fushë (Enforce Referential Integrity).
# Atributet: 0 = një-me-shumë; +256 = Cascade Update; +4096 = Cascade Delete.
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
    $ana = '(shumë)'
    if ($Atributet -band $dbRelationUnique) { $ana = '(1)' }
    $kask = ''
    if ($Atributet -band $dbRelUpdCascade) { $kask += ', Cascade Update' }
    if ($Atributet -band $dbRelDelCascade) { $kask += ', Cascade Delete' }
    Write-Host ('U krijua relacioni ' + $Emri + ': ' + $Tabela + '.' + $Fusha + ' (1) -> ' + $TabelaHuaj + '.' + $FushaHuaj + ' ' + $ana + $kask)
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

# =====================================================================
try {

# ---------- BLLOKU 0: Kontrolli i mjedisit (PowerShell, kodimi, ACE OLEDB, DAO) ----------
try {
    Write-Host '=== P04_Zgjidhja.ps1 - Magazina (vetitë, validimi, integriteti) ==='
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
        $EmriRezerve = 'P04_Zgjidhja_' + $Vula + '.bak.accdb'
        $nr = 2
        while (Test-Path -LiteralPath (Join-Path $Folder $EmriRezerve)) {
            $EmriRezerve = 'P04_Zgjidhja_' + $Vula + '_' + $nr + '.bak.accdb'
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

# ---------- BLLOKU 3: Tabelat dhe indekset me ADODB (DDL) ----------
# NOT NULL = Required = Yes; UNIQUE = Indexed: Yes (No Duplicates) (tema 1.6).
try {
    $Conn = New-Object -ComObject ADODB.Connection
    $Conn.Open('Provider=' + $Provider + ';Data Source=' + $DbPath + ';')

    # N02: emra të propozuar dhe emri i saktë sipas konventës
    New-AccTable 'Emertimet' 'CREATE TABLE Emertimet (EmertimiID COUNTER CONSTRAINT PK_Emertimet PRIMARY KEY, EmriIPropozuar TEXT(40), Objekti TEXT(10), Problemi TEXT(60), EmriISakte TEXT(40))'

    New-AccTable 'Kategorite' 'CREATE TABLE Kategorite (KategoriaID COUNTER CONSTRAINT PK_Kategorite PRIMARY KEY, EmriKategorise TEXT(30) NOT NULL CONSTRAINT UQ_Kategorite_Emri UNIQUE)'
    New-AccTable 'Furnitoret' 'CREATE TABLE Furnitoret (FurnitoriID COUNTER CONSTRAINT PK_Furnitoret PRIMARY KEY, EmriFurnitorit TEXT(50) NOT NULL, Qyteti TEXT(30), Telefoni TEXT(15), Email TEXT(50))'
    New-AccTable 'Artikujt' 'CREATE TABLE Artikujt (ArtikulliID COUNTER CONSTRAINT PK_Artikujt PRIMARY KEY, Kodi TEXT(8) NOT NULL CONSTRAINT UQ_Artikujt_Kodi UNIQUE, Emertimi TEXT(50) NOT NULL, KategoriaID LONG NOT NULL, FurnitoriID LONG, Cmimi CURRENCY NOT NULL, Sasia SMALLINT, MinSasia SMALLINT, DataFutjes DATETIME, Aktiv YESNO)'
    New-AccTable 'Levizjet' 'CREATE TABLE Levizjet (LevizjaID COUNTER CONSTRAINT PK_Levizjet PRIMARY KEY, ArtikulliID LONG NOT NULL, DataLevizjes DATETIME NOT NULL, Lloji TEXT(6) NOT NULL, Sasia SMALLINT NOT NULL)'

    # 2.2+N10: tabelat kopje për testet e Cascade (origjinalet nuk preken)
    New-AccTable 'Kategorite_Kopje' 'CREATE TABLE Kategorite_Kopje (KategoriaID LONG CONSTRAINT PK_Kategorite_Kopje PRIMARY KEY, EmriKategorise TEXT(30))'
    New-AccTable 'Artikujt_Kopje' 'CREATE TABLE Artikujt_Kopje (ArtikulliID LONG CONSTRAINT PK_Artikujt_Kopje PRIMARY KEY, Kodi TEXT(8), Emertimi TEXT(50), KategoriaID LONG, Sasia SMALLINT)'

    # 1.6: indekse shtesë
    Invoke-Sql 'CREATE UNIQUE INDEX IX_Furnitoret_EmriQyteti ON Furnitoret (EmriFurnitorit, Qyteti)'
    Invoke-Sql 'CREATE INDEX IX_Levizjet_Data ON Levizjet (DataLevizjes)'
    Write-Host 'U krijuan indekset IX_Furnitoret_EmriQyteti (unik, i përbërë) dhe IX_Levizjet_Data'

    $Conn.Close()
    Release-ComObjekt $Conn
    $Conn = $null
    Start-Sleep -Milliseconds 1000      # pritje e shkurtër që ndryshimet të shkruhen në disk
} catch { Ndalo 3 'Krijimi i tabelave dhe indekseve me ADODB (DDL)' $_ }

# ---------- BLLOKU 4: Vetitë e fushave, validimi dhe Lookup (DAO) ----------
try {
    $Db = $DbE.OpenDatabase($DbPath, $false, $false)
    $Db.TableDefs.Refresh()

    # 1.1: Caption, Format, Default Value
    Set-FieldProp 'Artikujt' 'Emertimi' 'Caption' $dbText 'Emërtimi'
    Set-FieldProp 'Artikujt' 'Cmimi' 'Caption' $dbText 'Çmimi (€)'
    Set-FieldProp 'Artikujt' 'Cmimi' 'Format' $dbText 'Euro'
    Set-FieldProp 'Artikujt' 'Cmimi' 'DecimalPlaces' $dbByte 2
    Set-FieldProp 'Artikujt' 'MinSasia' 'Caption' $dbText 'Sasia minimale'
    Set-FieldProp 'Artikujt' 'DataFutjes' 'Caption' $dbText 'Data e futjes'
    Set-FieldProp 'Artikujt' 'DataFutjes' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Artikujt' 'Aktiv' 'Format' $dbText 'Yes/No'
    Set-FieldProp 'Artikujt' 'Aktiv' 'DisplayControl' $dbInteger $acCheckBox
    Set-FieldProp 'Artikujt' 'Sasia' 'DefaultValue' $dbText '0'
    Set-FieldProp 'Artikujt' 'MinSasia' 'DefaultValue' $dbText '5'
    Set-FieldProp 'Artikujt' 'DataFutjes' 'DefaultValue' $dbText 'Date()'
    Set-FieldProp 'Artikujt' 'Aktiv' 'DefaultValue' $dbText 'Yes'
    Set-FieldProp 'Levizjet' 'DataLevizjes' 'DefaultValue' $dbText 'Date()'
    Set-FieldProp 'Levizjet' 'DataLevizjes' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Levizjet' 'Lloji' 'DefaultValue' $dbText '"Hyrje"'

    # 1.4: Validation Rule dhe Validation Text (vendosen para të dhënave: çdo INSERT kontrollohet)
    $Validimet = @(
        @('Artikujt', 'Cmimi', '>0', 'Çmimi duhet të jetë më i madh se 0.'),
        @('Artikujt', 'Sasia', '>=0', 'Sasia nuk mund të jetë negative.'),
        @('Artikujt', 'MinSasia', 'Between 0 And 100', 'Sasia minimale duhet të jetë nga 0 deri në 100.'),
        @('Artikujt', 'DataFutjes', '<=Date()', 'Data e futjes nuk mund të jetë në të ardhmen.'),
        @('Levizjet', 'Lloji', 'In ("Hyrje","Dalje")', 'Lloji duhet të jetë Hyrje ose Dalje.'),
        @('Levizjet', 'Sasia', '>0', 'Sasia e lëvizjes duhet të jetë më e madhe se 0.')
    )
    foreach ($vv in $Validimet) {
        Set-FieldProp $vv[0] $vv[1] 'ValidationRule' $dbText $vv[2]
        Set-FieldProp $vv[0] $vv[1] 'ValidationText' $dbText $vv[3]
    }

    # 1.5: fushat Lookup (ComboBox me dy kolona; kolona e parë e fshehur)
    $Lookup = @(
        @('Artikujt', 'KategoriaID', 'SELECT Kategorite.KategoriaID, Kategorite.EmriKategorise FROM Kategorite ORDER BY Kategorite.EmriKategorise;', 'Kategoria', 2),
        @('Artikujt', 'FurnitoriID', 'SELECT Furnitoret.FurnitoriID, Furnitoret.EmriFurnitorit, Furnitoret.Qyteti FROM Furnitoret ORDER BY Furnitoret.EmriFurnitorit;', 'Furnitori', 3),
        @('Levizjet', 'ArtikulliID', 'SELECT Artikujt.ArtikulliID, Artikujt.Kodi, Artikujt.Emertimi FROM Artikujt ORDER BY Artikujt.Kodi;', 'Artikulli', 3)
    )
    foreach ($lk in $Lookup) {
        $nrKol = [int]$lk[4]
        Set-FieldProp $lk[0] $lk[1] 'DisplayControl' $dbInteger $acComboBox
        Set-FieldProp $lk[0] $lk[1] 'RowSourceType' $dbText 'Table/Query'
        Set-FieldProp $lk[0] $lk[1] 'RowSource' $dbText $lk[2]
        Set-FieldProp $lk[0] $lk[1] 'BoundColumn' $dbInteger 1
        Set-FieldProp $lk[0] $lk[1] 'ColumnCount' $dbInteger $nrKol
        Set-FieldProp $lk[0] $lk[1] 'ColumnWidths' $dbText ('0;' + ((@('2268') * ($nrKol - 1)) -join ';'))
        Set-FieldProp $lk[0] $lk[1] 'LimitToList' $dbBoolean $true
        Set-FieldProp $lk[0] $lk[1] 'Caption' $dbText $lk[3]
    }
} catch { Ndalo 4 'Vetitë e fushave, validimi dhe Lookup (DAO)' $_ }

# ---------- BLLOKU 5: Relacionet (DAO) ----------
# 2.2+N10: origjinalet pa Cascade (përveç Furnitoret: Cascade Update);
# kopjet me Cascade Update + Cascade Delete, për testet e nxënësve.
try {
    New-AccRelation 'FK_Artikujt_Kategorite' 'Kategorite' 'Artikujt' 'KategoriaID' 'KategoriaID' 0
    New-AccRelation 'FK_Artikujt_Furnitoret' 'Furnitoret' 'Artikujt' 'FurnitoriID' 'FurnitoriID' $dbRelUpdCascade
    New-AccRelation 'FK_Levizjet_Artikujt' 'Artikujt' 'Levizjet' 'ArtikulliID' 'ArtikulliID' 0
    New-AccRelation 'FK_ArtikujtKopje_KategoriteKopje' 'Kategorite_Kopje' 'Artikujt_Kopje' 'KategoriaID' 'KategoriaID' ($dbRelUpdCascade + $dbRelDelCascade)
} catch { Ndalo 5 'Krijimi i relacioneve (DAO)' $_ }

# ---------- BLLOKU 6: Të dhënat shembull (prindërit para fëmijëve) ----------
# Të dhëna fiktive. Datat #yyyy-mm-dd#, numrat dhjetorë me pikë, apostrofi shkruhet ''.
try {
    # Emertimet (N02): 10 emra të propozuar dhe emri i saktë
    Add-Rreshta 'Emertimet' 'EmertimiID, EmriIPropozuar, Objekti, Problemi, EmriISakte' @(
        "1, 'Tabela e Artikujve', 'Tabelë', 'ka hapësira', 'Artikujt'",
        "2, 'tbl_Artikujt', 'Tabelë', 'tabelat nuk kanë parashtesë', 'Artikujt'",
        "3, 'Çmimi', 'Fushë', 'ka ç', 'Cmimi'",
        "4, 'Date', 'Fushë', 'fjalë e rezervuar', 'DataFutjes'",
        "5, 'Name', 'Fushë', 'fjalë e rezervuar', 'Emertimi'",
        "6, 'Sasia në stok', 'Fushë', 'hapësira dhe ë', 'SasiaStok'",
        "7, 'Level', 'Fushë', 'fjalë e rezervuar', 'NiveliMin'",
        "8, 'qryArtikujt', 'Pyetësor', 'parashtesa duhet Q_', 'Q_Artikujt'",
        "9, 'Forma Artikujt', 'Formë', 'hapësirë; parashtesa F_', 'F_Artikujt'",
        "10, 'Raporti1', 'Raport', 'parashtesa R_; emër pa kuptim', 'R_Stoku'"
    )

    # Kategorite (prind): 8 kategori, një për çdo kompjuter
    Add-Rreshta 'Kategorite' 'KategoriaID, EmriKategorise' @(
        "1, 'Shkrim'",
        "2, 'Letër'",
        "3, 'Kompjuterë'",
        "4, 'Rrjeta'",
        "5, 'Pastrim'",
        "6, 'Mobilje'",
        "7, 'Elektrike'",
        "8, 'Arkivim'"
    )

    # Furnitoret (prind): 5 furnitorë fiktivë; «Zyra Plus» në dy qytete (indeksi unik i përbërë e lejon)
    Add-Rreshta 'Furnitoret' 'FurnitoriID, EmriFurnitorit, Qyteti, Telefoni, Email' @(
        "1, 'Zyra Plus', 'Prishtinë', '044 510 101', 'shitja@zyraplus.example'",
        "2, 'TekNet', 'Prizren', '045 510 202', 'info@teknet.example'",
        "3, 'Pastra', 'Gjakovë', '049 510 303', 'porosi@pastra.example'",
        "4, 'Mobilia', 'Pejë', '044 510 404', 'kontakt@mobilia.example'",
        "5, 'Zyra Plus', 'Prizren', '045 510 505', 'prizren@zyraplus.example'"
    )

    # Artikujt: 20 artikuj (kontrollohen nga Validation Rule gjatë futjes)
    Add-Rreshta 'Artikujt' 'ArtikulliID, Kodi, Emertimi, KategoriaID, FurnitoriID, Cmimi, Sasia, MinSasia, DataFutjes, Aktiv' @(
        "1, 'ART-001', 'Stilolaps blu', 1, 1, 0.35, 7, 10, #2026-01-18#, True",
        "2, 'ART-002', 'Laps grafit HB', 1, 1, 0.20, 14, 15, #2026-01-31#, True",
        "3, 'ART-003', 'Marker i zi', 1, 1, 1.10, 21, 5, #2026-02-13#, True",
        "4, 'ART-004', 'Letër A4 500 fletë', 2, 1, 4.90, 28, 10, #2026-02-26#, True",
        "5, 'ART-005', 'Fletore A5', 2, 1, 1.25, 35, 15, #2026-03-11#, True",
        "6, 'ART-006', 'Letër ngjitëse', 2, 1, 2.40, 2, 5, #2026-03-24#, True",
        "7, 'ART-007', 'Maus pa tel', 3, 2, 12.50, 9, 10, #2026-04-06#, True",
        "8, 'ART-008', 'Tastierë USB', 3, 2, 15.90, 16, 15, #2026-04-19#, True",
        "9, 'ART-009', 'Kabllo UTP 5 m', 4, 2, 3.80, 23, 5, #2026-05-02#, False",
        "10, 'ART-010', 'Switch 8 porta', 4, 2, 24.00, 30, 10, #2026-05-15#, True",
        "11, 'ART-011', 'Lidhës RJ45', 4, 2, 0.15, 37, 15, #2026-05-28#, True",
        "12, 'ART-012', 'Detergjent 1 l', 5, 3, 2.70, 4, 5, #2026-06-10#, True",
        "13, 'ART-013', 'Peceta letre', 5, 3, 1.60, 11, 10, #2026-06-23#, True",
        "14, 'ART-014', 'Karrige zyre', 6, 4, 79.00, 18, 15, #2026-07-06#, True",
        "15, 'ART-015', 'Raft metalik', 6, 4, 55.50, 25, 5, #2026-07-19#, True",
        "16, 'ART-016', 'Zgjatues 5 prizë', 7, 5, 8.90, 32, 10, #2026-08-01#, True",
        "17, 'ART-017', 'Llambë LED 9 W', 7, 5, 2.30, 39, 15, #2026-08-14#, True",
        "18, 'ART-018', 'Bateri AA 4 copë', 7, 5, 3.20, 6, 5, #2026-08-27#, False",
        "19, 'ART-019', 'Dosje me kapëse', 8, 5, 1.95, 13, 10, #2026-09-09#, True",
        "20, 'ART-020', 'Kuti arkivi', 8, 5, 3.40, 20, 15, #2026-09-22#, True"
    )

    # Levizjet (fëmijë): 36 lëvizje hyrje/dalje
    Add-Rreshta 'Levizjet' 'LevizjaID, ArtikulliID, DataLevizjes, Lloji, Sasia' @(
        "1, 1, #2026-09-06#, 'Hyrje', 21",
        "2, 1, #2026-09-11#, 'Dalje', 7",
        "3, 1, #2026-09-16#, 'Dalje', 6",
        "4, 4, #2026-09-21#, 'Hyrje', 20",
        "5, 4, #2026-09-26#, 'Dalje', 6",
        "6, 4, #2026-10-01#, 'Dalje', 5",
        "7, 7, #2026-10-06#, 'Hyrje', 23",
        "8, 7, #2026-10-11#, 'Dalje', 5",
        "9, 7, #2026-10-16#, 'Dalje', 4",
        "10, 9, #2026-10-21#, 'Hyrje', 22",
        "11, 9, #2026-10-26#, 'Dalje', 8",
        "12, 9, #2026-10-31#, 'Dalje', 3",
        "13, 12, #2026-11-05#, 'Hyrje', 21",
        "14, 12, #2026-11-10#, 'Dalje', 7",
        "15, 12, #2026-11-15#, 'Dalje', 6",
        "16, 14, #2026-09-01#, 'Hyrje', 20",
        "17, 14, #2026-09-06#, 'Dalje', 6",
        "18, 14, #2026-09-11#, 'Dalje', 5",
        "19, 16, #2026-09-16#, 'Hyrje', 23",
        "20, 16, #2026-09-21#, 'Dalje', 5",
        "21, 16, #2026-09-26#, 'Dalje', 4",
        "22, 19, #2026-10-01#, 'Hyrje', 22",
        "23, 19, #2026-10-06#, 'Dalje', 8",
        "24, 19, #2026-10-11#, 'Dalje', 3",
        "25, 2, #2026-10-16#, 'Hyrje', 11",
        "26, 3, #2026-10-21#, 'Hyrje', 12",
        "27, 5, #2026-10-26#, 'Hyrje', 13",
        "28, 6, #2026-10-31#, 'Hyrje', 10",
        "29, 8, #2026-11-05#, 'Hyrje', 11",
        "30, 10, #2026-11-10#, 'Hyrje', 12",
        "31, 11, #2026-11-15#, 'Hyrje', 13",
        "32, 13, #2026-09-01#, 'Hyrje', 10",
        "33, 15, #2026-09-06#, 'Hyrje', 11",
        "34, 17, #2026-09-11#, 'Hyrje', 12",
        "35, 18, #2026-09-16#, 'Hyrje', 13",
        "36, 20, #2026-09-21#, 'Hyrje', 10"
    )

    # Kopjet për testet e Cascade (2.2+N10)
    Add-Rreshta 'Kategorite_Kopje' 'KategoriaID, EmriKategorise' @(
        "1, 'Shkrim'",
        "2, 'Letër'",
        "3, 'Kompjuterë'",
        "4, 'Rrjeta'",
        "5, 'Pastrim'",
        "6, 'Mobilje'",
        "7, 'Elektrike'",
        "8, 'Arkivim'"
    )
    Add-Rreshta 'Artikujt_Kopje' 'ArtikulliID, Kodi, Emertimi, KategoriaID, Sasia' @(
        "1, 'ART-001', 'Stilolaps blu', 1, 7",
        "2, 'ART-002', 'Laps grafit HB', 1, 14",
        "3, 'ART-003', 'Marker i zi', 1, 21",
        "4, 'ART-004', 'Letër A4 500 fletë', 2, 28",
        "5, 'ART-005', 'Fletore A5', 2, 35",
        "6, 'ART-006', 'Letër ngjitëse', 2, 2",
        "7, 'ART-007', 'Maus pa tel', 3, 9",
        "8, 'ART-008', 'Tastierë USB', 3, 16",
        "9, 'ART-009', 'Kabllo UTP 5 m', 4, 23",
        "10, 'ART-010', 'Switch 8 porta', 4, 30",
        "11, 'ART-011', 'Lidhës RJ45', 4, 37",
        "12, 'ART-012', 'Detergjent 1 l', 5, 4",
        "13, 'ART-013', 'Peceta letre', 5, 11",
        "14, 'ART-014', 'Karrige zyre', 6, 18",
        "15, 'ART-015', 'Raft metalik', 6, 25",
        "16, 'ART-016', 'Zgjatues 5 prizë', 7, 32",
        "17, 'ART-017', 'Llambë LED 9 W', 7, 39",
        "18, 'ART-018', 'Bateri AA 4 copë', 7, 6",
        "19, 'ART-019', 'Dosje me kapëse', 8, 13",
        "20, 'ART-020', 'Kuti arkivi', 8, 20"
    )
} catch { Ndalo 6 'Futja e të dhënave' $_ }

# ---------- BLLOKU 7: Pyetësorët e ruajtur (DAO CreateQueryDef) ----------
try {
    # Artikujt me emrin e kategorisë dhe të furnitorit (LEFT JOIN: artikulli mund të jetë pa furnitor)
    New-AccQuery 'Q_Artikujt_Plote' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Furnitoret.EmriFurnitorit, Artikujt.Cmimi, Artikujt.Sasia, Artikujt.MinSasia FROM (Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID) LEFT JOIN Furnitoret ON Artikujt.FurnitoriID = Furnitoret.FurnitoriID ORDER BY Artikujt.Kodi;"
    # Artikujt që duhen porositur (Sasia <= MinSasia)
    New-AccQuery 'Q_NenMinimum' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Artikujt.Sasia, Artikujt.MinSasia FROM Artikujt WHERE Artikujt.Sasia <= Artikujt.MinSasia ORDER BY Artikujt.Kodi;"

    # Kontrolli i kategorisë A për çdo kompjuter
    New-AccQuery 'Q_Kontroll_K1' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 1 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K2' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 2 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K3' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 3 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K4' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 4 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K5' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 5 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K6' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 6 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K7' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 7 ORDER BY Artikujt.Kodi;"
    New-AccQuery 'Q_Kontroll_K8' "SELECT Artikujt.Kodi, Artikujt.Emertimi, Kategorite.EmriKategorise, Artikujt.Cmimi, Artikujt.Sasia FROM Artikujt INNER JOIN Kategorite ON Artikujt.KategoriaID = Kategorite.KategoriaID WHERE Artikujt.KategoriaID = 8 ORDER BY Artikujt.Kodi;"
} catch { Ndalo 7 'Krijimi i pyetësorëve' $_ }

# ---------- BLLOKU 8: Verifikimi ----------
try {
    Write-Host ''
    Write-Host '=== VERIFIKIMI ==='
    $Db.TableDefs.Refresh()
    $Db.QueryDefs.Refresh()

    # Vlerat e pritura (të njëjtat si në çelësin e përgjigjeve)
    $PritetTabela = [ordered]@{ 'Emertimet' = 10; 'Kategorite' = 8; 'Furnitoret' = 5; 'Artikujt' = 20; 'Levizjet' = 36; 'Kategorite_Kopje' = 8; 'Artikujt_Kopje' = 20 }
    $PritetRel = [ordered]@{ 'FK_Artikujt_Kategorite' = 0; 'FK_Artikujt_Furnitoret' = 256; 'FK_Levizjet_Artikujt' = 0; 'FK_ArtikujtKopje_KategoriteKopje' = 4352 }      # 256 = Cascade Update, 4096 = Cascade Delete
    $PritetQ = [ordered]@{ 'Q_Artikujt_Plote' = 20; 'Q_NenMinimum' = 5; 'Q_Kontroll_K1' = 3; 'Q_Kontroll_K2' = 3; 'Q_Kontroll_K3' = 2; 'Q_Kontroll_K4' = 3; 'Q_Kontroll_K5' = 2; 'Q_Kontroll_K6' = 2; 'Q_Kontroll_K7' = 3; 'Q_Kontroll_K8' = 2 }
    $ParamShembull = @{}      # P04 nuk ka pyetësorë me parametra
    $PritetNenMin = 5
    $Variantet = @(
        [pscustomobject]@{ K = 1; KategoriaID = 1; Kategoria = 'Shkrim'; A = 3; ArtikulliID = 1; Artikulli = 'ART-001 Stilolaps blu'; B = 3 },
        [pscustomobject]@{ K = 2; KategoriaID = 2; Kategoria = 'Letër'; A = 3; ArtikulliID = 4; Artikulli = 'ART-004 Letër A4 500 fletë'; B = 3 },
        [pscustomobject]@{ K = 3; KategoriaID = 3; Kategoria = 'Kompjuterë'; A = 2; ArtikulliID = 7; Artikulli = 'ART-007 Maus pa tel'; B = 3 },
        [pscustomobject]@{ K = 4; KategoriaID = 4; Kategoria = 'Rrjeta'; A = 3; ArtikulliID = 9; Artikulli = 'ART-009 Kabllo UTP 5 m'; B = 3 },
        [pscustomobject]@{ K = 5; KategoriaID = 5; Kategoria = 'Pastrim'; A = 2; ArtikulliID = 12; Artikulli = 'ART-012 Detergjent 1 l'; B = 3 },
        [pscustomobject]@{ K = 6; KategoriaID = 6; Kategoria = 'Mobilje'; A = 2; ArtikulliID = 14; Artikulli = 'ART-014 Karrige zyre'; B = 3 },
        [pscustomobject]@{ K = 7; KategoriaID = 7; Kategoria = 'Elektrike'; A = 3; ArtikulliID = 16; Artikulli = 'ART-016 Zgjatues 5 prizë'; B = 3 },
        [pscustomobject]@{ K = 8; KategoriaID = 8; Kategoria = 'Arkivim'; A = 2; ArtikulliID = 19; Artikulli = 'ART-019 Dosje me kapëse'; B = 3 }
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

    # 8b. Fushat me tipin, Required (R), Default Value (D) dhe Validation Rule (V)
    foreach ($t in @('Kategorite', 'Furnitoret', 'Artikujt', 'Levizjet')) {
        $tdfX = $Db.TableDefs.Item($t)
        $pjeset = for ($i = 0; $i -lt $tdfX.Fields.Count; $i++) {
            $f = $tdfX.Fields.Item($i)
            $sh = $f.Name + ' ' + (Get-LlojiFushes $f)
            if ($f.Required) { $sh += ' R' }
            if ([string]$f.DefaultValue -ne '') { $sh += ' D=' + $f.DefaultValue }
            if ([string]$f.ValidationRule -ne '') { $sh += ' V=' + $f.ValidationRule }
            $sh
            Release-ComObjekt $f
        }
        Release-ComObjekt $tdfX
        Write-Host ('  ' + $t + ': ' + ($pjeset -join '; '))
    }
    Write-Host ''

    # 8c. Relacionet
    Write-Host 'Relacionet:'
    for ($i = 0; $i -lt $Db.Relations.Count; $i++) {
        $r = $Db.Relations.Item($i)
        if ($r.Name -like 'FK_*') {
            $cascade = 'pa Cascade'
            if (([int]$r.Attributes -band 256) -or ([int]$r.Attributes -band 4096)) { $cascade = 'me Cascade' }
            $rif = 'Enforce Referential Integrity'
            if ([int]$r.Attributes -band 2) { $rif = 'PA integritet referencial' }
            $ana = '(shumë)'
            if ([int]$r.Attributes -band $dbRelationUnique) { $ana = '(1)' }
            if ([int]$r.Attributes -band 256) { $cascade = 'Cascade Update' }
            if ([int]$r.Attributes -band 4096) { $cascade += ' + Cascade Delete' }
            Write-Host ('  ' + $r.Name + ': ' + $r.Table + ' (1) -> ' + $r.ForeignTable + ' ' + $ana + ', ' + $rif + ', ' + $cascade)
        }
        Release-ComObjekt $r
    }
    foreach ($emriRel in $PritetRel.Keys) {
        $gjetRel = $false
        for ($i = 0; $i -lt $Db.Relations.Count; $i++) {
            $r = $Db.Relations.Item($i)
            if ($r.Name -eq $emriRel) {
                $gjetRel = $true
                $pa = [bool]([int]$r.Attributes -band 2)
                $kaskada = [int]$r.Attributes -band ($dbRelUpdCascade + $dbRelDelCascade)
                if ($pa -or ($kaskada -ne $PritetRel[$emriRel])) { $script:Mosperputhje++; Write-Host ('  GABIM: lloji i relacionit ' + $emriRel) -ForegroundColor Red }
            }
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
        $a = Get-Scalar ('SELECT Count(*) FROM Artikujt WHERE KategoriaID = ' + $v.KategoriaID)
        $b = Get-Scalar ('SELECT Count(*) FROM Levizjet WHERE ArtikulliID = ' + $v.ArtikulliID)
        $ok = ($a -eq $v.A) -and ($b -eq $v.B)
        if (-not $ok) { $script:Mosperputhje++ }
        [pscustomobject]@{ K = $v.K; A_Kategoria = $v.Kategoria + ': ' + $a; B_Artikulli = $v.Artikulli + ': ' + $b; Kontrolli = $(if ($ok) { 'OK' } else { 'GABIM' }) }
    }
    Write-Host 'Variantet (A = artikuj në kategori, B = lëvizje të artikullit):'
    $raportV | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 8f. Kontrolle të veçanta
    # 1.4: rregullat e validimit punojnë (provë me rresht të gabuar; rreshti nuk ruhet)
    $provat = @(
        @('Cmimi = 0', "INSERT INTO Artikujt (Kodi, Emertimi, KategoriaID, Cmimi) VALUES ('PROVE-01', 'Provë', 1, 0)"),
        @('Lloji = Shitje', "INSERT INTO Levizjet (ArtikulliID, DataLevizjes, Lloji, Sasia) VALUES (1, #2026-10-01#, 'Shitje', 1)"),
        @('Kodi i dyfishtë', "INSERT INTO Artikujt (Kodi, Emertimi, KategoriaID, Cmimi) VALUES ('ART-001', 'Provë', 1, 1)"),
        @('KategoriaID = 99', "INSERT INTO Artikujt (Kodi, Emertimi, KategoriaID, Cmimi) VALUES ('PROVE-02', 'Provë', 99, 1)")
    )
    foreach ($pr in $provat) {
        $refuzuar = $false
        try { Invoke-DaoSql $pr[1] } catch { $refuzuar = $true }
        if ($refuzuar) { Write-Host ('Provë: ' + $pr[0] + ' -> refuzohet (OK)') -ForegroundColor Green }
        else { $script:Mosperputhje++; Write-Host ('GABIM: ' + $pr[0] + ' u pranua') -ForegroundColor Red }
    }
    $nProve = Get-Scalar "SELECT Count(*) FROM Artikujt WHERE Kodi LIKE 'PROVE*'"
    if ($nProve -ne 0) { $script:Mosperputhje++; Write-Host 'GABIM: një rresht prove mbeti në tabelë.' -ForegroundColor Red }

    # 1.5: Lookup
    $tdfL = $Db.TableDefs.Item('Artikujt'); $fL = $tdfL.Fields.Item('KategoriaID')
    $dc = [int]$fL.Properties.Item('DisplayControl').Value
    $rsrc = [string]$fL.Properties.Item('RowSource').Value
    Release-ComObjekt $fL; Release-ComObjekt $tdfL
    Write-Host ('1.5: Artikujt.KategoriaID DisplayControl = ' + $dc + ' (111 = Combo Box); RowSource = ' + $rsrc)
    if ($dc -ne $acComboBox) { $script:Mosperputhje++ }

    # 1.6: indekset
    $tdfI = $Db.TableDefs.Item('Furnitoret'); $gjetIdx = $false
    for ($i = 0; $i -lt $tdfI.Indexes.Count; $i++) { $ix = $tdfI.Indexes.Item($i); if ($ix.Name -eq 'IX_Furnitoret_EmriQyteti' -and $ix.Unique) { $gjetIdx = $true }; Release-ComObjekt $ix }
    Release-ComObjekt $tdfI
    Write-Host ('1.6: indeksi unik i përbërë IX_Furnitoret_EmriQyteti: ' + $gjetIdx)
    if (-not $gjetIdx) { $script:Mosperputhje++ }

    $nNen = Get-Scalar 'SELECT Count(*) FROM Q_NenMinimum'
    Write-Host ('Q_NenMinimum: ' + $nNen + ' artikuj me Sasia <= MinSasia (pritej ' + $PritetNenMin + ')')
    if ($nNen -ne $PritetNenMin) { $script:Mosperputhje++ }

    $kat2 = [string](Get-Scalar 'SELECT EmriKategorise FROM Kategorite WHERE KategoriaID = 2')
    if ($kat2 -ceq 'Letër') { Write-Host 'Kodimi: shkronjat ë/ç u ruajtën saktë në bazë (KategoriaID 2 = Letër).' -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM kodimi: KategoriaID 2 = ' + $kat2) -ForegroundColor Red }

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
    Write-Host 'Përfundoi. Hap C:\BazatDB\P04\P04_Zgjidhja.accdb në Access.' -ForegroundColor Green
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
