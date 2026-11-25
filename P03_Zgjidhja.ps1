# =====================================================================
# P03_Zgjidhja.ps1  -  "Bazat e të dhënave", Klasa 12, Projekti P03
# Klinika e familjes: rregullat biznesore, importi dhe konvertimi i
# bazave të vjetra
#
# Qëllimi : krijon C:\BazatDB\P03\P03_Zgjidhja.accdb (zgjidhja: tabelat e
#           importuara, rregullat, Required/Default, relacionet, pyetësorët)
#           dhe skedarët burim për nxënësit në C:\BazatDB\P03\Burimet\
#           (pacientet.xlsx, mjeket.csv, vizitat.csv, klinika_vjeter.mdb).
#           Mësimdhënësi e kopjon dosjen P03 me USB në 8 kompjuterët.
#           Nxënësit nuk e ekzekutojnë skriptin.
# Temat   : N08, N09+9.1 (.xlsx, fleta «Plani i orëve», Prioriteti = Thelbësore)
# Kërkesat: Windows PowerShell 5.1; Microsoft Access me ACE OLEDB
#           (16.0 ose 12.0) dhe DAO.DBEngine.120, me të njëjtin bitness
#           si PowerShell-i që e ekzekuton skriptin.
#           Pa Access.Application, pa VBE, pa SendKeys.
# Ruaje skedarin si "UTF-8 with BOM" që shkronjat ë, ç, Ë, Ç të mos prishen.
#
# Ekzekutimi, Office 64-bit:
#   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P03_Zgjidhja.ps1"
# Ekzekutimi, Office 32-bit:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P03_Zgjidhja.ps1"
#
# Rendi: dosja -> kopja rezervë -> baza -> tabelat -> vetitë e fushave ->
#        relacionet -> të dhënat (prindërit para fëmijëve) -> pyetësorët ->
#        skedarët burim -> verifikimi. Kodet e daljes: 0 = në rregull,
#        10-19 = gabim në bllokun 0-9, 99 = gabim i papritur.
# =====================================================================

$ErrorActionPreference = 'Stop'

$Folder   = 'C:\BazatDB\P03'
$DbPath   = Join-Path $Folder 'P03_Zgjidhja.accdb'
$LockPath = Join-Path $Folder 'P03_Zgjidhja.laccdb'
$Burimet  = Join-Path $Folder 'Burimet'

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

# Krijon me DAO një relacion me një fushë. Atributet 0 = një-me-shumë, 1 = një-me-një;
# të dyja me Enforce Referential Integrity dhe pa Cascade.
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
    Write-Host ('U krijua relacioni ' + $Emri + ': ' + $Tabela + '.' + $Fusha + ' (1) -> ' + $TabelaHuaj + '.' + $FushaHuaj + ' ' + $ana)
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
    Write-Host '=== P03_Zgjidhja.ps1 - Klinika e familjes (rregullat dhe importi) ==='
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
        $EmriRezerve = 'P03_Zgjidhja_' + $Vula + '.bak.accdb'
        $nr = 2
        while (Test-Path -LiteralPath (Join-Path $Folder $EmriRezerve)) {
            $EmriRezerve = 'P03_Zgjidhja_' + $Vula + '_' + $nr + '.bak.accdb'
            $nr++
        }
        Rename-Item -LiteralPath $DbPath -NewName $EmriRezerve
        Write-Host ('Baza e vjetër u ruajt si kopje rezervë: ' + $EmriRezerve)
    }
    # Skedarët burim: dosja e vjetër Burimet ruhet me vulë kohe, pastaj krijohet e reja
    if (Test-Path -LiteralPath $Burimet) {
        $EmriB = 'Burimet_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.bak'
        Rename-Item -LiteralPath $Burimet -NewName $EmriB
        Write-Host ('Dosja e vjetër Burimet u ruajt si ' + $EmriB)
    }
    New-Item -ItemType Directory -Path $Burimet -Force | Out-Null
    Write-Host ('U krijua dosja ' + $Burimet)
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
# P03 nuk ka Hyperlink, Attachment ose Calculated.
try {
    $Conn = New-Object -ComObject ADODB.Connection
    $Conn.Open('Provider=' + $Provider + ';Data Source=' + $DbPath + ';')

    # N08: rregullat biznesore të klinikës dhe mekanizmi që i zbaton
    New-AccTable 'Rregullat' 'CREATE TABLE Rregullat (RregullaID COUNTER CONSTRAINT PK_Rregullat PRIMARY KEY, Rregulla TEXT(100), Tabela TEXT(20), Fusha TEXT(40), Mekanizmi TEXT(60), Projekti TEXT(3))'

    # N09+9.1: tabelat që nxënësit i importojnë. NOT NULL = Required = Yes (N08).
    # PacientiID dhe MjekuID vijnë nga skedarët burim, prandaj janë LONG (jo AutoNumber).
    New-AccTable 'Pacientet' 'CREATE TABLE Pacientet (PacientiID LONG CONSTRAINT PK_Pacientet PRIMARY KEY, Emri TEXT(30) NOT NULL, Mbiemri TEXT(30) NOT NULL, DataLindjes DATETIME, Telefoni TEXT(15), Qyteti TEXT(30))'
    New-AccTable 'Mjeket' 'CREATE TABLE Mjeket (MjekuID LONG CONSTRAINT PK_Mjeket PRIMARY KEY, EmriMjekut TEXT(30), MbiemriMjekut TEXT(30), Specializimi TEXT(30))'
    New-AccTable 'Vizitat' 'CREATE TABLE Vizitat (VizitaID LONG CONSTRAINT PK_Vizitat PRIMARY KEY, PacientiID LONG NOT NULL, MjekuID LONG NOT NULL, DataVizites DATETIME NOT NULL, Arsyeja TEXT(100))'
    # N09: tabela që importohet nga baza e vjetër klinika_vjeter.mdb
    New-AccTable 'Barnat' 'CREATE TABLE Barnat (BarnaID COUNTER CONSTRAINT PK_Barnat PRIMARY KEY, EmriBarnes TEXT(40), Forma TEXT(20), Cmimi CURRENCY, NeStok YESNO)'

    $Conn.Close()
    Release-ComObjekt $Conn
    $Conn = $null
    Start-Sleep -Milliseconds 1000      # pritje e shkurtër që ndryshimet të shkruhen në disk
} catch { Ndalo 3 'Krijimi i tabelave me ADODB (DDL)' $_ }

# ---------- BLLOKU 4: Vetitë e fushave dhe të tabelave (DAO) ----------
try {
    $Db = $DbE.OpenDatabase($DbPath, $false, $false)
    $Db.TableDefs.Refresh()

    # N08: Default Value (Required = Yes u vendos me NOT NULL në BLLOKUN 3)
    Set-FieldProp 'Vizitat' 'DataVizites' 'DefaultValue' $dbText 'Date()'
    Set-FieldProp 'Mjeket' 'Specializimi' 'DefaultValue' $dbText '"Mjek familjeje"'

    # Formati i shfaqjes (Kosova: dd.mm.yyyy; euro me 2 shifra dhjetore)
    Set-FieldProp 'Pacientet' 'DataLindjes' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Vizitat' 'DataVizites' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Barnat' 'Cmimi' 'Format' $dbText 'Euro'
    Set-FieldProp 'Barnat' 'Cmimi' 'DecimalPlaces' $dbByte 2
    Set-FieldProp 'Barnat' 'NeStok' 'Format' $dbText 'Yes/No'
    Set-FieldProp 'Barnat' 'NeStok' 'DisplayControl' $dbInteger $acCheckBox

    # Nëntabelat (+): pacienti dhe mjeku tregojnë vizitat e tyre
    $Nentabelat = @(
        @('Pacientet', 'Table.Vizitat', 'PacientiID'),
        @('Mjeket', 'Table.Vizitat', 'MjekuID')
    )
    foreach ($nt in $Nentabelat) {
        try {
            Set-TableProp $nt[0] 'SubdatasheetName' $dbText $nt[1]
            Set-TableProp $nt[0] 'LinkChildFields' $dbText $nt[2]
            Set-TableProp $nt[0] 'LinkMasterFields' $dbText $nt[2]
        } catch {
            Add-Paralajmerim ('Nëntabela e ' + $nt[0] + ' nuk u vendos: ' + (Get-MesazhiGabimit $_) +
                              ' Vendose me dorë: ' + $nt[0] + ' -> Design View -> Property Sheet -> Subdatasheet Name = ' + $nt[1] + '.')
        }
    }
} catch { Ndalo 4 'Vetitë e fushave dhe të tabelave (DAO)' $_ }

# ---------- BLLOKU 5: Relacionet (DAO) ----------
# N08: «pacienti dhe mjeku i vizitës duhet të ekzistojnë» -> Enforce Referential
# Integrity, pa Cascade (Cascade dhe testet janë te P04).
try {
    New-AccRelation 'FK_Vizitat_Pacientet' 'Pacientet' 'Vizitat' 'PacientiID' 'PacientiID' 0
    New-AccRelation 'FK_Vizitat_Mjeket' 'Mjeket' 'Vizitat' 'MjekuID' 'MjekuID' 0
} catch { Ndalo 5 'Krijimi i relacioneve (DAO)' $_ }

# ---------- BLLOKU 6: Të dhënat shembull (prindërit para fëmijëve) ----------
# Të dhëna fiktive. Datat #yyyy-mm-dd#, numrat dhjetorë me pikë, apostrofi shkruhet ''.
try {
    # Rregullat (N08): 8 rregulla biznesore; 7 zbatohen në P03, 1 te P04
    Add-Rreshta 'Rregullat' 'RregullaID, Rregulla, Tabela, Fusha, Mekanizmi, Projekti' @(
        "1, 'Çdo pacient ka një numër unik', 'Pacientet', 'PacientiID', 'Primary Key', 'P03'",
        "2, 'Emri dhe mbiemri i pacientit janë të detyrueshëm', 'Pacientet', 'Emri, Mbiemri', 'Required = Yes', 'P03'",
        "3, 'Çdo vizitë ka një pacient dhe një mjek', 'Vizitat', 'PacientiID, MjekuID', 'Required = Yes', 'P03'",
        "4, 'Pacienti dhe mjeku i vizitës duhet të ekzistojnë', 'Vizitat', 'PacientiID, MjekuID', 'Lidhje me Enforce Referential Integrity', 'P03'",
        "5, 'Data e vizitës është e detyrueshme; si parazgjedhje është data e sotme', 'Vizitat', 'DataVizites', 'Required = Yes; Default Value = Date()', 'P03'",
        "6, 'Mjeku i ri është mjek familjeje, nëse nuk thuhet ndryshe', 'Mjeket', 'Specializimi', 'Default Value = «Mjek familjeje»', 'P03'",
        "7, 'Çmimi i barit është më i madh se 0', 'Barnat', 'Cmimi', 'Validation Rule (P04)', 'P04'",
        "8, 'Telefoni ruhet me zero në fillim', 'Pacientet', 'Telefoni', 'Short Text (jo Number)', 'P03'"
    )

    # Pacientet (prind): 20 pacientë fiktivë, si pas importit nga pacientet.xlsx
    Add-Rreshta 'Pacientet' 'PacientiID, Emri, Mbiemri, DataLindjes, Telefoni, Qyteti' @(
        "1, 'Albina', 'Krasniqi', #1957-06-04#, '045 201 311', 'Prizren'",
        "2, 'Bekim', 'Gashi', #1964-11-07#, '046 202 322', 'Prizren'",
        "3, 'Dafina', 'Hoxha', #1971-04-10#, '044 203 333', 'Prizren'",
        "4, 'Edon', 'Berisha', #1978-09-13#, '045 204 344', 'Gjakovë'",
        "5, 'Flaka', 'Morina', #1985-02-16#, '046 205 355', 'Gjakovë'",
        "6, 'Genc', 'Shala', #1992-07-19#, '044 206 366', 'Gjakovë'",
        "7, 'Hira', 'Kelmendi', #1999-12-22#, '045 207 377', 'Rahovec'",
        "8, 'Ilir', 'Rexhepi', #2006-05-25#, '046 208 388', 'Rahovec'",
        "9, 'Jehona', 'Bytyqi', #1953-10-28#, '044 209 399', 'Rahovec'",
        "10, 'Kreshnik', 'Hasani', #1960-03-03#, '045 210 410', 'Suharekë'",
        "11, 'Lule', 'Mustafa', #1967-08-06#, '046 211 421', 'Suharekë'",
        "12, 'Mal', 'Kryeziu', #1974-01-09#, '044 212 432', 'Malishevë'",
        "13, 'Nora', 'Rama', #1981-06-12#, '045 213 443', 'Malishevë'",
        "14, 'Petrit', 'Sylejmani', #1988-11-15#, '046 214 454', 'Pejë'",
        "15, 'Qëndresa', 'Ahmeti', #1995-04-18#, '044 215 465', 'Pejë'",
        "16, 'Rrezart', 'Bajrami', #2002-09-21#, '045 216 476', 'Pejë'",
        "17, 'Shqipe', 'Lleshi', #2009-02-24#, '046 217 487', 'Ferizaj'",
        "18, 'Tringa', 'Zeqiri', #1956-07-27#, '044 218 498', 'Ferizaj'",
        "19, 'Ujkan', 'Avdiu', #1963-12-02#, '045 219 509', 'Gjilan'",
        "20, 'Vesa', 'Dema', #1970-05-05#, '046 220 520', 'Gjilan'"
    )

    # Mjeket (prind): 8 mjekë fiktivë, si pas importit nga mjeket.csv
    Add-Rreshta 'Mjeket' 'MjekuID, EmriMjekut, MbiemriMjekut, Specializimi' @(
        "1, 'Arta', 'Gjoka', 'Mjek familjeje'",
        "2, 'Besnik', 'Qorraj', 'Mjek familjeje'",
        "3, 'Drita', 'Lumi', 'Pediatër'",
        "4, 'Fatos', 'Nushi', 'Kardiolog'",
        "5, 'Gentiana', 'Rugova', 'Mjek familjeje'",
        "6, 'Ilirjana', 'Sopi', 'Pediatër'",
        "7, 'Kujtim', 'Tahiri', 'Dermatolog'",
        "8, 'Liridona', 'Veseli', 'Gjinekologe'"
    )

    # Vizitat (fëmijë): 24 vizita, si pas importit nga vizitat.csv dhe korrigjimit të vizitës 17
    Add-Rreshta 'Vizitat' 'VizitaID, PacientiID, MjekuID, DataVizites, Arsyeja' @(
        "1, 1, 1, #2026-09-14#, 'Vaksinim'",
        "2, 1, 2, #2026-09-28#, 'Analiza gjaku'",
        "3, 4, 3, #2026-10-12#, 'Temperaturë dhe kollë'",
        "4, 4, 4, #2026-10-23#, 'Dhimbje koke'",
        "5, 7, 5, #2026-11-05#, 'Ekzaminim i lëkurës'",
        "6, 7, 6, #2026-11-18#, 'Matje e tensionit'",
        "7, 10, 7, #2026-09-07#, 'Kontroll pas terapisë'",
        "8, 10, 8, #2026-09-21#, 'Kontroll i rregullt'",
        "9, 12, 1, #2026-10-05#, 'Vaksinim'",
        "10, 12, 2, #2026-10-16#, 'Analiza gjaku'",
        "11, 14, 3, #2026-10-29#, 'Temperaturë dhe kollë'",
        "12, 14, 4, #2026-11-11#, 'Dhimbje koke'",
        "13, 17, 5, #2026-11-24#, 'Ekzaminim i lëkurës'",
        "14, 17, 6, #2026-09-14#, 'Matje e tensionit'",
        "15, 19, 7, #2026-09-28#, 'Kontroll pas terapisë'",
        "16, 19, 8, #2026-10-09#, 'Kontroll i rregullt'",
        "17, 2, 1, #2026-11-05#, 'Vaksinim'",
        "18, 3, 2, #2026-11-04#, 'Analiza gjaku'",
        "19, 5, 3, #2026-11-17#, 'Temperaturë dhe kollë'",
        "20, 6, 4, #2026-09-07#, 'Dhimbje koke'",
        "21, 8, 5, #2026-09-21#, 'Ekzaminim i lëkurës'",
        "22, 9, 6, #2026-10-02#, 'Matje e tensionit'",
        "23, 11, 7, #2026-10-15#, 'Kontroll pas terapisë'",
        "24, 13, 8, #2026-10-28#, 'Kontroll i rregullt'"
    )

    # Barnat: 10 barna, si pas importit nga klinika_vjeter.mdb
    Add-Rreshta 'Barnat' 'BarnaID, EmriBarnes, Forma, Cmimi, NeStok' @(
        "1, 'Paracetamol 500 mg', 'Tableta', 2.40, True",
        "2, 'Ibuprofen 400 mg', 'Tableta', 3.10, True",
        "3, 'Amoksicilinë 500 mg', 'Kapsula', 5.80, False",
        "4, 'Shurup për kollë', 'Shurup', 4.50, True",
        "5, 'Krem hidratues', 'Krem', 6.90, True",
        "6, 'Vitaminë C', 'Tableta', 3.95, True",
        "7, 'Pika për sy', 'Pika', 4.20, False",
        "8, 'Losartan 50 mg', 'Tableta', 7.30, True",
        "9, 'Sprej për hundë', 'Sprej', 5.15, True",
        "10, 'Pomadë Eozinë', 'Pomadë', 2.85, False"
    )
} catch { Ndalo 6 'Futja e të dhënave' $_ }

# ---------- BLLOKU 7: Pyetësorët e ruajtur (DAO CreateQueryDef) ----------
try {
    # Tabelat e importuara, të bashkuara përmes lidhjeve
    New-AccQuery 'Q_Vizitat_Plote' "SELECT Vizitat.VizitaID, Vizitat.DataVizites, Pacientet.Emri, Pacientet.Mbiemri, Pacientet.Qyteti, Mjeket.MbiemriMjekut, Mjeket.Specializimi, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID ORDER BY Vizitat.VizitaID;"

    # Kontrolli i vizitave të mjekut B për çdo kompjuter (Niveli 2)
    New-AccQuery 'Q_Kontroll_K1' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 1 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K2' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 2 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K3' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 3 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K4' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 4 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K5' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 5 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K6' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 6 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K7' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 7 ORDER BY Vizitat.DataVizites;"
    New-AccQuery 'Q_Kontroll_K8' "SELECT Mjeket.EmriMjekut, Mjeket.MbiemriMjekut, Vizitat.VizitaID, Pacientet.Emri, Pacientet.Mbiemri, Vizitat.DataVizites, Vizitat.Arsyeja FROM (Vizitat INNER JOIN Pacientet ON Vizitat.PacientiID = Pacientet.PacientiID) INNER JOIN Mjeket ON Vizitat.MjekuID = Mjeket.MjekuID WHERE Vizitat.MjekuID = 8 ORDER BY Vizitat.DataVizites;"
} catch { Ndalo 7 'Krijimi i pyetësorëve' $_ }

# ---------- BLLOKU 8: Skedarët burim për nxënësit (C:\BazatDB\P03\Burimet) ----------
# N09+9.1: Excel, CSV dhe një bazë e vjetër .mdb. Të njëjtat të dhëna si në zgjidhje,
# por vizitat.csv ka me qëllim një datë të gabuar (vizita 17) që të dalë tabela Import Errors.
try {
    $Utf8Bom = New-Object System.Text.UTF8Encoding($true)

    # mjeket.csv: tekstet në thonjëza, presje si ndarës, UTF-8
    $RreshtatMjeket = [string[]]@(
        'MjekuID,EmriMjekut,MbiemriMjekut,Specializimi',
        '1,"Arta","Gjoka","Mjek familjeje"',
        '2,"Besnik","Qorraj","Mjek familjeje"',
        '3,"Drita","Lumi","Pediatër"',
        '4,"Fatos","Nushi","Kardiolog"',
        '5,"Gentiana","Rugova","Mjek familjeje"',
        '6,"Ilirjana","Sopi","Pediatër"',
        '7,"Kujtim","Tahiri","Dermatolog"',
        '8,"Liridona","Veseli","Gjinekologe"'
    )
    [System.IO.File]::WriteAllLines((Join-Path $Burimet 'mjeket.csv'), $RreshtatMjeket, $Utf8Bom)
    Write-Host ('U krijua mjeket.csv (' + ($RreshtatMjeket.Count - 1) + ' rreshta)')

    # vizitat.csv: datat yyyy-mm-dd; vizita 17 ka datën 2026-13-05 (muaji 13 nuk ekziston)
    $RreshtatVizitat = [string[]]@(
        'VizitaID,PacientiID,MjekuID,DataVizites,Arsyeja',
        '1,1,1,2026-09-14,"Vaksinim"',
        '2,1,2,2026-09-28,"Analiza gjaku"',
        '3,4,3,2026-10-12,"Temperaturë dhe kollë"',
        '4,4,4,2026-10-23,"Dhimbje koke"',
        '5,7,5,2026-11-05,"Ekzaminim i lëkurës"',
        '6,7,6,2026-11-18,"Matje e tensionit"',
        '7,10,7,2026-09-07,"Kontroll pas terapisë"',
        '8,10,8,2026-09-21,"Kontroll i rregullt"',
        '9,12,1,2026-10-05,"Vaksinim"',
        '10,12,2,2026-10-16,"Analiza gjaku"',
        '11,14,3,2026-10-29,"Temperaturë dhe kollë"',
        '12,14,4,2026-11-11,"Dhimbje koke"',
        '13,17,5,2026-11-24,"Ekzaminim i lëkurës"',
        '14,17,6,2026-09-14,"Matje e tensionit"',
        '15,19,7,2026-09-28,"Kontroll pas terapisë"',
        '16,19,8,2026-10-09,"Kontroll i rregullt"',
        '17,2,1,2026-13-05,"Vaksinim"',
        '18,3,2,2026-11-04,"Analiza gjaku"',
        '19,5,3,2026-11-17,"Temperaturë dhe kollë"',
        '20,6,4,2026-09-07,"Dhimbje koke"',
        '21,8,5,2026-09-21,"Ekzaminim i lëkurës"',
        '22,9,6,2026-10-02,"Matje e tensionit"',
        '23,11,7,2026-10-15,"Kontroll pas terapisë"',
        '24,13,8,2026-10-28,"Kontroll i rregullt"'
    )
    [System.IO.File]::WriteAllLines((Join-Path $Burimet 'vizitat.csv'), $RreshtatVizitat, $Utf8Bom)
    Write-Host ('U krijua vizitat.csv (' + ($RreshtatVizitat.Count - 1) + ' rreshta)')

    # klinika_vjeter.mdb: baza e vjetër (format Access 2002-2003, Jet 4) me tabelën Barnat
    $MdbPath = Join-Path $Burimet 'klinika_vjeter.mdb'
    $CatV = New-Object -ComObject ADOX.Catalog
    try {
        [void]$CatV.Create('Provider=' + $Provider + ';Data Source=' + $MdbPath + ';Jet OLEDB:Engine Type=5;')
        try { $CatV.ActiveConnection.Close() } catch { }
    } finally { Release-ComObjekt $CatV; $CatV = $null }
    $cV = New-Object -ComObject ADODB.Connection
    try {
        $cV.Open('Provider=' + $Provider + ';Data Source=' + $MdbPath + ';')
        [void]$cV.Execute('CREATE TABLE Barnat (BarnaID COUNTER CONSTRAINT PK_Barnat PRIMARY KEY, EmriBarnes TEXT(40), Forma TEXT(20), Cmimi CURRENCY, NeStok YESNO)')
        foreach ($vl in [string[]]@(
            '1, ''Paracetamol 500 mg'', ''Tableta'', 2.40, True',
            '2, ''Ibuprofen 400 mg'', ''Tableta'', 3.10, True',
            '3, ''Amoksicilinë 500 mg'', ''Kapsula'', 5.80, False',
            '4, ''Shurup për kollë'', ''Shurup'', 4.50, True',
            '5, ''Krem hidratues'', ''Krem'', 6.90, True',
            '6, ''Vitaminë C'', ''Tableta'', 3.95, True',
            '7, ''Pika për sy'', ''Pika'', 4.20, False',
            '8, ''Losartan 50 mg'', ''Tableta'', 7.30, True',
            '9, ''Sprej për hundë'', ''Sprej'', 5.15, True',
            '10, ''Pomadë Eozinë'', ''Pomadë'', 2.85, False'
        )) {
            [void]$cV.Execute('INSERT INTO Barnat (BarnaID, EmriBarnes, Forma, Cmimi, NeStok) VALUES (' + $vl + ')')
        }
    } finally { if ($cV.State -ne 0) { $cV.Close() }; Release-ComObjekt $cV; $cV = $null }
    Write-Host 'U krijua klinika_vjeter.mdb (tabela Barnat, 10 rreshta)'

    # pacientet.xlsx: krijohet me ofruesin ACE (Excel 12.0 Xml), pa Excel.Application.
    # Nëse nuk krijohet, skripti shkruan pacientet_rezerve.csv dhe vazhdon me paralajmërim.
    $XlsxPath = Join-Path $Burimet 'pacientet.xlsx'
    $LidhjaXl = 'Provider=' + $Provider + ';Data Source=' + $XlsxPath + ';Extended Properties="Excel 12.0 Xml;HDR=YES";'
    $cX = $null; $CatX = $null
    try {
        $CatX = New-Object -ComObject ADOX.Catalog
        [void]$CatX.Create($LidhjaXl)
        try { $CatX.ActiveConnection.Close() } catch { }
        $cX = New-Object -ComObject ADODB.Connection
        $cX.Open($LidhjaXl)
        [void]$cX.Execute('CREATE TABLE Pacientet (PacientiID NUMBER, Emri TEXT, Mbiemri TEXT, DataLindjes DATETIME, Telefoni TEXT, Qyteti TEXT)')
        foreach ($vl in [string[]]@(
            '1, ''Albina'', ''Krasniqi'', #1957-06-04#, ''045 201 311'', ''Prizren''',
            '2, ''Bekim'', ''Gashi'', #1964-11-07#, ''046 202 322'', ''Prizren''',
            '3, ''Dafina'', ''Hoxha'', #1971-04-10#, ''044 203 333'', ''Prizren''',
            '4, ''Edon'', ''Berisha'', #1978-09-13#, ''045 204 344'', ''Gjakovë''',
            '5, ''Flaka'', ''Morina'', #1985-02-16#, ''046 205 355'', ''Gjakovë''',
            '6, ''Genc'', ''Shala'', #1992-07-19#, ''044 206 366'', ''Gjakovë''',
            '7, ''Hira'', ''Kelmendi'', #1999-12-22#, ''045 207 377'', ''Rahovec''',
            '8, ''Ilir'', ''Rexhepi'', #2006-05-25#, ''046 208 388'', ''Rahovec''',
            '9, ''Jehona'', ''Bytyqi'', #1953-10-28#, ''044 209 399'', ''Rahovec''',
            '10, ''Kreshnik'', ''Hasani'', #1960-03-03#, ''045 210 410'', ''Suharekë''',
            '11, ''Lule'', ''Mustafa'', #1967-08-06#, ''046 211 421'', ''Suharekë''',
            '12, ''Mal'', ''Kryeziu'', #1974-01-09#, ''044 212 432'', ''Malishevë''',
            '13, ''Nora'', ''Rama'', #1981-06-12#, ''045 213 443'', ''Malishevë''',
            '14, ''Petrit'', ''Sylejmani'', #1988-11-15#, ''046 214 454'', ''Pejë''',
            '15, ''Qëndresa'', ''Ahmeti'', #1995-04-18#, ''044 215 465'', ''Pejë''',
            '16, ''Rrezart'', ''Bajrami'', #2002-09-21#, ''045 216 476'', ''Pejë''',
            '17, ''Shqipe'', ''Lleshi'', #2009-02-24#, ''046 217 487'', ''Ferizaj''',
            '18, ''Tringa'', ''Zeqiri'', #1956-07-27#, ''044 218 498'', ''Ferizaj''',
            '19, ''Ujkan'', ''Avdiu'', #1963-12-02#, ''045 219 509'', ''Gjilan''',
            '20, ''Vesa'', ''Dema'', #1970-05-05#, ''046 220 520'', ''Gjilan'''
        )) {
            [void]$cX.Execute('INSERT INTO [Pacientet] (PacientiID, Emri, Mbiemri, DataLindjes, Telefoni, Qyteti) VALUES (' + $vl + ')')
        }
        Write-Host 'U krijua pacientet.xlsx (fleta Pacientet, 20 rreshta)'
    } catch {
        $RreshtatPacientet = [string[]]@(
            'PacientiID,Emri,Mbiemri,DataLindjes,Telefoni,Qyteti',
            '1,"Albina","Krasniqi",1957-06-04,"045 201 311","Prizren"',
            '2,"Bekim","Gashi",1964-11-07,"046 202 322","Prizren"',
            '3,"Dafina","Hoxha",1971-04-10,"044 203 333","Prizren"',
            '4,"Edon","Berisha",1978-09-13,"045 204 344","Gjakovë"',
            '5,"Flaka","Morina",1985-02-16,"046 205 355","Gjakovë"',
            '6,"Genc","Shala",1992-07-19,"044 206 366","Gjakovë"',
            '7,"Hira","Kelmendi",1999-12-22,"045 207 377","Rahovec"',
            '8,"Ilir","Rexhepi",2006-05-25,"046 208 388","Rahovec"',
            '9,"Jehona","Bytyqi",1953-10-28,"044 209 399","Rahovec"',
            '10,"Kreshnik","Hasani",1960-03-03,"045 210 410","Suharekë"',
            '11,"Lule","Mustafa",1967-08-06,"046 211 421","Suharekë"',
            '12,"Mal","Kryeziu",1974-01-09,"044 212 432","Malishevë"',
            '13,"Nora","Rama",1981-06-12,"045 213 443","Malishevë"',
            '14,"Petrit","Sylejmani",1988-11-15,"046 214 454","Pejë"',
            '15,"Qëndresa","Ahmeti",1995-04-18,"044 215 465","Pejë"',
            '16,"Rrezart","Bajrami",2002-09-21,"045 216 476","Pejë"',
            '17,"Shqipe","Lleshi",2009-02-24,"046 217 487","Ferizaj"',
            '18,"Tringa","Zeqiri",1956-07-27,"044 218 498","Ferizaj"',
            '19,"Ujkan","Avdiu",1963-12-02,"045 219 509","Gjilan"',
            '20,"Vesa","Dema",1970-05-05,"046 220 520","Gjilan"'
        )
        [System.IO.File]::WriteAllLines((Join-Path $Burimet 'pacientet_rezerve.csv'), $RreshtatPacientet, $Utf8Bom)
        Add-Paralajmerim ('pacientet.xlsx nuk u krijua: ' + (Get-MesazhiGabimit $_) + ' U krijua pacientet_rezerve.csv. Hape në Excel -> File -> Save As -> Excel Workbook (*.xlsx) -> pacientet.xlsx, fleta Pacientet.')
    } finally {
        if ($null -ne $cX) { try { if ($cX.State -ne 0) { $cX.Close() } } catch { } }
        Release-ComObjekt $cX; Release-ComObjekt $CatX; $cX = $null; $CatX = $null
    }
} catch { Ndalo 8 'Krijimi i skedarëve burim' $_ }

# ---------- BLLOKU 9: Verifikimi ----------
try {
    Write-Host ''
    Write-Host '=== VERIFIKIMI ==='
    $Db.TableDefs.Refresh()
    $Db.QueryDefs.Refresh()

    # Vlerat e pritura (të njëjtat si në çelësin e përgjigjeve)
    $PritetTabela = [ordered]@{ 'Rregullat' = 8; 'Pacientet' = 20; 'Mjeket' = 8; 'Vizitat' = 24; 'Barnat' = 10 }
    $PritetRel = [ordered]@{ 'FK_Vizitat_Pacientet' = $false; 'FK_Vizitat_Mjeket' = $false }      # $true = një-me-një
    $PritetQ = [ordered]@{ 'Q_Vizitat_Plote' = 24; 'Q_Kontroll_K1' = 3; 'Q_Kontroll_K2' = 3; 'Q_Kontroll_K3' = 3; 'Q_Kontroll_K4' = 3; 'Q_Kontroll_K5' = 3; 'Q_Kontroll_K6' = 3; 'Q_Kontroll_K7' = 3; 'Q_Kontroll_K8' = 3 }
    $ParamShembull = @{}      # P03 nuk ka pyetësorë me parametra
    $PritetRregullaP03 = 7
    $VizitaKorrigjuar = 17
    $PritetBurime = @(
        [pscustomobject]@{ Skedari = 'pacientet.xlsx'; Lloji = 'xlsx'; Rreshta = 20; Opsional = $true },
        [pscustomobject]@{ Skedari = 'mjeket.csv'; Lloji = 'csv'; Rreshta = 8; Opsional = $false },
        [pscustomobject]@{ Skedari = 'vizitat.csv'; Lloji = 'csv'; Rreshta = 24; Opsional = $false },
        [pscustomobject]@{ Skedari = 'klinika_vjeter.mdb'; Lloji = 'mdb'; Rreshta = 10; Opsional = $false }
    )
    $Variantet = @(
        [pscustomobject]@{ K = 1; Qyteti = 'Prizren'; A = 3; MjekuID = 1; Mjeku = 'Arta Gjoka'; B = 3; PacientiID = 1; Pacienti = 'Albina Krasniqi'; C = 2 },
        [pscustomobject]@{ K = 2; Qyteti = 'Gjakovë'; A = 3; MjekuID = 2; Mjeku = 'Besnik Qorraj'; B = 3; PacientiID = 4; Pacienti = 'Edon Berisha'; C = 2 },
        [pscustomobject]@{ K = 3; Qyteti = 'Rahovec'; A = 3; MjekuID = 3; Mjeku = 'Drita Lumi'; B = 3; PacientiID = 7; Pacienti = 'Hira Kelmendi'; C = 2 },
        [pscustomobject]@{ K = 4; Qyteti = 'Suharekë'; A = 2; MjekuID = 4; Mjeku = 'Fatos Nushi'; B = 3; PacientiID = 10; Pacienti = 'Kreshnik Hasani'; C = 2 },
        [pscustomobject]@{ K = 5; Qyteti = 'Malishevë'; A = 2; MjekuID = 5; Mjeku = 'Gentiana Rugova'; B = 3; PacientiID = 12; Pacienti = 'Mal Kryeziu'; C = 2 },
        [pscustomobject]@{ K = 6; Qyteti = 'Pejë'; A = 3; MjekuID = 6; Mjeku = 'Ilirjana Sopi'; B = 3; PacientiID = 14; Pacienti = 'Petrit Sylejmani'; C = 2 },
        [pscustomobject]@{ K = 7; Qyteti = 'Ferizaj'; A = 2; MjekuID = 7; Mjeku = 'Kujtim Tahiri'; B = 3; PacientiID = 17; Pacienti = 'Shqipe Lleshi'; C = 2 },
        [pscustomobject]@{ K = 8; Qyteti = 'Gjilan'; A = 2; MjekuID = 8; Mjeku = 'Liridona Veseli'; B = 3; PacientiID = 19; Pacienti = 'Ujkan Avdiu'; C = 2 }
    )

    # 9a. Tabelat: rreshtat, çelësi primar, çelësat e huaj
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

    # 9b. Fushat e tabelave të importuara, me Required (R) dhe Default Value (D)
    foreach ($t in @('Pacientet', 'Mjeket', 'Vizitat', 'Barnat')) {
        $tdfX = $Db.TableDefs.Item($t)
        $pjeset = for ($i = 0; $i -lt $tdfX.Fields.Count; $i++) {
            $f = $tdfX.Fields.Item($i)
            $sh = $f.Name + ' ' + (Get-LlojiFushes $f)
            if ($f.Required) { $sh += ' R' }
            if ([string]$f.DefaultValue -ne '') { $sh += ' D=' + $f.DefaultValue }
            $sh
            Release-ComObjekt $f
        }
        Release-ComObjekt $tdfX
        Write-Host ('  ' + $t + ': ' + ($pjeset -join '; '))
    }
    Write-Host ''

    # 9c. Relacionet
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
                $nje = [bool]([int]$r.Attributes -band $dbRelationUnique)
                $pa = [bool]([int]$r.Attributes -band 2)
                if ($pa -or ($nje -ne $PritetRel[$emriRel])) { $script:Mosperputhje++; Write-Host ('  GABIM: lloji i relacionit ' + $emriRel) -ForegroundColor Red }
            }
            Release-ComObjekt $r
        }
        if (-not $gjetRel) { $script:Mosperputhje++; Write-Host ('  GABIM: mungon relacioni ' + $emriRel) -ForegroundColor Red }
    }

    # 9d. Pyetësorët: lloji dhe numri i rreshtave
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

    # 9e. Rezultatet e 8 varianteve (Kompjuteri 1-8), krahasuar me çelësin e përgjigjeve
    $raportV = foreach ($v in $Variantet) {
        $a = Get-Scalar ("SELECT Count(*) FROM Pacientet WHERE Qyteti = '" + $v.Qyteti + "'")
        $b = Get-Scalar ('SELECT Count(*) FROM Vizitat WHERE MjekuID = ' + $v.MjekuID)
        $c = Get-Scalar ('SELECT Count(*) FROM Vizitat WHERE PacientiID = ' + $v.PacientiID)
        $ok = ($a -eq $v.A) -and ($b -eq $v.B) -and ($c -eq $v.C)
        if (-not $ok) { $script:Mosperputhje++ }
        [pscustomobject]@{
            K          = $v.K
            A_Qyteti   = $v.Qyteti + ': ' + $a
            B_Mjeku    = $v.Mjeku + ': ' + $b
            C_Pacienti = $v.Pacienti + ': ' + $c
            Kontrolli  = $(if ($ok) { 'OK' } else { 'GABIM' })
        }
    }
    Write-Host 'Variantet (A = pacientë në qytet, B = vizita të mjekut, C = vizita të pacientit):'
    $raportV | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 9f. Kontrolle të veçanta
    # N08: Required dhe Default Value
    $kontrolleR = @(@('Pacientet', 'Emri'), @('Pacientet', 'Mbiemri'), @('Vizitat', 'PacientiID'), @('Vizitat', 'MjekuID'), @('Vizitat', 'DataVizites'))
    foreach ($kr in $kontrolleR) {
        $tdfR = $Db.TableDefs.Item($kr[0]); $fR = $tdfR.Fields.Item($kr[1])
        if (-not $fR.Required) { $script:Mosperputhje++; Write-Host ('GABIM N08: ' + $kr[0] + '.' + $kr[1] + ' nuk është Required.') -ForegroundColor Red }
        Release-ComObjekt $fR; Release-ComObjekt $tdfR
    }
    $tdfR = $Db.TableDefs.Item('Vizitat'); $fR = $tdfR.Fields.Item('DataVizites'); $defV = [string]$fR.DefaultValue
    Release-ComObjekt $fR; Release-ComObjekt $tdfR
    $tdfR = $Db.TableDefs.Item('Mjeket'); $fR = $tdfR.Fields.Item('Specializimi'); $defM = [string]$fR.DefaultValue
    Release-ComObjekt $fR; Release-ComObjekt $tdfR
    Write-Host ('N08: Required te 5 fusha; Default Value: Vizitat.DataVizites = ' + $defV + ', Mjeket.Specializimi = ' + $defM)
    if ($defV -ne 'Date()' -or $defM -ne '"Mjek familjeje"') { $script:Mosperputhje++; Write-Host 'GABIM N08: Default Value nuk është vendosur saktë.' -ForegroundColor Red }

    $nP3 = Get-Scalar ('SELECT Count(*) FROM Rregullat WHERE Projekti = ' + "'P03'")
    Write-Host ('N08: rregulla të zbatuara në P03 = ' + $nP3 + ' (pritej ' + $PritetRregullaP03 + ')')
    if ($nP3 -ne $PritetRregullaP03) { $script:Mosperputhje++ }

    $dataK = Get-Scalar ('SELECT DataVizites FROM Vizitat WHERE VizitaID = ' + $VizitaKorrigjuar)
    Write-Host ('N09: vizita ' + $VizitaKorrigjuar + ' (data e gabuar në CSV) ka në zgjidhje datën ' + ([datetime]$dataK).ToString('dd.MM.yyyy'))

    $telZero = Get-Scalar "SELECT Count(*) FROM Pacientet WHERE Telefoni LIKE '0*'"
    Write-Host ('Telefoni si Short Text: ' + $telZero + ' telefona fillojnë me 0 (pritej ' + $PritetTabela['Pacientet'] + ').')
    if ($telZero -ne $PritetTabela['Pacientet']) { $script:Mosperputhje++ }

    $nPlote = Get-Scalar 'SELECT Count(*) FROM Q_Vizitat_Plote'
    if ($nPlote -eq $PritetTabela['Vizitat']) { Write-Host ('Q_Vizitat_Plote bashkon 3 tabelat e importuara: ' + $nPlote + ' rreshta.') -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM: Q_Vizitat_Plote jep ' + $nPlote + ' rreshta.') -ForegroundColor Red }

    $bar = [string](Get-Scalar 'SELECT EmriBarnes FROM Barnat WHERE BarnaID = 3')
    if ($bar -ceq 'Amoksicilinë 500 mg') { Write-Host 'Kodimi: shkronjat ë/ç u ruajtën saktë në bazë (BarnaID 3 = Amoksicilinë 500 mg).' -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM kodimi: BarnaID 3 ka EmriBarnes = ' + $bar) -ForegroundColor Red }

    # Skedarët burim (BLLOKU 8)
    $raportS = foreach ($sb in $PritetBurime) {
        $rr = '-'
        $shtegu = Join-Path $Burimet $sb.Skedari
        if (-not (Test-Path -LiteralPath $shtegu)) {
            if ($sb.Opsional) { [pscustomobject]@{ Skedari = $sb.Skedari; Rreshta = '-'; Kontrolli = 'mungon (shih KUJDES)' }; continue }
            $script:Mosperputhje++
            [pscustomobject]@{ Skedari = $sb.Skedari; Rreshta = '-'; Kontrolli = 'GABIM: mungon' }
            continue
        }
        if ($sb.Lloji -eq 'csv') {
            $rr = @(Get-Content -LiteralPath $shtegu -Encoding UTF8 | Where-Object { $_ -ne '' }).Count - 1
        } elseif ($sb.Lloji -eq 'mdb') {
            $dbV = $DbE.OpenDatabase($shtegu, $false, $true)
            $rsV = $dbV.OpenRecordset('SELECT Count(*) FROM Barnat', $dbOpenSnapshot)
            $rr = [int]$rsV.Fields.Item(0).Value
            $rsV.Close(); Release-ComObjekt $rsV; $dbV.Close(); Release-ComObjekt $dbV
        } elseif ($sb.Lloji -eq 'xlsx') {
            $cX = New-Object -ComObject ADODB.Connection
            try {
                $cX.Open('Provider=' + $Provider + ';Data Source=' + $shtegu + ';Extended Properties="Excel 12.0 Xml;HDR=YES";')
                $rsX = $cX.Execute('SELECT Count(*) FROM [Pacientet$]')
                $rr = [int]$rsX.Fields.Item(0).Value
                $rsX.Close(); Release-ComObjekt $rsX
            } finally { if ($cX.State -ne 0) { $cX.Close() }; Release-ComObjekt $cX }
        }
        $okS = ($rr -eq $sb.Rreshta)
        if (-not $okS) { $script:Mosperputhje++ }
        [pscustomobject]@{ Skedari = $sb.Skedari; Rreshta = $rr; Kontrolli = $(if ($okS) { 'OK' } else { 'GABIM (pritej ' + $sb.Rreshta + ')' }) }
    }
    Write-Host 'Skedarët burim në C:\BazatDB\P03\Burimet:'
    $raportS | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

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
    Write-Host 'Përfundoi. Hap C:\BazatDB\P03\P03_Zgjidhja.accdb në Access.' -ForegroundColor Green
} catch { Ndalo 9 'Verifikimi' $_ }

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
