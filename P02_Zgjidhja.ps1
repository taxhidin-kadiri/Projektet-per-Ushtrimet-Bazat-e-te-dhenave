# =====================================================================
# P02_Zgjidhja.ps1  -  "Bazat e të dhënave", Klasa 12, Projekti P02
# Shkolla: nga diagrama ER te tabelat e lidhura (1:N, M:N, 1:1) dhe
# evoluimi i modeleve të bazave të të dhënave
#
# Qëllimi : krijon C:\BazatDB\P02\P02_Zgjidhja.accdb me tabelat, vetitë
#           e fushave, relacionet, të dhënat shembull dhe pyetësorët e
#           zgjidhur të P02. Mësimdhënësi e kopjon bazën me USB në 8
#           kompjuterët e kabinetit. Nxënësit nuk e ekzekutojnë skriptin.
# Temat   : N06, N07 (.xlsx, fleta «Plani i orëve», Prioriteti = Thelbësore)
# Kërkesat: Windows PowerShell 5.1; Microsoft Access me ACE OLEDB
#           (16.0 ose 12.0) dhe DAO.DBEngine.120, me të njëjtin bitness
#           si PowerShell-i që e ekzekuton skriptin.
#           Pa Access.Application, pa VBE, pa SendKeys.
# Ruaje skedarin si "UTF-8 with BOM" që shkronjat ë, ç, Ë, Ç të mos prishen.
#
# Ekzekutimi, Office 64-bit:
#   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P02_Zgjidhja.ps1"
# Ekzekutimi, Office 32-bit:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\BazatDB\Skriptet\P02_Zgjidhja.ps1"
#
# Rendi: dosja -> kopja rezervë -> baza -> tabelat -> vetitë e fushave ->
#        relacionet -> të dhënat (prindërit para fëmijëve) -> pyetësorët ->
#        verifikimi. Kodet e daljes: 0 = në rregull, 10-18 = gabim në
#        bllokun 0-8, 99 = gabim i papritur.
# =====================================================================

$ErrorActionPreference = 'Stop'

$Folder   = 'C:\BazatDB\P02'
$DbPath   = Join-Path $Folder 'P02_Zgjidhja.accdb'
$LockPath = Join-Path $Folder 'P02_Zgjidhja.laccdb'

$script:KodiDaljes    = 0
$script:Mosperputhje  = 0
$script:Paralajmerime = New-Object System.Collections.ArrayList
$Provider = $null
$Cat      = $null
$Conn     = $null
$DbE      = $null
$Db       = $null

# Konstantet DAO me vlerat numerike (PowerShell nuk i njeh emrat dbText etj.)
$dbText           = 10
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
    Write-Host '=== P02_Zgjidhja.ps1 - Shkolla (modelimi ER) ==='
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
        $EmriRezerve = 'P02_Zgjidhja_' + $Vula + '.bak.accdb'
        $nr = 2
        while (Test-Path -LiteralPath (Join-Path $Folder $EmriRezerve)) {
            $EmriRezerve = 'P02_Zgjidhja_' + $Vula + '_' + $nr + '.bak.accdb'
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
# P02 nuk ka Hyperlink, Attachment ose Calculated.
try {
    $Conn = New-Object -ComObject ADODB.Connection
    $Conn.Open('Provider=' + $Provider + ';Data Source=' + $DbPath + ';')

    # N06: modelet e bazave të të dhënave (vija kohore)
    New-AccTable 'Modelet' 'CREATE TABLE Modelet (ModeliID COUNTER CONSTRAINT PK_Modelet PRIMARY KEY, EmriModelit TEXT(25), Periudha TEXT(12), VitiFillimit SMALLINT, Struktura TEXT(40), Shembulli TEXT(30), Perparesia TEXT(50), Kufizimi TEXT(50))'

    # N07: entitetet e diagramës ER «Shkolla» -> tabela. Çelësi i huaj është LONG
    # (Number, Long Integer), si AutoNumber-i i tabelës prind.
    New-AccTable 'Klasat' 'CREATE TABLE Klasat (KlasaID COUNTER CONSTRAINT PK_Klasat PRIMARY KEY, EmriKlases TEXT(5), Drejtimi TEXT(30), KlasaNr BYTE, Salla TEXT(5))'
    New-AccTable 'Nxenesit' 'CREATE TABLE Nxenesit (NxenesiID COUNTER CONSTRAINT PK_Nxenesit PRIMARY KEY, Emri TEXT(30), Mbiemri TEXT(30), Gjinia TEXT(1), DataLindjes DATETIME, Qyteti TEXT(30), KlasaID LONG)'
    # 1:1: NxenesiID është njëkohësisht çelës primar dhe çelës i huaj
    New-AccTable 'Dosjet' 'CREATE TABLE Dosjet (NxenesiID LONG CONSTRAINT PK_Dosjet PRIMARY KEY, NrAmzes TEXT(10), DataRegjistrimit DATETIME, ShkollaParaprake TEXT(50), TelefoniPrindit TEXT(15))'
    New-AccTable 'Mesuesit' 'CREATE TABLE Mesuesit (MesuesiID COUNTER CONSTRAINT PK_Mesuesit PRIMARY KEY, Emri TEXT(30), Mbiemri TEXT(30), Specializimi TEXT(30), Email TEXT(50))'
    New-AccTable 'Lendet' 'CREATE TABLE Lendet (LendaID COUNTER CONSTRAINT PK_Lendet PRIMARY KEY, EmriLendes TEXT(40), OretJave BYTE, MesuesiID LONG)'
    # M:N Nxenesit-Lendet: tabela ndërmjetëse me dy çelësa të huaj
    New-AccTable 'Notat' 'CREATE TABLE Notat (NotaID COUNTER CONSTRAINT PK_Notat PRIMARY KEY, NxenesiID LONG, LendaID LONG, Nota BYTE, DataNotes DATETIME)'

    $Conn.Close()
    Release-ComObjekt $Conn
    $Conn = $null
    Start-Sleep -Milliseconds 1000      # pritje e shkurtër që ndryshimet të shkruhen në disk
} catch { Ndalo 3 'Krijimi i tabelave me ADODB (DDL)' $_ }

# ---------- BLLOKU 4: Vetitë e fushave dhe të tabelave (DAO) ----------
try {
    $Db = $DbE.OpenDatabase($DbPath, $false, $false)
    $Db.TableDefs.Refresh()

    # Formati i datave (Kosova: dd.mm.yyyy)
    Set-FieldProp 'Nxenesit' 'DataLindjes' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Dosjet' 'DataRegjistrimit' 'Format' $dbText 'dd.mm.yyyy'
    Set-FieldProp 'Notat' 'DataNotes' 'Format' $dbText 'dd.mm.yyyy'

    # Nëntabelat (+) në Datasheet View: prindi tregon rreshtat e lidhur të fëmijës.
    # Nëse një veti nuk vendoset, skripti vazhdon me paralajmërim (nuk prek të dhënat).
    $Nentabelat = @(
        @('Klasat', 'Table.Nxenesit', 'KlasaID'),
        @('Nxenesit', 'Table.Notat', 'NxenesiID'),
        @('Lendet', 'Table.Notat', 'LendaID'),
        @('Mesuesit', 'Table.Lendet', 'MesuesiID')
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
# N07: lidhjet e diagramës ER. Enforce Referential Integrity, pa Cascade
# (Cascade, testi me vlerë që nuk ekziston dhe Relationship Report janë te P04).
try {
    New-AccRelation 'FK_Nxenesit_Klasat' 'Klasat' 'Nxenesit' 'KlasaID' 'KlasaID' 0
    New-AccRelation 'FK_Dosjet_Nxenesit' 'Nxenesit' 'Dosjet' 'NxenesiID' 'NxenesiID' $dbRelationUnique
    New-AccRelation 'FK_Lendet_Mesuesit' 'Mesuesit' 'Lendet' 'MesuesiID' 'MesuesiID' 0
    New-AccRelation 'FK_Notat_Nxenesit' 'Nxenesit' 'Notat' 'NxenesiID' 'NxenesiID' 0
    New-AccRelation 'FK_Notat_Lendet' 'Lendet' 'Notat' 'LendaID' 'LendaID' 0
} catch { Ndalo 5 'Krijimi i relacioneve (DAO)' $_ }

# ---------- BLLOKU 6: Të dhënat shembull (prindërit para fëmijëve) ----------
# Të dhëna fiktive. Datat #yyyy-mm-dd#, numrat dhjetorë me pikë, apostrofi shkruhet ''.
try {
    # Modelet (N06): 6 modele, nga skedarët te NoSQL
    Add-Rreshta 'Modelet' 'ModeliID, EmriModelit, Periudha, VitiFillimit, Struktura, Shembulli, Perparesia, Kufizimi' @(
        "1, 'Skedarë', '1950–1960', 1950, 'Skedarë të veçantë pa lidhje', 'Kartela, skedarë tekst', 'e thjeshtë', 'përsëritje, pa lidhje'",
        "2, 'Hierarkik', '1960–1970', 1960, 'Pemë: një prind, shumë fëmijë', 'IBM IMS', 'shpejtësi në pemë', 'një prind për çdo fëmijë'",
        "3, 'Rrjetë', '1969–1980', 1969, 'Rrjet: një fëmijë, disa prindër', 'CODASYL (IDMS)', 'lidhje shumë-me-shumë', 'strukturë e ndërlikuar'",
        "4, 'Relacional', '1970–sot', 1970, 'Tabela të lidhura me çelësa', 'Access, MySQL', 'tabela, SQL, integritet', 'varet nga skema'",
        "5, 'Objektor', '1990', 1990, 'Objekte me veti dhe metoda', 'db4o', 'objekte pa shndërrim', 'përdorim i kufizuar'",
        "6, 'NoSQL', '2000–sot', 2000, 'Dokumente, çelës–vlerë, grafe', 'MongoDB', 'shkallëzim i madh', 'konsistencë e lehtë'"
    )

    # Klasat (prind): 8 klasa, një për çdo kompjuter
    Add-Rreshta 'Klasat' 'KlasaID, EmriKlases, Drejtimi, KlasaNr, Salla' @(
        "1, '10-1', 'Informatikë', 10, 'A01'",
        "2, '10-2', 'Elektroteknikë', 10, 'A02'",
        "3, '11-1', 'Informatikë', 11, 'B11'",
        "4, '11-2', 'Mekatronikë', 11, 'B12'",
        "5, '11-3', 'Elektroteknikë', 11, 'B13'",
        "6, '12-1', 'Informatikë', 12, 'C21'",
        "7, '12-2', 'Mekatronikë', 12, 'C22'",
        "8, '12-3', 'Elektroteknikë', 12, 'C23'"
    )

    # Nxenesit (fëmijë i Klasat, prind i Dosjet dhe Notat): 20 nxënës fiktivë
    Add-Rreshta 'Nxenesit' 'NxenesiID, Emri, Mbiemri, Gjinia, DataLindjes, Qyteti, KlasaID' @(
        "1, 'Albin', 'Berisha', 'M', #2010-03-14#, 'Prizren', 1",
        "2, 'Besarta', 'Krasniqi', 'F', #2010-07-22#, 'Gjakovë', 1",
        "3, 'Dardan', 'Hoxha', 'M', #2010-11-05#, 'Rahovec', 1",
        "4, 'Erza', 'Gashi', 'F', #2010-01-30#, 'Suharekë', 2",
        "5, 'Fisnik', 'Morina', 'M', #2010-09-18#, 'Prizren', 2",
        "6, 'Gresa', 'Shala', 'F', #2009-04-11#, 'Malishevë', 3",
        "7, 'Ilir', 'Kelmendi', 'M', #2009-06-27#, 'Rahovec', 3",
        "8, 'Jeta', 'Rexhepi', 'F', #2009-12-03#, 'Gjakovë', 3",
        "9, 'Kushtrim', 'Bytyqi', 'M', #2009-02-16#, 'Prizren', 4",
        "10, 'Lirije', 'Hasani', 'F', #2009-08-09#, 'Suharekë', 4",
        "11, 'Mentor', 'Mustafa', 'M', #2009-05-21#, 'Rahovec', 5",
        "12, 'Njomza', 'Kryeziu', 'F', #2009-10-14#, 'Prizren', 5",
        "13, 'Orges', 'Rama', 'M', #2009-03-02#, 'Malishevë', 5",
        "14, 'Pranvera', 'Sylejmani', 'F', #2008-07-19#, 'Gjakovë', 6",
        "15, 'Qëndrim', 'Ahmeti', 'M', #2008-01-25#, 'Rahovec', 6",
        "16, 'Rinesa', 'Bajrami', 'F', #2008-09-30#, 'Prizren', 7",
        "17, 'Shpëtim', 'Lleshi', 'M', #2008-04-08#, 'Suharekë', 7",
        "18, 'Teuta', 'Zeqiri', 'F', #2008-12-12#, 'Malishevë', 7",
        "19, 'Uran', 'Avdiu', 'M', #2008-02-05#, 'Gjakovë', 8",
        "20, 'Valbona', 'Dema', 'F', #2008-06-17#, 'Rahovec', 8"
    )

    # Dosjet (1:1): 18 dosje; nxënësit 10 dhe 18 nuk kanë ende dosje
    Add-Rreshta 'Dosjet' 'NxenesiID, NrAmzes, DataRegjistrimit, ShkollaParaprake, TelefoniPrindit' @(
        "1, 'R2026-001', #2026-08-26#, 'SHMU Nr. 2, Prizren', '049 301 107'",
        "2, 'R2026-002', #2026-08-27#, 'SHMU Nr. 3, Gjakovë', '049 302 114'",
        "3, 'R2026-003', #2026-08-28#, 'SHMU Nr. 1, Rahovec', '049 303 121'",
        "4, 'R2026-004', #2026-08-29#, 'SHMU Nr. 2, Suharekë', '049 304 128'",
        "5, 'R2026-005', #2026-08-25#, 'SHMU Nr. 3, Prizren', '049 305 135'",
        "6, 'R2025-001', #2025-08-26#, 'SHMU Nr. 1, Malishevë', '049 306 142'",
        "7, 'R2025-002', #2025-08-27#, 'SHMU Nr. 2, Rahovec', '049 307 149'",
        "8, 'R2025-003', #2025-08-28#, 'SHMU Nr. 3, Gjakovë', '049 308 156'",
        "9, 'R2025-004', #2025-08-29#, 'SHMU Nr. 1, Prizren', '049 309 163'",
        "11, 'R2025-005', #2025-08-26#, 'SHMU Nr. 3, Rahovec', '049 311 177'",
        "12, 'R2025-006', #2025-08-27#, 'SHMU Nr. 1, Prizren', '049 312 184'",
        "13, 'R2025-007', #2025-08-28#, 'SHMU Nr. 2, Malishevë', '049 313 191'",
        "14, 'R2024-001', #2024-08-29#, 'SHMU Nr. 3, Gjakovë', '049 314 198'",
        "15, 'R2024-002', #2024-08-25#, 'SHMU Nr. 1, Rahovec', '049 315 205'",
        "16, 'R2024-003', #2024-08-26#, 'SHMU Nr. 2, Prizren', '049 316 212'",
        "17, 'R2024-004', #2024-08-27#, 'SHMU Nr. 3, Suharekë', '049 317 219'",
        "19, 'R2024-005', #2024-08-29#, 'SHMU Nr. 2, Gjakovë', '049 319 233'",
        "20, 'R2024-006', #2024-08-25#, 'SHMU Nr. 3, Rahovec', '049 320 240'"
    )

    # Mesuesit (prind i Lendet): 6 mësimdhënës fiktivë
    Add-Rreshta 'Mesuesit' 'MesuesiID, Emri, Mbiemri, Specializimi, Email' @(
        "1, 'Arbnora', 'Krasniqi', 'Informatikë', 'arbnora.krasniqi@shkolla.example'",
        "2, 'Burim', 'Gashi', 'Elektroteknikë', 'burim.gashi@shkolla.example'",
        "3, 'Drita', 'Berisha', 'Matematikë', 'drita.berisha@shkolla.example'",
        "4, 'Fatmir', 'Hoxha', 'Gjuhë shqipe', 'fatmir.hoxha@shkolla.example'",
        "5, 'Hana', 'Morina', 'Gjuhë angleze', 'hana.morina@shkolla.example'",
        "6, 'Kujtim', 'Shala', 'Fizikë', 'kujtim.shala@shkolla.example'"
    )

    # Lendet (fëmijë i Mesuesit, prind i Notat): 8 lëndë
    Add-Rreshta 'Lendet' 'LendaID, EmriLendes, OretJave, MesuesiID' @(
        "1, 'Bazat e të dhënave', 2, 1",
        "2, 'Programimi', 3, 1",
        "3, 'Rrjetat kompjuterike', 2, 2",
        "4, 'Elektronika digjitale', 3, 2",
        "5, 'Matematikë', 3, 3",
        "6, 'Gjuhë shqipe', 3, 4",
        "7, 'Gjuhë angleze', 2, 5",
        "8, 'Fizikë', 2, 6"
    )

    # Notat (tabela ndërmjetëse M:N, futet e fundit): 36 nota
    Add-Rreshta 'Notat' 'NotaID, NxenesiID, LendaID, Nota, DataNotes' @(
        "1, 1, 1, 5, #2026-09-25#",
        "2, 1, 5, 4, #2026-10-06#",
        "3, 1, 6, 5, #2026-10-19#",
        "4, 2, 1, 4, #2026-10-28#",
        "5, 2, 7, 3, #2026-11-09#",
        "6, 3, 2, 3, #2026-11-19#",
        "7, 4, 3, 4, #2026-09-23#",
        "8, 4, 5, 2, #2026-10-05#",
        "9, 4, 8, 3, #2026-10-15#",
        "10, 5, 4, 5, #2026-10-26#",
        "11, 6, 1, 3, #2026-11-06#",
        "12, 6, 6, 4, #2026-11-17#",
        "13, 7, 2, 5, #2026-09-21#",
        "14, 7, 8, 4, #2026-10-02#",
        "15, 8, 7, 5, #2026-10-13#",
        "16, 9, 2, 4, #2026-10-26#",
        "17, 9, 3, 3, #2026-11-04#",
        "18, 9, 4, 2, #2026-11-16#",
        "19, 10, 5, 5, #2026-09-21#",
        "20, 11, 3, 5, #2026-09-30#",
        "21, 11, 7, 4, #2026-10-12#",
        "22, 12, 4, 3, #2026-10-22#",
        "23, 12, 6, 2, #2026-11-02#",
        "24, 13, 8, 5, #2026-11-13#",
        "25, 14, 1, 5, #2026-09-17#",
        "26, 14, 2, 4, #2026-09-28#",
        "27, 14, 5, 3, #2026-10-09#",
        "28, 15, 6, 4, #2026-10-20#",
        "29, 16, 4, 4, #2026-11-02#",
        "30, 16, 8, 2, #2026-11-11#",
        "31, 17, 7, 5, #2026-09-15#",
        "32, 18, 3, 2, #2026-09-28#",
        "33, 19, 5, 4, #2026-10-07#",
        "34, 19, 6, 3, #2026-10-19#",
        "35, 19, 1, 2, #2026-10-29#",
        "36, 20, 2, 5, #2026-11-09#"
    )
} catch { Ndalo 6 'Futja e të dhënave' $_ }

# ---------- BLLOKU 7: Pyetësorët e ruajtur (DAO CreateQueryDef) ----------
try {
    # 1:N Klasat-Nxenesit: lista e nxënësve sipas klasës
    New-AccQuery 'Q_Nxenesit_Klasat' "SELECT Klasat.EmriKlases, Nxenesit.NxenesiID, Nxenesit.Emri, Nxenesit.Mbiemri, Nxenesit.Qyteti FROM Klasat INNER JOIN Nxenesit ON Klasat.KlasaID = Nxenesit.KlasaID ORDER BY Klasat.EmriKlases, Nxenesit.Mbiemri;"
    # 1:1 Nxenesit-Dosjet: çdo nxënës me dosjen e tij (dosja mund të mungojë)
    New-AccQuery 'Q_Nxenesit_Dosjet' "SELECT Nxenesit.NxenesiID, Nxenesit.Emri, Nxenesit.Mbiemri, Dosjet.NrAmzes, Dosjet.DataRegjistrimit FROM Nxenesit LEFT JOIN Dosjet ON Nxenesit.NxenesiID = Dosjet.NxenesiID ORDER BY Nxenesit.NxenesiID;"
    # M:N përmes Notat: 5 tabelat e lidhura japin pamjen e plotë të çdo note
    New-AccQuery 'Q_Notat_Plote' "SELECT Notat.NotaID, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Lendet.EmriLendes, Mesuesit.Mbiemri AS MbiemriMesuesit, Notat.Nota, Notat.DataNotes FROM (((Notat INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID) INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Mesuesit ON Lendet.MesuesiID = Mesuesit.MesuesiID ORDER BY Notat.NotaID;"

    # Kontrolli i nëntabelës së lëndës B për çdo kompjuter (Niveli 2)
    New-AccQuery 'Q_Kontroll_K1' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Bazat e të dhënave' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K2' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Programimi' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K3' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Rrjetat kompjuterike' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K4' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Elektronika digjitale' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K5' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Matematikë' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K6' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Gjuhë shqipe' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K7' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Gjuhë angleze' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
    New-AccQuery 'Q_Kontroll_K8' "SELECT Lendet.EmriLendes, Nxenesit.Emri, Nxenesit.Mbiemri, Klasat.EmriKlases, Notat.Nota, Notat.DataNotes FROM ((Notat INNER JOIN Lendet ON Notat.LendaID = Lendet.LendaID) INNER JOIN Nxenesit ON Notat.NxenesiID = Nxenesit.NxenesiID) INNER JOIN Klasat ON Nxenesit.KlasaID = Klasat.KlasaID WHERE Lendet.EmriLendes = 'Fizikë' ORDER BY Nxenesit.Mbiemri, Nxenesit.Emri;"
} catch { Ndalo 7 'Krijimi i pyetësorëve' $_ }

# ---------- BLLOKU 8: Verifikimi ----------
try {
    Write-Host ''
    Write-Host '=== VERIFIKIMI ==='
    $Db.TableDefs.Refresh()
    $Db.QueryDefs.Refresh()

    # Vlerat e pritura (të njëjtat si në çelësin e përgjigjeve)
    $PritetTabela = [ordered]@{ 'Modelet' = 6; 'Klasat' = 8; 'Nxenesit' = 20; 'Dosjet' = 18; 'Mesuesit' = 6; 'Lendet' = 8; 'Notat' = 36 }
    $PritetRel = [ordered]@{ 'FK_Nxenesit_Klasat' = $false; 'FK_Dosjet_Nxenesit' = $true; 'FK_Lendet_Mesuesit' = $false; 'FK_Notat_Nxenesit' = $false; 'FK_Notat_Lendet' = $false }      # $true = një-me-një
    $PritetQ = [ordered]@{ 'Q_Nxenesit_Klasat' = 20; 'Q_Nxenesit_Dosjet' = 20; 'Q_Notat_Plote' = 36; 'Q_Kontroll_K1' = 5; 'Q_Kontroll_K2' = 5; 'Q_Kontroll_K3' = 4; 'Q_Kontroll_K4' = 4; 'Q_Kontroll_K5' = 5; 'Q_Kontroll_K6' = 5; 'Q_Kontroll_K7' = 4; 'Q_Kontroll_K8' = 4 }
    $ParamShembull = @{}      # P02 nuk ka pyetësorë me parametra
    $PritetPaDosje = 2
    $Variantet = @(
        [pscustomobject]@{ K = 1; KlasaID = 1; Klasa = '10-1'; A = 3; LendaID = 1; Lenda = 'Bazat e të dhënave'; B = 5; NxenesiID = 1; Nxenesi = 'Albin Berisha'; C = 3; Amza = 'R2026-001' },
        [pscustomobject]@{ K = 2; KlasaID = 2; Klasa = '10-2'; A = 2; LendaID = 2; Lenda = 'Programimi'; B = 5; NxenesiID = 4; Nxenesi = 'Erza Gashi'; C = 3; Amza = 'R2026-004' },
        [pscustomobject]@{ K = 3; KlasaID = 3; Klasa = '11-1'; A = 3; LendaID = 3; Lenda = 'Rrjetat kompjuterike'; B = 4; NxenesiID = 6; Nxenesi = 'Gresa Shala'; C = 2; Amza = 'R2025-001' },
        [pscustomobject]@{ K = 4; KlasaID = 4; Klasa = '11-2'; A = 2; LendaID = 4; Lenda = 'Elektronika digjitale'; B = 4; NxenesiID = 9; Nxenesi = 'Kushtrim Bytyqi'; C = 3; Amza = 'R2025-004' },
        [pscustomobject]@{ K = 5; KlasaID = 5; Klasa = '11-3'; A = 3; LendaID = 5; Lenda = 'Matematikë'; B = 5; NxenesiID = 11; Nxenesi = 'Mentor Mustafa'; C = 2; Amza = 'R2025-005' },
        [pscustomobject]@{ K = 6; KlasaID = 6; Klasa = '12-1'; A = 2; LendaID = 6; Lenda = 'Gjuhë shqipe'; B = 5; NxenesiID = 14; Nxenesi = 'Pranvera Sylejmani'; C = 3; Amza = 'R2024-001' },
        [pscustomobject]@{ K = 7; KlasaID = 7; Klasa = '12-2'; A = 3; LendaID = 7; Lenda = 'Gjuhë angleze'; B = 4; NxenesiID = 16; Nxenesi = 'Rinesa Bajrami'; C = 2; Amza = 'R2024-003' },
        [pscustomobject]@{ K = 8; KlasaID = 8; Klasa = '12-3'; A = 2; LendaID = 8; Lenda = 'Fizikë'; B = 4; NxenesiID = 19; Nxenesi = 'Uran Avdiu'; C = 3; Amza = 'R2024-005' }
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

    # 8b. Fushat e tabelave të Shkollës (N07: atribut -> fushë, me tipin e saj)
    foreach ($t in @('Klasat', 'Nxenesit', 'Dosjet', 'Mesuesit', 'Lendet', 'Notat')) {
        $tdfX = $Db.TableDefs.Item($t)
        $pjeset = for ($i = 0; $i -lt $tdfX.Fields.Count; $i++) {
            $f = $tdfX.Fields.Item($i)
            $f.Name + ' ' + (Get-LlojiFushes $f)
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
        $a  = Get-Scalar ('SELECT Count(*) FROM Nxenesit WHERE KlasaID = ' + $v.KlasaID)
        $b  = Get-Scalar ('SELECT Count(*) FROM Notat WHERE LendaID = ' + $v.LendaID)
        $c  = Get-Scalar ('SELECT Count(*) FROM Notat WHERE NxenesiID = ' + $v.NxenesiID)
        $dd = [string](Get-Scalar ('SELECT NrAmzes FROM Dosjet WHERE NxenesiID = ' + $v.NxenesiID))
        $ok = ($a -eq $v.A) -and ($b -eq $v.B) -and ($c -eq $v.C) -and ($dd -eq $v.Amza)
        if (-not $ok) { $script:Mosperputhje++ }
        [pscustomobject]@{
            K          = $v.K
            A_Klasa    = $v.Klasa + ': ' + $a
            B_Lenda    = $v.Lenda + ': ' + $b
            C_Nxenesi  = $v.Nxenesi + ': ' + $c
            C_Dosja    = $dd
            Kontrolli  = $(if ($ok) { 'OK' } else { 'GABIM' })
        }
    }
    Write-Host 'Variantet (A = nxënës në klasë, B = nota në lëndë, C = nota të nxënësit dhe dosja e tij):'
    $raportV | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # 8f. Kontrolle të veçanta
    $nPaDosje = Get-Scalar 'SELECT Count(*) FROM Nxenesit LEFT JOIN Dosjet ON Nxenesit.NxenesiID = Dosjet.NxenesiID WHERE Dosjet.NxenesiID Is Null'
    Write-Host ('1:1: nxënës pa dosje = ' + $nPaDosje + ' (pritej ' + $PritetPaDosje + '); asnjë nxënës nuk ka më shumë se një dosje (çelësi primar).')
    if ($nPaDosje -ne $PritetPaDosje) { $script:Mosperputhje++ }

    $nPlote = Get-Scalar 'SELECT Count(*) FROM Q_Notat_Plote'
    if ($nPlote -eq $PritetTabela['Notat']) { Write-Host ('N07: Q_Notat_Plote bashkon 5 tabela dhe jep ' + $nPlote + ' rreshta, një për çdo notë.') -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM N07: Q_Notat_Plote jep ' + $nPlote + ' rreshta.') -ForegroundColor Red }

    $nVit = Get-Scalar 'SELECT TOP 1 EmriModelit FROM Modelet ORDER BY VitiFillimit DESC, ModeliID DESC'
    Write-Host ('N06: modeli më i ri sipas VitiFillimit = ' + $nVit + ' (pritej NoSQL).')
    if ([string]$nVit -ne 'NoSQL') { $script:Mosperputhje++ }

    $lenda6 = [string](Get-Scalar 'SELECT EmriLendes FROM Lendet WHERE LendaID = 6')
    if ($lenda6 -ceq 'Gjuhë shqipe') { Write-Host 'Kodimi: shkronjat ë/ç u ruajtën saktë në bazë (LendaID 6 = Gjuhë shqipe).' -ForegroundColor Green }
    else { $script:Mosperputhje++; Write-Host ('GABIM kodimi: LendaID 6 ka EmriLendes = ' + $lenda6) -ForegroundColor Red }

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
    Write-Host 'Përfundoi. Hap C:\BazatDB\P02\P02_Zgjidhja.accdb në Access.' -ForegroundColor Green
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
