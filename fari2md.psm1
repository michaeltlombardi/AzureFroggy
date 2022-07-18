function ConvertTo-TitleCase {
  [cmdletbinding()]
  [OutputType('String[]')]
  param(
    [Parameter(ValueFromPipeline)]
    [string[]]$Text
  )

  begin {
    $TextInfo = (Get-Culture).TextInfo
  }

  process {
    foreach ($Entry in $Text) {
      $TextInfo.ToTitleCase($Entry.ToLower())
    }
  }
}

function Resolve-HtmlEntity {
  [cmdletbinding()]
  [OutputType('String[]')]
  param(
    [Parameter(ValueFromPipeline)]
    [string[]]$Text
  )

  begin {
    $Pattern = '(?n)(?<EntityCode>&([a-zA-Z]+|#[0-9]+);)'
  }

  process {
    foreach ($Entry in $Text) {
      if ($Entry -match $Pattern) {
        $PlainText = [System.Web.HttpUtility]::HtmlDecode($Matches.EntityCode.ToLower())
        $Entry -replace [regex]::Escape($Matches.EntityCode), $PlainText
      } else {
        $Entry
      }
    }
  }
}

class Header {
  [int]$Level
  [string]$Text

  Header([string]$Text) {
    $this.Text = $Text
    $this.Level = 1
  }
  Header([string]$Text, [int]$Level) {
    $this.Text = $Text
    $this.Level = $Level
  }

  [string] ToMarkdown() {
    return ('#' * $this.Level) + ' ' + $this.Text
  }
}

class Die {
  [int]$Size = 6

  Die() {
    $this.Size = 6
  }
  Die([int]$Size) {
    $this.Size = $Size
  }

  Die([string]$DiceString) {
    $this.Size = $DiceString -split 'd' | Select-Object -Last 1
  }

  [string] ToString() {
    return "d$($this.Size)"
  }
}

class Dice {
  [int]$Count = 1
  [int]$Size = 6

  Dice() {
    $this.Count = 1
    $this.Size = 6
  }

  Dice([int]$Count, [int]$Size) {
    $this.Count = $Count
    $this.Size = $Size
  }

  Dice([string]$DiceString) {
    $this.Count, $this.Size = $DiceString -split 'd'
  }

  [string] ToString() {
    return "$($this.Count)d$($this.Size)"
  }
}

class HitDice : Dice {
  [int]$Marks = 0

  HitDice() {
    $this.Count = 1
    $this.Size = 6
    $this.Marks = 0
  }

  HitDice([int]$Count, [int]$Size) {
    $this.Count = $Count
    $this.Size = $Size
    $this.Marks = 0
  }

  HitDice([int]$Count, [int]$Size, [int]$Marks) {
    $this.Count = $Count
    $this.Size = $Size
    $this.Marks = $Marks
  }

  HitDice([string]$DiceString, [int]$Marks) {
    $this.Count, $this.Size = $DiceString -split 'd'
    $this.Marks = $Marks
  }

  [string] ToMarkdownSection() {
    return @(
      '## Hit Dice'
      @(
        "- Size: $($this.Count)d$($this.Size)"
        "- Marks: $($this.Marks)"
      ) -join "`n"
    ) -join "`n`n"
  }
}

class Debt {
  [string]$Creditor
  [string]$Description
  [int]$Silver
}

class Possessions {
  [string]$Stored
  [string]$Carried
  [int]$Silver
  [Debt[]]$Debts

  [string] ToMarkdownSection() {
    $Markdown = @(
      '## Possessions'
      @(
        "- Silver: $($this.Silver)"
        "- Carried: $($this.Carried -replace "`n", "`n  ")"
        "- Stored: $($this.Stored -replace "`n", "`n  ")"
      ) -join "`n"
      "### Debts`n`n"
    ) -join "`n`n"

    if ($this.Debts.Count) {
      $Markdown += $this.Debts | ForEach-Object {
        "- $($_.Silver) to $($_.Creditor): $($_.Description -replace "`n", "`n  ")"
      } | Join-String -Separator "`n"
    } else {
      $Markdown += 'None.'
    }
    return $Markdown
  }
}

class Trick {
  [string]$Name
  [Dice]$Dice
  [string]$Description
  [int]$Uses
  [int]$Marks

  [string] ToMarkdownSection() {
    return @(
      "> ##### $($this.Name)"
      $this.Description.Trim() -split "`n" | ForEach-Object { "> $_" } | Join-String -Separator "`n"
      @(
        "> - Dice: $($this.Dice.ToString())"
        "> - Uses: $($this.Uses)"
        "> - Marks: $($this.Marks)"
      ) -join "`n"
    ) -join "`n>`n"
  }
}

class Domain {
  [string]$Name
  [Die]$Die
  [int]$Marks
  [int]$TrickSlots
  [Trick[]]$Tricks

  [string] ToMarkdownSection() {
    $Markdown = @(
      "### $($this.Name)"
      @(
        "- Die Size: $($this.Die)"
        "- Trick Slots: $($this.Tricks.Count) / $($this.TrickSlots)"
        "- Marks: $($this.Marks)"
      ) -join "`n"
      "#### Tricks`n`n"
    ) -join "`n`n"

    if ($this.Tricks.Count) {
      $Markdown += $this.Tricks.ToMarkdownSection() -join "`n`n"
    } else {
      $Markdown += 'None.'
    }

    return $Markdown
  }
}

class Picaroon {
  [string]$Name
  [string]$Pronouns
  [string]$Group
  [int]$HitPoints
  [HitDice]$HitDice = [HitDice]::new()
  [Domain[]]$Domains
  [Possessions]$Possessions = [Possessions]::new()
}

function Import-FariPicaroon {
  [CmdletBinding()]
  [OutputType('Picaroon')]
  param(
    [string]$JsonPath
  )

  begin {
    $Character = [Picaroon]::new()
  }

  process {
    $Raw = Get-Content -Path $JsonPath -Raw
    $Json = $Raw | ConvertFrom-Json -Depth 100

    if ($Json.Name -match '^(?<Name>.+) \((?<Pronouns>.+)\)') {
      $Character.Name = $Matches.Name
      $Character.Pronouns = $Matches.Pronouns
    } else {
      $Character.Name = $Json.Name
      $Character.Pronouns = 'Pronouns'
    }

    $Character.Group = $Json.Group | Resolve-HtmlEntity | ConvertTo-TitleCase

    foreach ($Page in $Json.Pages) {
      $Section = $Page.Label | Resolve-HtmlEntity | ConvertTo-TitleCase
      switch ($Section) {
        'Character' {
          $Character = Import-CharacterSection -Character $Character -Page $Page
        }
        'Domains' {
          $Character = Import-DomainSection -Character $Character -Page $Page
        }
        'Tricks' {
          $Character = Import-TrickSection -Character $Character -Page $Page
        }
        default {}
      }
    }

    $Character
  }
}

function Import-CharacterSection {
  [CmdletBinding()]
  [OutputType('Picaroon')]
  param (
    [Picaroon]$Character = [Picaroon]::new(),
    [pscustomobject]$Page
  )

  process {
    $Sections = $Page.Rows.Columns.Sections
    foreach ($Section in $Sections) {
      switch ($Section.Label) {
        'Hit Dice' {
          foreach ($Block in $Section.Blocks) {
            switch ($Block.Label) {
              'Count' {
                $Character.HitDice.Count, $Character.HitDice.Size = $Block.Meta.Commands -split 'd'
              }
              'Hit Points' {
                $Character.HitPoints = $Block.Value
              }
              'Marks' {
                $Character.HitDice.Marks = $Block.Value
              }
            }
          }
        }
        'Possessions' {
          foreach ($Block in $Section.Blocks) {
            switch ($Block.Label) {
              'Silver on Hand' {
                $Character.Possessions.Silver = $Block.Value
              }
              'Things Carried' {
                $Character.Possessions.Carried = $Block.Value
              }
              'Things Stored' {
                $Character.Possessions.Stored = $Block.Value
              }
            }
          }

          $DebtBlocks = $Section.Blocks | Where-Object {
            $_.Label -match '^Silver(ed)? owed'
          }
          foreach ($Block in $DebtBlocks) {

            if ($Block.label -match '^Silver(ed)? owed to \(?(?<Creditor>(?!Creditor)[^\(\)]+)\)?') {
              $Character.Possessions.Debts += [Debt]@{
                Creditor    = $Matches.Creditor
                Silver      = $Block.Value
                Description = $Block.Meta.HelperText
              }
            }
          }
        }
      }
    }
    $Character
  }
}

Function Import-DomainSection {
  [CmdletBinding()]
  [OutputType('Picaroon')]
  param (
    [Picaroon]$Character = [Picaroon]::new(),
    [pscustomobject]$Page
  )

  process {
    $Character.Domains = @()
    $Sections = $Page.Rows.Columns.Sections
    foreach ($Section in $Sections) {
      $Domain = [Domain]@{
        Name       = $Section.Label
        Die        = [Die]::new()
        Marks      = $Section.Blocks.Where({ $_.Label -eq 'Marks' }).Value
        TrickSlots = $Section.Blocks.Where({
            $_.Label -eq 'Available Tricks'
          }).Value.Count
      }

      $Section.Blocks
      | Where-Object { $_.Label -like 'Domain*' }
      | ForEach-Object {
        $Domain.Die.Size = $_.Meta.Commands -split 'd' | Select-Object -Last 1
      }
      $Character.Domains += $Domain
    }

    $Character
  }
}

Function Import-TrickSection {
  [CmdletBinding()]
  [OutputType('Picaroon')]
  param (
    [Picaroon]$Character = [Picaroon]::new(),
    [pscustomobject]$Page
  )

  process {
    $Sections = $Page.Rows.Columns.Sections
    $TrickList = @()

    foreach ($Section in $Sections) {
      if ($Section.Label -match '(?<TrickName>.+) \((?<DomainName>[^\)]+)') {
        $TrickName = $Matches.TrickName
        $DomainName = $Matches.DomainName
      } else {
        Write-Error "Could not determine name and domain for entry $($Section.Label)"
        continue
      }

      if ($Character.Domains.Where({ $_.Name -eq $DomainName })) {
        $Trick = [Trick]@{
          Name        = $TrickName
          Dice        = @{
            Count = 1
            Size  = 6
          }
          Description = $Section.Blocks.Where({ $_.Label -match 'Effect' }).Value
          Uses        = $Section.Blocks.Where({ $_.Label -eq 'Uses' }).Value.Count
          Marks       = $Section.Blocks.Where({ $_.Label -eq 'Marks' }).Value
        }

        $Section.Blocks
        | Where-Object { $_.Label -like 'Trick Dice' }
        | ForEach-Object {
          $Trick.Dice.Count, $Trick.Dice.Size = $_.Meta.Commands -split 'd'
        }

        $TrickList += @{
          Domain = $DomainName
          Trick  = $Trick
        }
      } else {
        $Message = @(
          "Character does not seem to have domain '$DomainName';"
          "can't add trick '$TrickName'"
        ) -join ' '
        Write-Error $Message
      }
    }

    $Character.Domains = $Character.Domains | ForEach-Object -Process {
      $WorkingDomain = $_
      $WorkingDomain.Tricks = $TrickList | Group-Object -Property Domain
      | Where-Object { $_.Name -eq $WorkingDomain.Name }
      | Select-Object -ExpandProperty Group
      | Select-Object -ExpandProperty Trick
      $WorkingDomain
    }

    $Character
  }
}

function ConvertTo-Slug {
  [CmdletBinding()]
  [OutputType('String')]
  param(
    [Parameter(ValueFromPipeline)]
    [string]$Text
  )

  process {
    $Text.ToLowerInvariant() -replace '[^a-zA-Z0-9\s-]', '' -replace '\s+', ' ' -replace '\s', '-'
  }
}

function Export-Picaroon {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline)]
    [Picaroon]$Character = [Picaroon]::new(),
    [string]$FolderPath
  )

  process {
    $OutputPath = Join-Path -Path $FolderPath -ChildPath "$($Character.Name | ConvertTo-Slug).md"
    $Builder = New-Object -TypeName System.Text.StringBuilder

    $null = $Builder.AppendLine("# $($Character.Name)").AppendLine()
    $null = $Builder.AppendLine("**Pronouns:** $($Character.Pronouns)").AppendLine()
    $null = $Builder.AppendLine($Character.HitDice.ToMarkdownSection())
    $null = $Builder.AppendLine("- Current Hit Points: $($Character.HitPoints)").AppendLine()
    $null = $Builder.AppendLine($Character.Possessions.ToMarkdownSection()).AppendLine()
    $null = $Builder.AppendLine('## Domains').AppendLine()
    $null = $Builder.AppendLine($Character.Domains.ToMarkdownSection() -join "`n`n").AppendLine()
    $Builder.ToString().Trim() -replace "`r`n", "`n" | Set-Content -Path $OutputPath -NoNewline
    Add-Content -Path $OutputPath -Value "`n" -NoNewline
  }
}

function Update-CharacterMarkdown {
  [cmdletbinding()]
  param(
    [string]$FolderPath = './characters'
  )

  process {
    foreach ($BlobFile in (Get-ChildItem -Path $FolderPath -Recurse -Include '*.fari.json')) {
      Write-Verbose "Processing $($BlobFile.Name)"
      Import-FariPicaroon -JsonPath $BlobFile.FullName | Export-Picaroon -FolderPath $FolderPath
    }
  }
}
