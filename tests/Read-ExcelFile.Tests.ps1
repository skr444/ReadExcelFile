$testData = @(
    @{
        Path = ""
        Sheets = @(
            @{
                Sheet = "Sheet1"
                Columns = @(
                    @{
                        Name = "Col one"
                        Type = "int"
                        Index = 0
                    },
                    @{
                        Name = "Col two"
                        Type = "string"
                        Index = 1
                    },
                    @{
                        Name = "date"
                        Type = "date"
                        Index = 2
                    },
                    @{
                        Name = "time"
                        Type = "time"
                        Index = 3
                    },
                    @{
                        Name = "date time"
                        Type = "datetime"
                        Index = 4
                    }
                )
                Rows = @(
                    @{
                        Index = 0
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "1"
                            },
                            @{
                                ColIndex = 1
                                Content = "hey"
                            },
                            @{
                                ColIndex = 2
                                Content = "2022-03-20"
                            },
                            @{
                                ColIndex = 3
                                Content = "15:47"
                            },
                            @{
                                ColIndex = 4
                                Content = "2022-11-30 23:02:00"
                            }
                        )
                    },
                    @{
                        Index = 1
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "2"
                            },
                            @{
                                ColIndex = 1
                                Content = "wow"
                            },
                            @{
                                ColIndex = 2
                                Content = "2022-03-22"
                            },
                            @{
                                ColIndex = 3
                                Content = "15:49"
                            },
                            @{
                                ColIndex = 4
                                Content = "2022-11-30 23:02:00"
                            }
                        )
                    },
                    @{
                        Index = 2
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "3"
                            },
                            @{
                                ColIndex = 1
                                Content = "hou"
                            },
                            @{
                                ColIndex = 2
                                Content = "2022-03-25"
                            },
                            @{
                                ColIndex = 3
                                Content = "15:52"
                            },
                            @{
                                ColIndex = 4
                                Content = "2022-11-30 23:02:00"
                            }
                        )
                    },
                    @{
                        Index = 3
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "4"
                            },
                            @{
                                ColIndex = 1
                                Content = "hiwoijf"
                            },
                            @{
                                ColIndex = 2
                                Content = "2022-03-29"
                            },
                            @{
                                ColIndex = 3
                                Content = "15:56"
                            },
                            @{
                                ColIndex = 4
                                Content = "2022-11-30 23:02:00"
                            }
                        )
                    },
                    @{
                        Index = 4
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "5"
                            },
                            @{
                                ColIndex = 1
                                Content = "sldkfjs"
                            },
                            @{
                                ColIndex = 2
                                Content = "2022-04-03"
                            },
                            @{
                                ColIndex = 3
                                Content = "16:01"
                            },
                            @{
                                ColIndex = 4
                                Content = "2022-11-30 23:02:00"
                            }
                        )
                    }
                )
            },
            @{
                Sheet = "Sheet2"
                Columns = @(
                    @{
                        Name = "Words"
                        Type = "string"
                        Index = 0
                    },
                    @{
                        Name = "Numbers"
                        Type = "float"
                        Index = 1
                    }
                )
                Rows = @(
                    @{
                        Index = 0
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "hello"
                            },
                            @{
                                ColIndex = 1
                                Content = "1"
                            }
                        )
                    },
                    @{
                        Index = 1
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "world"
                            },
                            @{
                                ColIndex = 1
                                Content = "234234.345345"
                            }
                        )
                    },
                    @{
                        Index = 2
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "what"
                            },
                            @{
                                ColIndex = 1
                                Content = "500"
                            }
                        )
                    },
                    @{
                        Index = 3
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "is going"
                            },
                            @{
                                ColIndex = 1
                                Content = "234"
                            }
                        )
                    },
                    @{
                        Index = 4
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "on?"
                            },
                            @{
                                ColIndex = 1
                                Content = "2"
                            }
                        )
                    }
                )
            },
            @{
                Sheet = "Sheet3"
                Columns = @(
                    @{
                        Name = "First name"
                        Type = "string"
                        Index = 0
                    },
                    @{
                        Name = "Last Name"
                        Type = "string"
                        Index = 1
                    }
                )
                Rows = @(
                    @{
                        Index = 0
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "Ben"
                            },
                            @{
                                ColIndex = 1
                                Content = "Foster"
                            }
                        )
                    },
                    @{
                        Index = 1
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "Helen"
                            },
                            @{
                                ColIndex = 1
                                Content = "Mirren"
                            }
                        )
                    },
                    @{
                        Index = 2
                        Cells = @(
                            @{
                                ColIndex = 0
                                Content = "Kurt"
                            },
                            @{
                                ColIndex = 1
                                Content = "Russel"
                            }
                        )
                    }
                )
            }
        )
    }
)

BeforeDiscovery {
    $VerbosePreference = "Continue"
    foreach ($item in $testData) {
        $item.Path = ($PSCommandPath | Split-Path -Parent).Replace('tests', 'samples') | Join-Path -ChildPath 'testdata\testdata.xlsx'
    }
}

BeforeAll {
    . $PSCommandPath.Replace('.Tests.ps1', '.ps1').Replace('tests', 'src')
    # $epplusPath = ($PSCommandPath | Split-Path -Parent).Replace('tests', 'src') | Join-Path -ChildPath 'EPPlus.dll'
    # if (Test-Path -Path $epplusPath) {
    #     [System.Reflection.Assembly]::LoadFile($epplusPath)
    #     Write-Verbose "Assembly loaded at '${epplusPath}'."
    # } else {
    #     throw "Requires '${epplusPath}'"
    # }
}

Describe "Read-ExcelFile" -ForEach $testData {
    
    Context "test data sheet" {
        It "is parsed correctly" {
            foreach ($sheet in $_.Sheets) {
                $result = Read-ExcelFile -File $_.Path -WorkSheetName $sheet.Sheet

                $result | Should -Not -BeNullOrEmpty

                $result.Count | Should -Be ($sheet.Rows.Count)

                $rowIndex = 0
                foreach ($row in $sheet.Rows) {
                    $colIndex = 0
                    foreach ($col in $sheet.Columns) {
                        $actual = $result[$rowIndex]."$($col.Name)"
                        $expected = ($row.Cells | Where-Object { $_.ColIndex -eq $colIndex }).Content
                        if ($col.Type -eq "float") {
                            $actual = [float]$actual
                            $expected = [float]$expected
                            $actual | Should -Be $expected
                        } elseif ($col.Type -eq "time") {
                            $parts = $expected -split ":"
                            $expected = Get-Date -Hour $parts[0] -Minute $parts[1] -Second 0 -Millisecond 0
                            $actual.Hours | Should -Be $expected.Hours
                            $actual.Minutes | Should -Be $expected.Minutes
                            $actual.Seconds | Should -Be $expected.Seconds
                        } elseif ($col.Type -eq "date") {
                            $parts = $expected -split "-"
                            $expected = Get-Date -Year $parts[0] -Month $parts[1] -Day $parts[2] `
                                -Hour 0 -Minute 0 -Second 0 -Millisecond 0
                            $actual.Years | Should -Be $expected.Years
                            $actual.Months | Should -Be $expected.Months
                            $actual.Days | Should -Be $expected.Days
                        } elseif ($col.Type -eq "datetime") {
                            $parts = $expected -split " "
                            $dateparts = $parts[0] -split "-"
                            $timeparts = $parts[1] -split ":"
                            $expected = Get-Date -Year $dateparts[0] -Month $dateparts[1] -Day $dateparts[2] `
                                -Hour $timeparts[0] -Minute $timeparts[1] -Second 0 -Millisecond 0
                            
                            $actual.Years | Should -Be $expected.Years
                            $actual.Months | Should -Be $expected.Months
                            $actual.Days | Should -Be $expected.Days
                            $actual.Hours | Should -Be $expected.Hours
                            $actual.Minutes | Should -Be $expected.Minutes
                            $actual.Seconds | Should -Be $expected.Seconds
                        } else {
                            $actual | Should -Be $expected
                        }
                        Write-Verbose "actual: ${actual} - expected: ${expected}"
                        Write-Verbose "actual: $($actual.GetType()) - expected: $($expected.GetType())"
                        $colIndex++
                    }
                    $rowIndex++
                }
            }
        }
    }
}
