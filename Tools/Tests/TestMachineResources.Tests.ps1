$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\Modules\Testing\TestMachineResources.psm1'
$schemaPath = Join-Path $repoRoot 'Specs\Testing\TestMachineResources.schema.json'
Import-Module $modulePath -Force

Describe 'Per-machine test resources' {
    It 'treats an absent manifest as legacy environment compatibility' {
        $configuration = Read-RSTestMachineResources -TestRoot $TestDrive -SchemaPath $schemaPath

        $configuration.Present | Should Be $false
        @($configuration.Resources).Count | Should Be 0
        $configuration.CapabilityDigest | Should Match '^[0-9a-f]{64}$'
    }

    It 'validates and projects connection, SMB, MTP, and alternate local-root resources without secrets' {
        $configRoot = Join-Path $TestDrive 'config'
        New-Item -ItemType Directory -Path $configRoot -Force | Out-Null
        $manifest = @{
            schema = 'red-salamander.test-machine-resources.v1'
            resources = @(
                @{ id = 'connection.s3.primary'; kind = 'connection-profile'; enabled = $true; profile = 'SINON S3 SelfTest' },
                @{ id = 'connection.onedrive.personal'; kind = 'connection-profile'; enabled = $true; profile = 'SINON OneDrive SelfTest' },
                @{ id = 'path.smb.primary'; kind = 'smb-share'; enabled = $true; root = '\\server\share\RedSalamander-SelfTest' },
                @{ id = 'device.mtp.primary'; kind = 'mtp-device'; enabled = $true; device = 'Phone'; scratch = '/Internal storage/RedSalamander-SelfTest' },
                @{ id = 'path.local.alternate'; kind = 'local-test-root'; enabled = $true; root = 'D:\RedSalamander.Perf' }
            )
        }
        $manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $configRoot 'machine-resources.json') -Encoding UTF8

        $configuration = Read-RSTestMachineResources -TestRoot $TestDrive -SchemaPath $schemaPath
        $configuration.Present | Should Be $true
        $configuration.Environment['REDSALAMANDER_SELFTEST_CONN_S3'] | Should Be 'SINON S3 SelfTest'
        $configuration.Environment['REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL'] | Should Be 'SINON OneDrive SelfTest'
        $configuration.Environment['REDSALAMANDER_SELFTEST_SMB_ROOT'] | Should Be '\\server\share\RedSalamander-SelfTest'
        $configuration.Environment['REDSALAMANDER_SELFTEST_MTP_DEVICE'] | Should Be 'Phone'
        $configuration.Environment['REDSALAMANDER_TEST_ALTERNATE_ROOT'] | Should Be 'D:\RedSalamander.Perf'
        $configuration.CapabilityDigest | Should Match '^[0-9a-f]{64}$'

        $oldS3 = $env:REDSALAMANDER_SELFTEST_CONN_S3
        $oldFtp = $env:REDSALAMANDER_SELFTEST_CONN_FTP
        $oldAlternateRoot = $env:REDSALAMANDER_TEST_ALTERNATE_ROOT
        try {
            $env:REDSALAMANDER_SELFTEST_CONN_FTP = 'ambient-value-must-be-cleared'
            $lease = Set-RSTestMachineResourceEnvironment -Configuration $configuration
            $env:REDSALAMANDER_SELFTEST_CONN_S3 | Should Be 'SINON S3 SelfTest'
            $env:REDSALAMANDER_SELFTEST_CONN_FTP | Should BeNullOrEmpty
            Restore-RSTestMachineResourceEnvironment -Lease $lease
            $env:REDSALAMANDER_SELFTEST_CONN_FTP | Should Be 'ambient-value-must-be-cleared'
        } finally {
            [Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_CONN_S3', $oldS3, 'Process')
            [Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_CONN_FTP', $oldFtp, 'Process')
            [Environment]::SetEnvironmentVariable('REDSALAMANDER_TEST_ALTERNATE_ROOT', $oldAlternateRoot, 'Process')
        }
    }

    It 'rejects an unsafe SMB share root' {
        $configRoot = Join-Path $TestDrive 'config'
        New-Item -ItemType Directory -Path $configRoot -Force | Out-Null
        @{
            schema = 'red-salamander.test-machine-resources.v1'
            resources = @(
                @{ id = 'path.smb.primary'; kind = 'smb-share'; enabled = $true; root = '\\server\share\Production' }
            )
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $configRoot 'machine-resources.json') -Encoding UTF8

        { Read-RSTestMachineResources -TestRoot $TestDrive -SchemaPath $schemaPath } |
            Should Throw "SMB root must include a path segment containing 'selftest'"
    }

    It 'rejects an arbitrary alternate local directory' {
        $configRoot = Join-Path $TestDrive 'config'
        New-Item -ItemType Directory -Path $configRoot -Force | Out-Null
        @{
            schema = 'red-salamander.test-machine-resources.v1'
            resources = @(
                @{ id = 'path.local.alternate'; kind = 'local-test-root'; enabled = $true; root = 'D:\Scratch' }
            )
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $configRoot 'machine-resources.json') -Encoding UTF8

        { Read-RSTestMachineResources -TestRoot $TestDrive -SchemaPath $schemaPath } |
            Should Throw 'Test machine resource manifest does not match the schema'
    }
}
