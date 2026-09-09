Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$inputSchema = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineCollectionDriverInput.schema.json')
$resultSchema = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineCollectionDriverResult.schema.json')
$descriptorSchema = Join-Path $repoRoot (
    'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json')

function Test-RSCollectionSchema {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Schema
    )

    return Test-Json `
        -Json ($Value | ConvertTo-Json -Depth 100 -Compress) `
        -SchemaFile $Schema `
        -ErrorAction SilentlyContinue
}

function New-RSWrapperInput {
    return [ordered]@{
        formatId = 'red-salamander-terminal-engine-wrapper-input'
        version = 1
        candidate = [ordered]@{
            candidateId = 'candidate-alpha'
            pin = 'pin-alpha'
            pinSha256 = '1' * 64
            sourceArchiveSha256 = '2' * 64
            patchSha256 = '3' * 64
            resultTree = '4' * 40
            collectionDescriptorSha256 = '5' * 64
            buildKind = 'candidate-wrapper-v1'
        }
        lane = [ordered]@{
            platform = 'x64'
            configuration = 'Release'
        }
        paths = [ordered]@{
            sourceRoot = 'C:\source'
            buildRoot = 'C:\build'
            noticeOutputPath = 'C:\source\notices.txt'
            adapterHeaderPath = 'C:\host\adapter.h'
            layoutHeaderPath = 'C:\host\layout.h'
            hostHeaderIncludeRoot = 'C:\host\wil-include'
            msbuildPath = 'C:\tool\msbuild.exe'
            buildReceiptPath = 'C:\build\receipt.json'
            sourceDependenciesPath = 'C:\build\dependencies.json'
        }
        pathBindings = [object[]]@()
        artifacts = [ordered]@{
            adapterDllLeaf = 'adapter-alpha.dll'
            privateRuntimeLibraries = [string[]]@()
        }
        gate0 = [ordered]@{
            adapterHeaderSha256 = '6' * 64
            layoutHeaderSha256 = '7' * 64
            hostHeaderTreeSha256 = '8' * 64
            capabilities = '127'
        }
    }
}

function New-RSGenericDescriptor {
    return [ordered]@{
        formatId = 'red-salamander-terminal-engine-collection-descriptor'
        version = 1
        candidateId = 'candidate-alpha'
        headerPackages = [object[]]@()
        legal = [ordered]@{
            closureId = 'closure-alpha'
            outputRelativePath = 'generated/notices.txt'
            dependencies = [object[]]@(
                [ordered]@{
                    dependencyId = 'dependency-alpha'
                    spdx = 'MIT'
                    inputs = [object[]]@(
                        [ordered]@{
                            inputId = 'license-alpha'
                            kind = 'license'
                            destinationLeaf = 'license-alpha.txt'
                            sha256 = '8' * 64
                            sizeBytes = '1'
                            source = [ordered]@{
                                sourceKind = 'frozen-input-whole'
                                sourceInputId = 'input-alpha'
                            }
                        })
                })
        }
        build = [ordered]@{
            kind = 'provider-generic-cmake-msvc-v1'
            buildTargets = [string[]]@('adapter-alpha')
            pathBindings = [object[]]@()
            adapter = [ordered]@{
                dllLeaf = 'adapter-alpha.dll'
                projectLeaf = 'adapter-alpha.vcxproj'
            }
            staticLibraries = [object[]]@()
            privateRuntimeLibraries = [object[]]@()
        }
    }
}

Describe 'Terminal-engine candidate-wrapper ABI' {
    It 'accepts the closed wrapper input and pass result' {
        (Test-RSCollectionSchema `
                -Value (New-RSWrapperInput) `
                -Schema $inputSchema) | Should Be $true
        $result = [ordered]@{
            formatId = 'red-salamander-terminal-engine-wrapper-result'
            version = 1
            status = 'pass'
            failure = $null
        }
        (Test-RSCollectionSchema -Value $result -Schema $resultSchema) |
            Should Be $true
    }

    It 'rejects a passing result with failure data' {
        $result = [ordered]@{
            formatId = 'red-salamander-terminal-engine-wrapper-result'
            version = 1
            status = 'pass'
            failure = [ordered]@{
                phase = 'build'
                code = 'build-process-failed'
            }
        }
        (Test-RSCollectionSchema -Value $result -Schema $resultSchema) |
            Should Be $false
    }

    It 'requires the explicit frozen host-header path and tree identity' {
        $input = New-RSWrapperInput
        $input.paths.Remove('hostHeaderIncludeRoot')
        (Test-RSCollectionSchema `
                -Value $input `
                -Schema $inputSchema) | Should Be $false

        $input = New-RSWrapperInput
        $input.gate0.Remove('hostHeaderTreeSha256')
        (Test-RSCollectionSchema `
                -Value $input `
                -Schema $inputSchema) | Should Be $false
    }

    It 'forbids driver data on a provider-generic build kind' {
        $descriptor = New-RSGenericDescriptor
        $descriptor.build['driverRelativePath'] = 'tools/driver.ps1'
        (Test-RSCollectionSchema `
                -Value $descriptor `
                -Schema $descriptorSchema) | Should Be $false
    }

    It 'requires a driver only for the counted wrapper kind' {
        $descriptor = New-RSGenericDescriptor
        $descriptor.build = [ordered]@{
            kind = 'candidate-wrapper-v1'
            driverRelativePath = 'tools/driver.ps1'
            pathBindings = [object[]]@()
            adapter = [ordered]@{
                dllLeaf = 'adapter-alpha.dll'
            }
            privateRuntimeLibraries = [object[]]@()
        }
        (Test-RSCollectionSchema `
                -Value $descriptor `
                -Schema $descriptorSchema) | Should Be $true
        $descriptor.build.Remove('driverRelativePath')
        (Test-RSCollectionSchema `
                -Value $descriptor `
                -Schema $descriptorSchema) | Should Be $false
    }

    It 'restricts wrapper drivers to PowerShell scripts' {
        $descriptor = New-RSGenericDescriptor
        $descriptor.build = [ordered]@{
            kind = 'candidate-wrapper-v1'
            driverRelativePath = 'tools/driver.exe'
            pathBindings = [object[]]@()
            adapter = [ordered]@{
                dllLeaf = 'adapter-alpha.dll'
            }
            privateRuntimeLibraries = [object[]]@()
        }
        (Test-RSCollectionSchema `
                -Value $descriptor `
                -Schema $descriptorSchema) | Should Be $false
    }
}
