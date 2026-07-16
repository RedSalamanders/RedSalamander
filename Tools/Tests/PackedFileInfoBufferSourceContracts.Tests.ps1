Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'Packed FileInfo buffer source contracts' {
    BeforeAll {
        $ownerSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\PackedFileInfoBuffer.h') -Raw
        $facadeHeaders = @(
            'Plugins\FileSystem7z\FileSystem7z.h'
            'Plugins\FileSystemCurl\FileSystemCurl.h'
            'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.h'
            'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.h'
            'Plugins\FileSystemMtp\FileSystemMtp.h'
            'Plugins\FileSystemS3\FileSystemS3.h'
        ) | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }
        $facadeSources = @(
            'Plugins\FileSystem7z\FileSystem7z.cpp'
            'Plugins\FileSystemCurl\FileSystemCurl.Shared.cpp'
            'Plugins\FileSystemGoogleDrive\FilesInformationGoogleDrive.cpp'
            'Plugins\FileSystemMicrosoftDrive\FilesInformationMicrosoftDrive.cpp'
            'Plugins\FileSystemMtp\FilesInformationMtp.cpp'
            'Plugins\FileSystemS3\FilesInformationS3.cpp'
        ) | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }
        $localHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystem\FileSystem.h') -Raw
        $dummySource = Get-Content -LiteralPath (Join-Path $repoRoot 'Plugins\FileSystemDummy\FileSystemDummy.cpp') -Raw
    }

    It 'owns checked sizing aligned construction and bounded traversal in one helper' {
        $ownerSource | Should Match 'class\s+PackedFileInfoBuffer\s+final'
        $ownerSource | Should Match 'kEntryAlignment\s*=\s*alignof\(FileInfo\)'
        $ownerSource | Should Match 'TryComputeEntrySize'
        $ownerSource | Should Match 'ERROR_ARITHMETIC_OVERFLOW'
        $ownerSource | Should Match 'ERROR_INVALID_DATA'
        $ownerSource | Should Match 'advance < entrySize'
        $ownerSource | Should Match 'advance % kEntryAlignment'
    }

    It 'places the shared owner beneath exactly the six equivalent buffered COM facades' {
        foreach ($header in $facadeHeaders) {
            $header | Should Match 'Common::Plugins::PackedFileInfoBuffer\s+_packedBuffer'
            $header | Should Not Match 'ComputeEntrySizeBytes'
            $header | Should Not Match 'std::vector<std::byte>\s+_buffer'
        }
        foreach ($source in $facadeSources) {
            $source | Should Match '_packedBuffer\.GetBuffer\(ppFileInfo\)'
            $source | Should Match '_packedBuffer\.GetBufferSize\(pSize\)'
            $source | Should Match '_packedBuffer\.GetAllocatedSize\(pSize\)'
            $source | Should Match '_packedBuffer\.GetCount\(pCount\)'
            $source | Should Match '_packedBuffer\.Get\(index, ppEntry\)'
            $source | Should Match '_packedBuffer\.Build\(entries,'
            $source | Should Not Match 'ComputeEntrySizeBytes'
        }
    }

    It 'keeps streaming local enumeration and prebuilt dummy fixtures outside the buffered owner' {
        $localHeader | Should Match 'class\s+FilesInformation\s+final'
        $localHeader | Should Not Match 'PackedFileInfoBuffer'
        $dummySource | Should Not Match 'PackedFileInfoBuffer'
    }
}
