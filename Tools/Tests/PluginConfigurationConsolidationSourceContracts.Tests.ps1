Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Describe 'Plugin configuration consolidation source contracts' {
    BeforeAll {
        $commonHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\PluginConfiguration.h') -Raw
        $commonSource = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\Common\PluginConfiguration.cpp') -Raw
        $manageSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\ManagePluginsDialog.cpp') -Raw
        $preferencesHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\Preferences.Internal.h') -Raw
        $preferencesSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\Preferences.Plugin.Configuration.cpp') -Raw
    }

    It 'keeps the schema model and codec in Common' {
        $commonHeader | Should Match 'namespace Common::PluginConfiguration'
        $commonHeader | Should Match 'SchemaParseResult ParseSchema\('
        $commonHeader | Should Match 'ConfigurationParseResult ParseConfiguration\('
        $commonHeader | Should Match 'HRESULT SerializeConfiguration\('
        $commonHeader | Should Match 'bool TryGetBoolToggleChoiceIndices\('
        $commonHeader | Should Match 'Members that are not represented by the schema are'
        $commonSource | Should Match 'root = CloneJsonValue\(root\);'
        $commonSource | Should Match 'if \(parseHr == E_OUTOFMEMORY\)'
    }

    It 'uses the shared model and codec from both editor entry points' {
        $manageSource | Should Match 'using PluginConfigField\s+= Common::PluginConfiguration::Field;'
        $preferencesHeader | Should Match 'using PrefsPluginConfigField\s+= Common::PluginConfiguration::Field;'

        foreach ($source in @($manageSource, $preferencesSource)) {
            $source | Should Match 'Common::PluginConfiguration::ParseSchema\('
            $source | Should Match 'Common::PluginConfiguration::ParseConfiguration\('
            $source | Should Match 'Common::PluginConfiguration::SerializeConfiguration\('
            $source | Should Match 'Common::PluginConfiguration::TryGetBoolToggleChoiceIndices\('
            $source | Should Not Match 'yyjson_(read|mut_doc_new|obj_get|mut_obj_add)'
        }
    }

    It 'preserves the original configuration object across Preferences commits' {
        $preferencesHeader | Should Match 'std::string pluginsDetailsConfigSourceJsonUtf8;'
        $preferencesSource | Should Match 'BuildConfigurationJson\(state\.pluginsDetailsConfigFields, state\.pluginsDetailsConfigSourceJsonUtf8\)'
        $preferencesSource | Should Match 'state\.pluginsDetailsConfigSourceJsonUtf8 = configJson;'
    }
}
