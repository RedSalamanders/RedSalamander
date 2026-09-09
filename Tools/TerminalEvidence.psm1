Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ('RedSalamander.Tools.TerminalEvidenceJcs' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;

namespace RedSalamander.Tools
{
    public static class TerminalEvidenceJcs
    {
        private const long MaxSafeInteger = 9007199254740991L;

        public static byte[] Serialize(object value)
        {
            StringBuilder output = new StringBuilder();
            HashSet<object> active = new HashSet<object>(ReferenceComparer.Instance);
            AppendValue(output, value, active, 0);
            return new UTF8Encoding(false, true).GetBytes(output.ToString());
        }

        public static string Sha256Hex(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }

            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] digest = sha256.ComputeHash(bytes);
                StringBuilder hex = new StringBuilder(64);
                foreach (byte value in digest)
                {
                    hex.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }
                return hex.ToString();
            }
        }

        public static object[] SortIdentityRecords(object[] records, string[] keyProperties)
        {
            if (records == null)
            {
                throw new ArgumentNullException("records");
            }
            if (keyProperties == null || keyProperties.Length == 0)
            {
                throw new ArgumentException("At least one key property is required.", "keyProperties");
            }

            HashSet<string> keyNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (string keyProperty in keyProperties)
            {
                if (String.IsNullOrEmpty(keyProperty))
                {
                    throw new ArgumentException("Identity key property names must be nonempty strings.", "keyProperties");
                }
                if (!keyNames.Add(keyProperty))
                {
                    throw new ArgumentException("Identity key property names must be unique.", "keyProperties");
                }
            }

            IdentityRecord[] sortable = new IdentityRecord[records.Length];
            for (int index = 0; index < records.Length; ++index)
            {
                object record = Unwrap(records[index]);
                if (record == null)
                {
                    throw new ArgumentException("Identity records cannot be null.", "records");
                }

                string[] keys = new string[keyProperties.Length];
                for (int keyIndex = 0; keyIndex < keyProperties.Length; ++keyIndex)
                {
                    object keyValue;
                    if (!TryGetProperty(record, keyProperties[keyIndex], out keyValue))
                    {
                        throw new ArgumentException(
                            "Identity record " + index.ToString(CultureInfo.InvariantCulture) +
                            " is missing sort key '" + keyProperties[keyIndex] + "'.",
                            "records");
                    }

                    keyValue = Unwrap(keyValue);
                    if (keyValue == null)
                    {
                        throw new ArgumentException(
                            "Identity record " + index.ToString(CultureInfo.InvariantCulture) +
                            " has a null sort key '" + keyProperties[keyIndex] + "'.",
                            "records");
                    }
                    string text = keyValue as string;
                    if (text == null)
                    {
                        throw new ArgumentException(
                            "Identity sort key '" + keyProperties[keyIndex] + "' must be a string.",
                            "records");
                    }
                    ValidateUnicode(text, "Identity sort key");
                    keys[keyIndex] = text;
                }
                sortable[index] = new IdentityRecord(record, keys);
            }

            Array.Sort(sortable, IdentityRecordComparer.Instance);
            for (int index = 1; index < sortable.Length; ++index)
            {
                if (IdentityRecordComparer.Instance.Compare(sortable[index - 1], sortable[index]) == 0)
                {
                    throw new ArgumentException("Identity records contain a duplicate sort-key tuple.", "records");
                }
            }

            object[] sorted = new object[sortable.Length];
            for (int index = 0; index < sortable.Length; ++index)
            {
                sorted[index] = sortable[index].Record;
            }
            return sorted;
        }

        public static string GetMachineHash()
        {
            string seed = null;
            try
            {
                using (Microsoft.Win32.RegistryKey localMachine =
                    Microsoft.Win32.RegistryKey.OpenBaseKey(
                        Microsoft.Win32.RegistryHive.LocalMachine,
                        Microsoft.Win32.RegistryView.Default))
                using (Microsoft.Win32.RegistryKey cryptography =
                    localMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography", false))
                {
                    if (cryptography != null)
                    {
                        seed = cryptography.GetValue(
                            "MachineGuid",
                            null,
                            Microsoft.Win32.RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                    }
                }
            }
            catch (System.Security.SecurityException)
            {
                seed = null;
            }
            catch (UnauthorizedAccessException)
            {
                seed = null;
            }
            catch (System.IO.IOException)
            {
                seed = null;
            }

            if (String.IsNullOrEmpty(seed))
            {
                seed = Environment.MachineName;
            }
            if (String.IsNullOrEmpty(seed))
            {
                throw new InvalidOperationException("Neither MachineGuid nor the computer name is readable and nonempty.");
            }

            byte[] seedBytes = new UTF8Encoding(false, true).GetBytes(seed);
            string fullHash = Sha256Hex(seedBytes);
            return fullHash.Substring(0, 12);
        }

        private static void AppendValue(
            StringBuilder output,
            object originalValue,
            HashSet<object> active,
            int depth)
        {
            if (depth > 128)
            {
                throw new ArgumentException("JCS input exceeds the maximum nesting depth of 128.");
            }

            object value = Unwrap(originalValue);
            if (value == null)
            {
                output.Append("null");
                return;
            }

            if (value is bool)
            {
                output.Append((bool)value ? "true" : "false");
                return;
            }

            string stringValue = value as string;
            if (stringValue != null)
            {
                AppendQuotedString(output, stringValue);
                return;
            }
            if (value is char)
            {
                AppendQuotedString(output, value.ToString());
                return;
            }

            if (IsInteger(value))
            {
                AppendInteger(output, value);
                return;
            }
            if (value is double)
            {
                output.Append(FormatDouble((double)value));
                return;
            }
            if (value is float)
            {
                output.Append(FormatDouble((double)(float)value));
                return;
            }
            if (value is decimal)
            {
                throw new ArgumentException("System.Decimal is not an RFC 8785 IEEE-754 number. Use a string or Double.");
            }

            IDictionary dictionary = value as IDictionary;
            if (dictionary != null)
            {
                EnterContainer(value, active);
                try
                {
                    AppendDictionary(output, dictionary, active, depth);
                }
                finally
                {
                    active.Remove(value);
                }
                return;
            }

            if (!(value is PSCustomObject) && value is IEnumerable)
            {
                EnterContainer(value, active);
                try
                {
                    AppendArray(output, (IEnumerable)value, active, depth);
                }
                finally
                {
                    active.Remove(value);
                }
                return;
            }

            if (value is PSCustomObject)
            {
                EnterContainer(value, active);
                try
                {
                    AppendPSObject(output, PSObject.AsPSObject(value), active, depth);
                }
                finally
                {
                    active.Remove(value);
                }
                return;
            }

            throw new ArgumentException(
                "Unsupported JCS input type '" + value.GetType().FullName + "'.");
        }

        private static object Unwrap(object value)
        {
            PSObject psObject = value as PSObject;
            if (psObject != null && psObject.BaseObject != null &&
                !Object.ReferenceEquals(psObject, psObject.BaseObject))
            {
                return psObject.BaseObject;
            }
            return value;
        }

        private static bool IsInteger(object value)
        {
            return value is sbyte || value is byte ||
                   value is short || value is ushort ||
                   value is int || value is uint ||
                   value is long || value is ulong;
        }

        private static void AppendInteger(StringBuilder output, object value)
        {
            if (value is ulong)
            {
                ulong unsignedValue = (ulong)value;
                if (unsignedValue > (ulong)MaxSafeInteger)
                {
                    throw new ArgumentOutOfRangeException(
                        "value",
                        "JSON integer values must fit exactly in IEEE-754 binary64; represent uint64 evidence values as decimal strings.");
                }
                output.Append(unsignedValue.ToString(CultureInfo.InvariantCulture));
                return;
            }

            long signedValue = Convert.ToInt64(value, CultureInfo.InvariantCulture);
            if (signedValue < -MaxSafeInteger || signedValue > MaxSafeInteger)
            {
                throw new ArgumentOutOfRangeException(
                    "value",
                    "JSON integer values must fit exactly in IEEE-754 binary64; represent int64 evidence values as decimal strings.");
            }
            output.Append(signedValue.ToString(CultureInfo.InvariantCulture));
        }

        private static string FormatDouble(double value)
        {
            if (Double.IsNaN(value) || Double.IsInfinity(value))
            {
                throw new ArgumentException("RFC 8785 forbids non-finite JSON numbers.");
            }
            if (String.Equals(
                    typeof(object).Assembly.GetName().Name,
                    "mscorlib",
                    StringComparison.Ordinal))
            {
                throw new PlatformNotSupportedException(
                    "RFC 8785 floating-point serialization requires the modern .NET shortest-round-trip formatter. " +
                    "Evidence manifests remain supported on this runtime when all JSON numbers are exact safe integers.");
            }
            if (value == 0.0)
            {
                return "0";
            }

            bool negative = value < 0.0;
            string roundTrip = Math.Abs(value).ToString("R", CultureInfo.InvariantCulture);
            int exponentMarker = roundTrip.IndexOfAny(new char[] { 'E', 'e' });
            string significand = exponentMarker >= 0
                ? roundTrip.Substring(0, exponentMarker)
                : roundTrip;
            int exponent = exponentMarker >= 0
                ? Int32.Parse(roundTrip.Substring(exponentMarker + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture)
                : 0;

            int decimalPoint = significand.IndexOf('.');
            int decimalPosition = decimalPoint >= 0 ? decimalPoint : significand.Length;
            string rawDigits = decimalPoint >= 0
                ? significand.Remove(decimalPoint, 1)
                : significand;

            int leadingZeros = 0;
            while (leadingZeros < rawDigits.Length && rawDigits[leadingZeros] == '0')
            {
                ++leadingZeros;
            }
            string digits = rawDigits.Substring(leadingZeros);
            int k = decimalPosition + exponent - leadingZeros;

            int end = digits.Length;
            while (end > 1 && digits[end - 1] == '0')
            {
                --end;
            }
            digits = digits.Substring(0, end);

            StringBuilder formatted = new StringBuilder();
            if (negative)
            {
                formatted.Append('-');
            }

            int digitCount = digits.Length;
            if (k > 0 && k <= 21)
            {
                if (k >= digitCount)
                {
                    formatted.Append(digits);
                    formatted.Append('0', k - digitCount);
                }
                else
                {
                    formatted.Append(digits, 0, k);
                    formatted.Append('.');
                    formatted.Append(digits, k, digitCount - k);
                }
            }
            else if (k <= 0 && k > -6)
            {
                formatted.Append("0.");
                formatted.Append('0', -k);
                formatted.Append(digits);
            }
            else
            {
                formatted.Append(digits[0]);
                if (digitCount > 1)
                {
                    formatted.Append('.');
                    formatted.Append(digits, 1, digitCount - 1);
                }
                formatted.Append('e');
                int scientificExponent = k - 1;
                if (scientificExponent >= 0)
                {
                    formatted.Append('+');
                }
                formatted.Append(scientificExponent.ToString(CultureInfo.InvariantCulture));
            }
            return formatted.ToString();
        }

        private static void AppendQuotedString(StringBuilder output, string value)
        {
            ValidateUnicode(value, "JSON string");
            output.Append('"');
            for (int index = 0; index < value.Length; ++index)
            {
                char character = value[index];
                switch (character)
                {
                    case '"':
                        output.Append("\\\"");
                        break;
                    case '\\':
                        output.Append("\\\\");
                        break;
                    case '\b':
                        output.Append("\\b");
                        break;
                    case '\t':
                        output.Append("\\t");
                        break;
                    case '\n':
                        output.Append("\\n");
                        break;
                    case '\f':
                        output.Append("\\f");
                        break;
                    case '\r':
                        output.Append("\\r");
                        break;
                    default:
                        if (character <= '\u001f')
                        {
                            output.Append("\\u");
                            output.Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            output.Append(character);
                        }
                        break;
                }
            }
            output.Append('"');
        }

        private static void ValidateUnicode(string value, string description)
        {
            for (int index = 0; index < value.Length; ++index)
            {
                char character = value[index];
                if (Char.IsHighSurrogate(character))
                {
                    if (index + 1 >= value.Length || !Char.IsLowSurrogate(value[index + 1]))
                    {
                        throw new ArgumentException(description + " contains an unpaired UTF-16 high surrogate.");
                    }
                    ++index;
                }
                else if (Char.IsLowSurrogate(character))
                {
                    throw new ArgumentException(description + " contains an unpaired UTF-16 low surrogate.");
                }
            }
        }

        private static void AppendDictionary(
            StringBuilder output,
            IDictionary dictionary,
            HashSet<object> active,
            int depth)
        {
            List<PropertyValue> properties = new List<PropertyValue>();
            HashSet<string> propertyNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (DictionaryEntry entry in dictionary)
            {
                string name = entry.Key as string;
                if (name == null)
                {
                    throw new ArgumentException("JCS object property names must be strings.");
                }
                ValidateUnicode(name, "JSON property name");
                if (!propertyNames.Add(name))
                {
                    throw new ArgumentException("JCS input contains duplicate object property '" + name + "'.");
                }
                properties.Add(new PropertyValue(name, entry.Value));
            }
            AppendProperties(output, properties, active, depth);
        }

        private static void AppendPSObject(
            StringBuilder output,
            PSObject psObject,
            HashSet<object> active,
            int depth)
        {
            List<PropertyValue> properties = new List<PropertyValue>();
            HashSet<string> propertyNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (PSPropertyInfo property in psObject.Properties)
            {
                if (!property.IsGettable)
                {
                    throw new ArgumentException("JCS object property '" + property.Name + "' is not readable.");
                }
                ValidateUnicode(property.Name, "JSON property name");
                if (!propertyNames.Add(property.Name))
                {
                    throw new ArgumentException("JCS input contains duplicate object property '" + property.Name + "'.");
                }
                properties.Add(new PropertyValue(property.Name, property.Value));
            }
            AppendProperties(output, properties, active, depth);
        }

        private static void AppendProperties(
            StringBuilder output,
            List<PropertyValue> properties,
            HashSet<object> active,
            int depth)
        {
            properties.Sort(PropertyValueComparer.Instance);
            output.Append('{');
            for (int index = 0; index < properties.Count; ++index)
            {
                if (index != 0)
                {
                    output.Append(',');
                }
                AppendQuotedString(output, properties[index].Name);
                output.Append(':');
                AppendValue(output, properties[index].Value, active, depth + 1);
            }
            output.Append('}');
        }

        private static void AppendArray(
            StringBuilder output,
            IEnumerable values,
            HashSet<object> active,
            int depth)
        {
            output.Append('[');
            bool first = true;
            foreach (object value in values)
            {
                if (!first)
                {
                    output.Append(',');
                }
                first = false;
                AppendValue(output, value, active, depth + 1);
            }
            output.Append(']');
        }

        private static void EnterContainer(object value, HashSet<object> active)
        {
            if (!active.Add(value))
            {
                throw new ArgumentException("JCS input contains a reference cycle.");
            }
        }

        private static bool TryGetProperty(object record, string propertyName, out object value)
        {
            IDictionary dictionary = record as IDictionary;
            if (dictionary != null)
            {
                foreach (DictionaryEntry entry in dictionary)
                {
                    string key = entry.Key as string;
                    if (key != null && String.Equals(key, propertyName, StringComparison.Ordinal))
                    {
                        value = entry.Value;
                        return true;
                    }
                }
                value = null;
                return false;
            }

            foreach (PSPropertyInfo property in PSObject.AsPSObject(record).Properties)
            {
                if (String.Equals(property.Name, propertyName, StringComparison.Ordinal))
                {
                    if (!property.IsGettable)
                    {
                        value = null;
                        return false;
                    }
                    value = property.Value;
                    return true;
                }
            }
            value = null;
            return false;
        }

        private sealed class PropertyValue
        {
            public PropertyValue(string name, object value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; private set; }
            public object Value { get; private set; }
        }

        private sealed class PropertyValueComparer : IComparer<PropertyValue>
        {
            public static readonly PropertyValueComparer Instance = new PropertyValueComparer();

            public int Compare(PropertyValue left, PropertyValue right)
            {
                return StringComparer.Ordinal.Compare(left.Name, right.Name);
            }
        }

        private sealed class IdentityRecord
        {
            public IdentityRecord(object record, string[] keys)
            {
                Record = record;
                Keys = keys;
            }

            public object Record { get; private set; }
            public string[] Keys { get; private set; }
        }

        private sealed class IdentityRecordComparer : IComparer<IdentityRecord>
        {
            public static readonly IdentityRecordComparer Instance = new IdentityRecordComparer();

            public int Compare(IdentityRecord left, IdentityRecord right)
            {
                for (int index = 0; index < left.Keys.Length; ++index)
                {
                    int comparison = StringComparer.Ordinal.Compare(left.Keys[index], right.Keys[index]);
                    if (comparison != 0)
                    {
                        return comparison;
                    }
                }
                return 0;
            }
        }

        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceComparer Instance = new ReferenceComparer();

            public new bool Equals(object left, object right)
            {
                return Object.ReferenceEquals(left, right);
            }

            public int GetHashCode(object value)
            {
                return RuntimeHelpers.GetHashCode(value);
            }
        }
    }
}
'@
}

function ConvertTo-RSJcsUtf8Bytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    return ,([RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($InputObject))
}

function Get-RSJcsSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject
    )

    $bytes = [RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($InputObject)
    return [RedSalamander.Tools.TerminalEvidenceJcs]::Sha256Hex($bytes)
}

function Get-RSJcsIdentitySetSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Records,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$KeyProperty
    )

    $sorted = [RedSalamander.Tools.TerminalEvidenceJcs]::SortIdentityRecords(
        [object[]]$Records,
        [string[]]$KeyProperty)
    $bytes = [RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($sorted)
    return [RedSalamander.Tools.TerminalEvidenceJcs]::Sha256Hex($bytes)
}

function Test-RSEvidenceLeaf {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Leaf,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    if ([string]::IsNullOrWhiteSpace($Leaf) -or
        -not [string]::Equals([IO.Path]::GetFileName($Leaf), $Leaf, [StringComparison]::Ordinal) -or
        $Leaf -eq '.' -or $Leaf -eq '..') {
        throw "$ParameterName must be one nonempty file-name leaf without directory traversal."
    }
}

function Get-RSRequiredManifestString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $value = $null
    $found = $false
    if ($Manifest -is [Collections.IDictionary]) {
        foreach ($entry in $Manifest.GetEnumerator()) {
            if ($entry.Key -is [string] -and
                [string]::Equals([string]$entry.Key, $Name, [StringComparison]::Ordinal)) {
                $value = $entry.Value
                $found = $true
                break
            }
        }
    } elseif ($null -ne $Manifest) {
        foreach ($property in $Manifest.PSObject.Properties) {
            if ([string]::Equals($property.Name, $Name, [StringComparison]::Ordinal)) {
                $value = $property.Value
                $found = $true
                break
            }
        }
    }

    if (-not $found -or $value -isnot [string] -or [string]::IsNullOrEmpty([string]$value)) {
        throw "Manifest must contain a nonempty string '$Name' property."
    }
    return [string]$value
}

function Write-RSCreateNewBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $stream = [IO.FileStream]::new(
        $Path,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None,
        4096,
        [IO.FileOptions]::WriteThrough)
    try {
        $stream.Write($Bytes, 0, $Bytes.Length)
        $stream.Flush($true)
    } finally {
        $stream.Dispose()
    }
}

function Write-RSEvidenceRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Manifest,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$JsonLeaf,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CompanionLeaf
    )

    Test-RSEvidenceLeaf -Leaf $JsonLeaf -ParameterName 'JsonLeaf'
    Test-RSEvidenceLeaf -Leaf $CompanionLeaf -ParameterName 'CompanionLeaf'
    if ([string]::Equals($JsonLeaf, $CompanionLeaf, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'JsonLeaf and CompanionLeaf must identify distinct files.'
    }

    $resolvedDirectory = [IO.Path]::GetFullPath($Directory).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar)
    if (-not [IO.Directory]::Exists($resolvedDirectory)) {
        throw "Evidence directory does not exist: $resolvedDirectory"
    }

    $runId = Get-RSRequiredManifestString -Manifest $Manifest -Name 'runId'
    $machineHash = Get-RSRequiredManifestString -Manifest $Manifest -Name 'machineHash'
    if ($runId -cnotmatch '^[0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{6}$') {
        throw "Manifest runId '$runId' is not UTC yyyy-MM-dd_HHmmss."
    }
    if ($machineHash -cnotmatch '^[0-9a-f]{12}$') {
        throw "Manifest machineHash '$machineHash' is not 12 lowercase hexadecimal characters."
    }
    $runDirectoryName = [IO.Path]::GetFileName($resolvedDirectory)
    $areaDirectory = [IO.Path]::GetDirectoryName($resolvedDirectory)
    $machineDirectory = if ([string]::IsNullOrEmpty($areaDirectory)) {
        $null
    } else {
        [IO.Path]::GetDirectoryName($areaDirectory)
    }
    $machineDirectoryName = if ([string]::IsNullOrEmpty($machineDirectory)) {
        $null
    } else {
        [IO.Path]::GetFileName($machineDirectory)
    }
    if (-not [string]::Equals($runId, $runDirectoryName, [StringComparison]::Ordinal)) {
        throw "Manifest runId '$runId' does not match evidence directory '$runDirectoryName'."
    }
    if (-not [string]::Equals($machineHash, $machineDirectoryName, [StringComparison]::Ordinal)) {
        throw "Manifest machineHash '$machineHash' does not match archive profile directory '$machineDirectoryName'."
    }
    $actualMachineHash = Get-RSTestMachineHash
    if (-not [string]::Equals($machineHash, $actualMachineHash, [StringComparison]::Ordinal)) {
        throw "Manifest machineHash '$machineHash' does not match this machine '$actualMachineHash'."
    }

    $manifestPath = [IO.Path]::GetFullPath((Join-Path $resolvedDirectory $JsonLeaf))
    $companionPath = [IO.Path]::GetFullPath((Join-Path $resolvedDirectory $CompanionLeaf))
    if (-not [string]::Equals(
            [IO.Path]::GetDirectoryName($manifestPath),
            $resolvedDirectory,
            [StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals(
            [IO.Path]::GetDirectoryName($companionPath),
            $resolvedDirectory,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Evidence leaves must resolve directly inside Directory.'
    }

    $manifestBytes = [RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($Manifest)
    $digest = [RedSalamander.Tools.TerminalEvidenceJcs]::Sha256Hex($manifestBytes)
    $companionBytes = [Text.Encoding]::ASCII.GetBytes($digest + "`n")

    $operationId = [guid]::NewGuid().ToString('N')
    $lockPath = Join-Path $resolvedDirectory '.rs-evidence-write.lock'
    $manifestTemporaryPath = Join-Path $resolvedDirectory (".rs-evidence-$operationId.json.tmp")
    $companionTemporaryPath = Join-Path $resolvedDirectory (".rs-evidence-$operationId.sha256.tmp")
    $lockStream = $null
    $publishedManifest = $false
    $publishedCompanion = $false

    try {
        $lockStream = [IO.FileStream]::new(
            $lockPath,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write,
            [IO.FileShare]::None,
            1,
            [IO.FileOptions]::DeleteOnClose)

        if ([IO.File]::Exists($manifestPath) -or [IO.File]::Exists($companionPath)) {
            throw 'Evidence record is create-new; a target file already exists.'
        }

        Write-RSCreateNewBytes -Path $manifestTemporaryPath -Bytes $manifestBytes
        Write-RSCreateNewBytes -Path $companionTemporaryPath -Bytes $companionBytes

        [IO.File]::Move($manifestTemporaryPath, $manifestPath)
        $publishedManifest = $true
        [IO.File]::Move($companionTemporaryPath, $companionPath)
        $publishedCompanion = $true
    } catch {
        if ($publishedCompanion -and [IO.File]::Exists($companionPath)) {
            [IO.File]::Delete($companionPath)
        }
        if ($publishedManifest -and [IO.File]::Exists($manifestPath)) {
            [IO.File]::Delete($manifestPath)
        }
        throw
    } finally {
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
        }
        if ([IO.File]::Exists($manifestTemporaryPath)) {
            [IO.File]::Delete($manifestTemporaryPath)
        }
        if ([IO.File]::Exists($companionTemporaryPath)) {
            [IO.File]::Delete($companionTemporaryPath)
        }
    }

    return $manifestPath
}

function Get-RSTestMachineHash {
    [CmdletBinding()]
    param()

    return [RedSalamander.Tools.TerminalEvidenceJcs]::GetMachineHash()
}

Export-ModuleMember -Function `
    ConvertTo-RSJcsUtf8Bytes, `
    Get-RSJcsSha256, `
    Get-RSJcsIdentitySetSha256, `
    Write-RSEvidenceRecord, `
    Get-RSTestMachineHash
