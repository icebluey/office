#!/usr/bin/env bash
export PATH=$PATH:/bin:/sbin:/usr/bin:/usr/sbin:/usr/local/bin:/usr/local/sbin
TZ='UTC'; export TZ
umask 022

_install_jq() {
    set -euo pipefail
    local _tmp_dir="$(mktemp -d)"
    cd "${_tmp_dir}"
    _jq_url="$(wget -qO- 'https://jqlang.org/download/' | grep 'jq-linux-amd64' | sed 's/"/\n/g' | grep 'http.*jq-linux-amd64' | sort -V | tail -n 1)"
    wget -q -c -t 9 -T 9 "${_jq_url}" -O jq-linux-amd64
    mv jq-linux-amd64 jq
    chmod 0755 jq
    file jq | sed -n -E 's/^(.*):[[:space:]]*ELF.*, not stripped.*/\1/p' | xargs --no-run-if-empty -I '{}' strip '{}'
    rm -f /usr/bin/jq
    rm -f /usr/local/bin/jq
    install -v -c -m 0755 jq /usr/bin/jq
    cp -f /usr/bin/jq /usr/local/bin/jq
    cd /tmp
    rm -fr "${_tmp_dir}"
}
_install_7z() {
    set -euo pipefail
    local _tmp_dir="$(mktemp -d)"
    cd "${_tmp_dir}"
    _7zip_loc="$(wget -qO- 'https://www.7-zip.org/download.html' | grep -i '\-linux-x64.tar' | grep -i 'href="' | sed 's|"|\n|g' | grep -i '\-linux-x64.tar' | sort -V | tail -n 1)"
    wget -q -c -t 9 -T 9 "https://www.7-zip.org/${_7zip_loc}"
    tar -xof *.tar*
    sleep 1
    rm -f *.tar*
    file 7zzs | sed -n -E 's/^(.*):[[:space:]]*ELF.*, not stripped.*/\1/p' | xargs --no-run-if-empty -I '{}' strip '{}'
    rm -f /usr/bin/7z
    rm -f /usr/local/bin/7z
    install -v -c -m 0755 7zzs /usr/bin/7z
    cp -f /usr/bin/7z /usr/local/bin/7z
    cd /tmp
    rm -fr "${_tmp_dir}"
}

# Microsoft Office Downloader Script
# Production-ready version with robust error handling and clean architecture
###############################################################################
# https://docs.microsoft.com/en-us/officeupdates/update-history-office-2019
# https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData
# https://www.coolhub.top/tech-articles/channels.html

# https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/7983bac0-e531-40cf-be00-fd24fe66619c
# https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/${FFN}

# https://officecdn.microsoft.com/pr/

# Current Channel
# Production::CC
# Channel=Current
# 492350f6-3a01-4f97-b9c0-c7c6ddf67d60
# https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60

# Monthly Enterprise Channel
# Production::MEC
# Channel=MonthlyEnterprise
# 55336b82-a18d-4dd6-b5f6-9e5095c314a6
# https://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6

# Semi-Annual Enterprise Channel
# Production::DC
# Channel=SemiAnnual
# 7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
# https://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114

# Office 2019 Perpetual Enterprise Channel
# Production::LTSC
# Channel=PerpetualVL2019
# f2e724c1-748f-4b47-8fb8-8e0d210e9208
# https://officecdn.microsoft.com/pr/f2e724c1-748f-4b47-8fb8-8e0d210e9208

# Office 2021 Perpetual Enterprise Channel
# Production::LTSC2021
# Channel=PerpetualVL2021
# 5030841d-c919-4594-8d2d-84ae4f96e58e
# https://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e

# Office 2024 Perpetual Enterprise Channel
# Production::LTSC2024
# Channel=PerpetualVL2024
# 7983bac0-e531-40cf-be00-fd24fe66619c
# https://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c

# Channel=Current
# Channel=PerpetualVL2019
# Channel=PerpetualVL2021
# Channel=PerpetualVL2024

# "https://config.office.com/api/filelist?Channel=PerpetualVL2021&Version=${_latest_ver}&Arch=${_arch}&Lid=en-US&Lid=zh-CN&Lid=zh-TW"

# https://config.office.com/api/filelist?Channel=PerpetualVL2021
# https://config.office.com/api/filelist?Channel=PerpetualVL2021&Version=16.0.14332.20685
# https://config.office.com/api/filelist?Channel=PerpetualVL2021&Version=16.0.14332.20685&Arch=x64
# https://config.office.com/api/filelist?Channel=PerpetualVL2021&Version=16.0.14332.20685&Arch=x64&Lid=en-US&Lid=zh-CN&Lid=zh-TW

# https://config.office.com/api/filelist?Channel=PerpetualVL2024&Version=16.0.17932.20496&Arch=x64&Lid=en-US&Lid=zh-CN&Lid=zh-TW
# https://config.office.com/api/filelist?Channel=PerpetualVL2024&Version=16.0.17932.20496&Arch=x86&Lid=en-US&Lid=zh-CN&Lid=zh-TW

# https://config.office.com/api/filelist?Channel=Current&Version=16.0.19029.20184&Arch=x64&Lid=en-US&Lid=zh-CN&Lid=zh-TW
# https://config.office.com/api/filelist?Channel=Current&Version=16.0.19029.20184&Arch=x86&Lid=en-US&Lid=zh-CN&Lid=zh-TW

# channel: PerpetualVL2021
# latest_ver: 16.0.14332.20685
# arch: x64
# langs: Lid=en-US&Lid=zh-CN&Lid=zh-TW

###############################################################################

set -euo pipefail

# Configuration constants
readonly SCRIPT_NAME="${0##*/}"
readonly SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
readonly TEMP_BASE_DIR="${TMPDIR:-/tmp}"
readonly DEFAULT_ARCH="x64"
readonly DEFAULT_LANGS="en-US,zh-CN,zh-TW"

# Color codes for output
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly NC='\033[0m'

# Channel configurations - immutable data structure
readonly -A CHANNELS=(
    ["current"]="492350f6-3a01-4f97-b9c0-c7c6ddf67d60:Current"
    ["2024"]="7983bac0-e531-40cf-be00-fd24fe66619c:PerpetualVL2024"
    ["2021"]="5030841d-c919-4594-8d2d-84ae4f96e58e:PerpetualVL2021"
    ["2019"]="f2e724c1-748f-4b47-8fb8-8e0d210e9208:PerpetualVL2019"
)

# Supported languages - constant array for validation
readonly SUPPORTED_LANGS=(
    "af-ZA" "am-ET" "ar-SA" "as-IN" "az-Latn-AZ" "be-BY" "bg-BG" "bn-BD" 
    "bn-IN" "bs-Latn-BA" "ca-ES" "cs-CZ" "cy-GB" "da-DK" "de-DE" "el-GR" 
    "en-GB" "en-US" "es-ES" "es-MX" "et-EE" "eu-ES" "fa-IR" "fil-PH" 
    "fi-FI" "fr-CA" "fr-FR" "ga-IE" "gl-ES" "gu-IN" "ha-Latn-NG" "he-IL" 
    "hi-IN" "hr-HR" "hu-HU" "hy-AM" "id-ID" "ig-NG" "is-IS" "it-IT" 
    "ja-JP" "ka-GE" "kk-KZ" "km-KH" "kn-IN" "kok-IN" "ko-KR" "ky-KG" 
    "lt-LT" "lv-LV" "mi-NZ" "mk-MK" "ml-IN" "mn-MN" "mr-IN" "ms-MY" 
    "mt-MT" "nb-NO" "ne-NP" "nl-NL" "nn-NO" "nso-ZA" "or-IN" "pa-IN" 
    "pl-PL" "ps-AF" "pt-BR" "pt-PT" "quz-PE" "rm-CH" "ro-RO" "ru-RU" 
    "si-LK" "sk-SK" "sl-SI" "sq-AL" "sv-SE" "sw-KE" "ta-IN" "te-IN" 
    "th-TH" "tk-TM" "tn-ZA" "tr-TR" "tt-RU" "uk-UA" "ur-PK" "uz-Latn-UZ" 
    "vi-VN" "xh-ZA" "yo-NG" "zh-CN" "zh-TW" "zu-ZA"
)

# LCID to BCP 47 mapping - constant associative array
readonly -A LCID_MAP=(
    [1025]="ar-SA" [1026]="bg-BG" [1027]="ca-ES" [1028]="zh-TW" 
    [1029]="cs-CZ" [1030]="da-DK" [1031]="de-DE" [1032]="el-GR" 
    [1033]="en-US" [1035]="fi-FI" [1036]="fr-FR" [1037]="he-IL" 
    [1038]="hu-HU" [1039]="is-IS" [1040]="it-IT" [1041]="ja-JP" 
    [1042]="ko-KR" [1043]="nl-NL" [1044]="nb-NO" [1045]="pl-PL" 
    [1046]="pt-BR" [1047]="rm-CH" [1048]="ro-RO" [1049]="ru-RU" 
    [1050]="hr-HR" [1051]="sk-SK" [1052]="sq-AL" [1053]="sv-SE" 
    [1054]="th-TH" [1055]="tr-TR" [1056]="ur-PK" [1057]="id-ID" 
    [1058]="uk-UA" [1059]="be-BY" [1060]="sl-SI" [1061]="et-EE" 
    [1062]="lv-LV" [1063]="lt-LT" [1065]="fa-IR" [1066]="vi-VN" 
    [1067]="hy-AM" [1068]="az-Latn-AZ" [1069]="eu-ES" [1071]="mk-MK" 
    [1074]="tn-ZA" [1076]="xh-ZA" [1077]="zu-ZA" [1078]="af-ZA" 
    [1079]="ka-GE" [1081]="hi-IN" [1082]="mt-MT" [1086]="ms-MY" 
    [1087]="kk-KZ" [1088]="ky-KG" [1089]="sw-KE" [1090]="tk-TM" 
    [1091]="uz-Latn-UZ" [1092]="tt-RU" [1093]="bn-IN" [1094]="pa-IN" 
    [1095]="gu-IN" [1096]="or-IN" [1097]="ta-IN" [1098]="te-IN" 
    [1099]="kn-IN" [1100]="ml-IN" [1101]="as-IN" [1102]="mr-IN" 
    [1104]="mn-MN" [1106]="cy-GB" [1107]="km-KH" [1110]="gl-ES" 
    [1111]="kok-IN" [1115]="si-LK" [1118]="am-ET" [1121]="ne-NP" 
    [1123]="ps-AF" [1124]="fil-PH" [1128]="ha-Latn-NG" [1130]="yo-NG" 
    [1132]="nso-ZA" [1136]="ig-NG" [1153]="mi-NZ" [2052]="zh-CN" 
    [2057]="en-GB" [2058]="es-MX" [2068]="nn-NO" [2070]="pt-PT" 
    [2108]="ga-IE" [2117]="bn-BD" [3082]="es-ES" [3084]="fr-CA" 
    [3179]="quz-PE" [5146]="bs-Latn-BA"
)

# Global state
declare temp_dir=""

# Logging functions with consistent error handling
log() {
    local level="$1"
    shift
    case "$level" in
        "INFO")  echo -e "${BLUE}[INFO]${NC} $*" >&2 ;;
        "WARN")  echo -e "${YELLOW}[WARN]${NC} $*" >&2 ;;
        "ERROR") echo -e "${RED}[ERROR]${NC} $*" >&2 ;;
        "SUCCESS") echo -e "${GREEN}[SUCCESS]${NC} $*" >&2 ;;
        *) echo -e "$*" >&2 ;;
    esac
}

# Cleanup function - robust error handling
cleanup() {
    local exit_code=$?
    if [[ -n "$temp_dir" && -d "$temp_dir" ]]; then
        log INFO "Cleaning up temporary directory: $temp_dir"
        rm -rf "$temp_dir" || log WARN "Failed to remove temporary directory"
    fi
    exit $exit_code
}

trap cleanup EXIT INT TERM

# Check dependencies with detailed error reporting
check_dependencies() {
    local -a deps=("wget" "jq" "openssl" "tar" "7z")
    local -a missing=()
    
    for dep in "${deps[@]}"; do
        if ! command -v "$dep" >/dev/null 2>&1; then
            missing+=("$dep")
        fi
    done
    
    if (( ${#missing[@]} > 0 )); then
        log ERROR "Missing required dependencies: ${missing[*]}"
        log ERROR "Please install them and try again"
        return 1
    fi
    
    return 0
}

# Fetch latest version with proper error handling
fetch_latest_version() {
    local ffn="$1"
    local url="https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/$ffn"
    
    log INFO "Fetching latest version information..."
    
    local response
    if ! response=$(wget --timeout=30 --tries=3 -qO- "$url" 2>/dev/null); then
        log ERROR "Failed to fetch version information from Microsoft servers"
        return 1
    fi
    
    # More robust version extraction
    local version
    if ! version=$(echo "$response" | \
                  grep -i 'AvailableBuild' | \
                  awk -F: '{print $2}' | \
                  sed 's/[", ]//g' | \
                  sort -V | \
                  tail -n 2 | \
                  head -n 1 | \
                  tr -d '\n\r'); then
        log ERROR "Failed to parse version information"
        return 1
    fi
    
    if [[ -z "$version" ]]; then
        log ERROR "Could not extract version information"
        return 1
    fi
    
    echo "$version"
}

# Validate architecture with clear error messages
validate_arch() {
    local arch="$1"
    case "$arch" in
        x64|x86-64|x86_64) echo "x64" ;;
        x86|x32) echo "x86" ;;
        *) 
            log ERROR "Unsupported architecture: $arch"
            log ERROR "Supported architectures: x64, x86"
            return 1
            ;;
    esac
}

# Get channel info with validation
get_channel_info() {
    local input_channel="$1"
    local channel_key
    
    case "${input_channel,,}" in
        current) channel_key="current" ;;
        2024|perpetualvl2024) channel_key="2024" ;;
        2021|perpetualvl2021) channel_key="2021" ;;
        2019|perpetualvl2019) channel_key="2019" ;;
        *)
            log ERROR "Unsupported channel: $input_channel"
            log ERROR "Supported channels: current, 2024, 2021, 2019"
            return 1
            ;;
    esac
    
    echo "${CHANNELS[$channel_key]}"
}

# Language validation with efficient lookup
is_language_supported() {
    local lang="$1"
    local supported_lang
    
    for supported_lang in "${SUPPORTED_LANGS[@]}"; do
        if [[ "$lang" == "$supported_lang" ]]; then
            return 0
        fi
    done
    return 1
}

# Normalize language code with case handling
normalize_language() {
    local lang="$1"
    
    # Handle standard patterns
    if [[ "$lang" =~ ^([a-zA-Z]{2})-([a-zA-Z]{2})$ ]]; then
        echo "${BASH_REMATCH[1],,}-${BASH_REMATCH[2]^^}"
    elif [[ "$lang" =~ ^([a-zA-Z]{2})-([a-zA-Z]{4})-([a-zA-Z]{2})$ ]]; then
        echo "${BASH_REMATCH[1],,}-${BASH_REMATCH[2]^}-${BASH_REMATCH[3]^^}"
    else
        echo "$lang"
    fi
}

# Validate and normalize languages
validate_languages() {
    local input_langs="$1"
    
    # Skip if already formatted
    if [[ "$input_langs" == *"Lid="* ]]; then
        echo "$input_langs"
        return 0
    fi
    
    local -a validated_langs=()
    local -a invalid_langs=()
    local lang
    
    IFS=',' read -ra lang_array <<< "$input_langs"
    
    for lang in "${lang_array[@]}"; do
        lang=$(echo "$lang" | tr -d ' ')
        [[ -z "$lang" ]] && continue
        
        lang=$(normalize_language "$lang")
        
        if is_language_supported "$lang"; then
            validated_langs+=("$lang")
        else
            invalid_langs+=("$lang")
        fi
    done
    
    if (( ${#invalid_langs[@]} > 0 )); then
        log ERROR "Invalid language codes: ${invalid_langs[*]}"
        echo
        log ERROR "Supported languages: ${SUPPORTED_LANGS[*]}"
        return 1
    fi
    
    local IFS=','
    echo "${validated_langs[*]}"
}

# Format language parameters for API
format_language_params() {
    local input_langs="$1"
    
    if [[ "$input_langs" == *"Lid="* ]]; then
        echo "$input_langs"
        return 0
    fi
    
    local -a formatted_langs=()
    local lang
    
    IFS=',' read -ra lang_array <<< "$input_langs"
    
    for lang in "${lang_array[@]}"; do
        lang=$(echo "$lang" | tr -d ' ')
        [[ -n "$lang" ]] && formatted_langs+=("Lid=$lang")
    done
    
    local IFS='&'
    echo "${formatted_langs[*]}"
}

# Download file list with error handling
download_file_list() {
    local channel="$1" version="$2" arch="$3" langs="$4" temp_file="$5"
    
    local formatted_langs
    formatted_langs=$(format_language_params "$langs")
    
    local url="https://config.office.com/api/filelist?Channel=$channel&Version=$version&Arch=$arch&$formatted_langs"
    
    log INFO "Downloading file list from Microsoft..."
    log INFO "URL: $url"
    
    if ! wget --timeout=30 --tries=3 -qO- "$url" | \
         sed 's|http:|https:|g' | \
         jq . > "$temp_file" 2>/dev/null; then
        log ERROR "Failed to download file list"
        return 1
    fi
    
    # Apply fixes for Office 2024 if needed
    if [[ "$channel" == "PerpetualVL2024" ]]; then
        fix_2024_channel_urls "$temp_file" "$version" "$channel"
    fi
    
    return 0
}

# Fix URLs for Office 2024 channel
fix_2024_channel_urls() {
    local file="$1" version="$2" channel="$3"
    
    log INFO "Applying fixes for Office 2024 channel..."
    
    local ffn_old ffn_new channel_old ver_old
    
    ffn_old=$(grep -i 'baseUrl' "$file" | \
              sed 's|.*/pr/||g; s|[", ]||g' | \
              head -n 1 | tr -d '\n\r')
    ffn_new="7983bac0-e531-40cf-be00-fd24fe66619c"
    
    channel_old=$(grep -i channel "$file" | \
                  sed 's|.*: ||g; s|[", ]||g' | \
                  head -n 1 | tr -d '\n\r')
    
    ver_old=$(grep -i version "$file" | \
              sed 's|.*: ||g; s|[", ]||g' | \
              head -n 1 | tr -d '\n\r')
    
    sed -i "s|$ffn_old|$ffn_new|g; s|$channel_old|$channel|g; s|$ver_old|$version|g" "$file"
}

# Convert LCID to BCP 47 language codes
convert_lcids_to_languages() {
    local lcids="$1"
    local -a language_codes=()
    local lcid
    
    IFS=',' read -ra lcid_array <<< "$lcids"
    
    for lcid in "${lcid_array[@]}"; do
        lcid=$(echo "$lcid" | tr -d ' ')
        
        if [[ -n "${LCID_MAP[$lcid]:-}" ]]; then
            language_codes+=("${LCID_MAP[$lcid]}")
        else
            language_codes+=("$lcid")
            log WARN "Unknown LCID: $lcid, keeping original value"
        fi
    done
    
    local IFS=','
    echo "${language_codes[*]}"
}

# Download Office files with progress tracking
download_office_files() {
    local json_file="$1" download_dir="$2"
    
    log INFO "Starting Office files download..."
    
    local total_files
    if ! total_files=$(jq '.files | length' "$json_file" 2>/dev/null); then
        log ERROR "Failed to parse file list"
        return 1
    fi
    
    log INFO "Total files to download: $total_files"
    
    local count=0
    while IFS= read -r file; do
        ((count++))
        
        local file_url relative_path file_name dir_path
        if ! file_url=$(echo "$file" | jq -r '.url' | sed 's|http:|https:|g') ||
           ! relative_path=$(echo "$file" | jq -r '.relativePath') ||
           ! file_name=$(echo "$file" | jq -r '.name'); then
            log ERROR "Failed to parse file information"
            return 1
        fi
        
        # Clean up the relative path
        relative_path="${relative_path#/}"
        relative_path="${relative_path%/}"
        
        if [[ -n "$relative_path" ]]; then
            dir_path="$download_dir/$relative_path"
        else
            dir_path="$download_dir"
        fi
        
        log INFO "Downloading file $count/$total_files: $file_name"
        
        # Create directory structure
        mkdir -p "$dir_path" || {
            log ERROR "Failed to create directory: $dir_path"
            return 1
        }
        
        # Download file with retry logic
        if ! wget -c -t 9 -T 9 --progress=bar:force \
                  "$file_url" -O "$dir_path/$file_name" 2>/dev/null; then
            log ERROR "Failed to download: $file_name"
            return 1
        fi
        
    done < <(jq -c '.files[]' "$json_file" 2>/dev/null)
    
    log SUCCESS "All files downloaded successfully"
    return 0
}

# Generate checksums and metadata
generate_metadata() {
    local target_dir="$1" json_file="$2"
    
    log INFO "Generating file metadata..."
    
    pushd "$target_dir" >/dev/null || {
        log ERROR "Failed to change to target directory"
        return 1
    }
    
    # Find office directory (case insensitive)
    local office_dir
    if [[ -d "office" ]]; then
        office_dir="office"
    elif [[ -d "Office" ]]; then
        office_dir="Office"
    else
        log ERROR "No office directory found"
        popd >/dev/null
        return 1
    fi
    
    # Generate file list
    find "$office_dir" -type f | sort -V > files.txt
    
    # Generate checksums
    log INFO "Computing SHA256 checksums..."
    while IFS= read -r file; do
        [[ -f "$file" ]] && openssl dgst -r -sha256 "$file"
    done < files.txt > sha256sums.txt
    
    # Generate version info
    local lcids languages
    if ! lcids=$(jq -r '.lcids' "$json_file" 2>/dev/null); then
        log ERROR "Failed to extract LCIDs from file list"
        popd >/dev/null
        return 1
    fi
    
    languages=$(convert_lcids_to_languages "$lcids")
    
    {
        printf '"channel":\t"%s"\n' "$(jq -r '.channel' "$json_file")"
        printf '"version":\t"%s"\n' "$(jq -r '.version' "$json_file")"
        printf '"architecture":\t"%s"\n' "$(jq -r '.architecture' "$json_file")"
        printf '"languages":\t"%s"\n' "$languages"
        printf '"baseUrl":\t"%s"\n' "$(jq -r '.baseUrl' "$json_file")"
    } > .version
    
    # Download setup.exe
    log INFO "Downloading setup.exe..."
    if ! wget -c -t 9 -T 9 --progress=bar:force \
             'https://officecdn.microsoft.com/pr/wsus/setup.exe' \
             -O setup.exe 2>/dev/null; then
        log WARN "Failed to download setup.exe"
        popd >/dev/null
        return 1
    fi
    
    popd >/dev/null
    return 0
}

# Create final package
create_package() {
    local package_dir="$1" output_path="$2"
    
    log INFO "Creating package archive..."
    
    if ! (cd "$(dirname "$package_dir")" && \
         /usr/bin/7z a -r -mmt=$(nproc) -mx9 -t7z -v1800m "$output_path.7z" \
             "$(basename "$package_dir")"); then
        log ERROR "Failed to create package archive"
        return 1
    fi

    # Generate package checksum
    pushd "$(dirname "$output_path")" >/dev/null || return 1
    openssl dgst -r -sha256 "$(basename "$output_path").7z"* > \
        "$(basename "$output_path").7z.sha256"
    popd >/dev/null
    
    #log SUCCESS "Package created: $output_path.7z"
    log INFO "Package checksum: $output_path.7z.sha256"
    return 0
}

# Display usage information
usage() {
    cat << 'EOF'
Usage: download-office.sh [OPTIONS]

Download Microsoft Office installation files from Microsoft CDN.

OPTIONS:
    -c, --channel CHANNEL    Office channel (current, 2024, 2021, 2019)
    -a, --arch ARCH         Architecture (x64, x86) [default: x64]
    -l, --langs LANGS       Language codes [default: en-US,zh-CN,zh-TW]
    -o, --output DIR        Output directory [default: current directory]
    -v, --verbose           Enable verbose output
    -h, --help              Show this help message

EXAMPLES:
    download-office.sh -c current -a x64
    download-office.sh -c 2021 -a x86 -o /downloads
    download-office.sh --channel 2024 --langs "en-US,zh-CN"

CHANNELS:
    current     - Current Channel (latest features)
    2024        - Office 2024 Perpetual Enterprise Channel  
    2021        - Office 2021 Perpetual Enterprise Channel
    2019        - Office 2019 Perpetual Enterprise Channel

ARCHITECTURES:
    x64         - 64-bit (recommended)
    x86         - 32-bit (legacy systems)

SUPPORTED LANGUAGES (94 total):
    af-ZA am-ET ar-SA as-IN az-Latn-AZ be-BY bg-BG bn-BD bn-IN bs-Latn-BA
    ca-ES cs-CZ cy-GB da-DK de-DE el-GR en-GB en-US es-ES es-MX et-EE eu-ES
    fa-IR fil-PH fi-FI fr-CA fr-FR ga-IE gl-ES gu-IN ha-Latn-NG he-IL hi-IN
    hr-HR hu-HU hy-AM id-ID ig-NG is-IS it-IT ja-JP ka-GE kk-KZ km-KH kn-IN
    kok-IN ko-KR ky-KG lt-LT lv-LV mi-NZ mk-MK ml-IN mn-MN mr-IN ms-MY mt-MT
    nb-NO ne-NP nl-NL nn-NO nso-ZA or-IN pa-IN pl-PL ps-AF pt-BR pt-PT quz-PE
    rm-CH ro-RO ru-RU si-LK sk-SK sl-SI sq-AL sv-SE sw-KE ta-IN te-IN th-TH
    tk-TM tn-ZA tr-TR tt-RU uk-UA ur-PK uz-Latn-UZ vi-VN xh-ZA yo-NG zh-CN
    zh-TW zu-ZA

EOF
}

# Main function with structured error handling
main() {
    local channel="" arch="$DEFAULT_ARCH" langs="$DEFAULT_LANGS"
    local output_dir="$(pwd)" verbose=false
    
    # Parse command line arguments
    while (( $# > 0 )); do
        case $1 in
            -c|--channel)
                [[ -n "${2:-}" ]] || { log ERROR "Channel option requires an argument"; return 1; }
                channel="$2"
                shift 2
                ;;
            -a|--arch)
                [[ -n "${2:-}" ]] || { log ERROR "Architecture option requires an argument"; return 1; }
                arch="$2"
                shift 2
                ;;
            -l|--langs)
                [[ -n "${2:-}" ]] || { log ERROR "Languages option requires an argument"; return 1; }
                langs="$2"
                shift 2
                ;;
            -o|--output)
                [[ -n "${2:-}" ]] || { log ERROR "Output option requires an argument"; return 1; }
                output_dir="$2"
                shift 2
                ;;
            -v|--verbose)
                verbose=true
                shift
                ;;
            -h|--help)
                usage
                exit 0
                ;;
            *)
                log ERROR "Unknown option: $1"
                usage
                exit 1
                ;;
        esac
    done
    
    # Validate required arguments
    if [[ -z "$channel" ]]; then
        log ERROR "Channel is required"
        usage
        exit 1
    fi

    _install_jq
    _install_7z

    # Check dependencies
    check_dependencies || exit 1
    
    # Validate and normalize inputs
    langs=$(validate_languages "$langs") || exit 1
    arch=$(validate_arch "$arch") || exit 1
    
    local channel_info
    channel_info=$(get_channel_info "$channel") || exit 1
    
    local ffn="${channel_info%:*}"
    local channel_name="${channel_info#*:}"
    local channel_short="${channel,,}"
    
    log INFO "Starting Office download process"
    log INFO "Channel: $channel_name"
    log INFO "Architecture: $arch"
    log INFO "Languages: $langs"
    log INFO "Output directory: $output_dir"
    
    # Create output directory
    mkdir -p "$output_dir" || {
        log ERROR "Failed to create output directory: $output_dir"
        exit 1
    }
    
    # Fetch latest version
    local latest_ver
    latest_ver=$(fetch_latest_version "$ffn") || exit 1
    log SUCCESS "Latest version: $latest_ver"
    
    # Setup working directory
    temp_dir=$(mktemp -d -p "$TEMP_BASE_DIR" "office-download.XXXXXX") || {
        log ERROR "Failed to create temporary directory"
        exit 1
    }
    
    local work_dir="$temp_dir/Office-${channel_short}-${latest_ver}-${arch}"
    local json_file="$temp_dir/filelist.json"
    
    # Download file list
    download_file_list "$channel_name" "$latest_ver" "$arch" "$langs" "$json_file" || exit 1
    
    # Create target directory
    mkdir -p "$work_dir" || {
        log ERROR "Failed to create work directory"
        exit 1
    }
    
    # Download Office files
    download_office_files "$json_file" "$work_dir" || exit 1
    
    # Generate metadata
    generate_metadata "$work_dir" "$json_file" || exit 1
    
    # Create final package
    local output_path="$output_dir/Office-${channel_short}-${latest_ver}-${arch}"
    create_package "$work_dir" "$output_path" || exit 1
    
    # Display summary
    echo
    log SUCCESS "Download completed successfully!"
    echo
    log INFO "Package details:"
    if [[ -f "$work_dir/.version" ]]; then
        cat "$work_dir/.version"
    fi
    echo
    log INFO "Files created:"
    #log INFO "  Archive: $output_path.7z"
    #log INFO "  Checksum: $output_path.7z.sha256"
    # Display file size
    #if command -v du >/dev/null 2>&1; then
    #    local size
    #    size=$(du -h "$output_path.7z" 2>/dev/null | cut -f1)
    #    [[ -n "$size" ]] && log INFO "  Size: $size"
    #fi
    
    /bin/ls -lh "$output_path.7z"*
    
    echo
    return 0
}

# Entry point with error handling
if [[ "${BASH_SOURCE[0]}" == "${0}" ]]; then
    main "$@"
fi
exit

