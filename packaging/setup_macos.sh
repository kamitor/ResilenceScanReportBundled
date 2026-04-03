#!/bin/bash
# setup_macos.sh -- installs R, Quarto, TinyTeX and required R packages on macOS.
# Called by the macOS pkg postinstall script in the background.
# Can also be run manually:
#   sudo /Applications/ResilenceScanReportBuilder.app/Contents/MacOS/_internal/setup_macos.sh

set -e

QUARTO_VERSION="1.6.39"
R_VERSION="4.5.1"
ARCH="$(uname -m)"   # arm64 or x86_64

# Resolve install directory: the script lives inside the .app bundle
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
INSTALL_DIR="$SCRIPT_DIR"
R_LIB="$INSTALL_DIR/r-library"

# R binary path (standard macOS R.framework location)
RSCRIPT="/Library/Frameworks/R.framework/Resources/bin/Rscript"

log() { echo "[SETUP] $1"; }

# Default to FAIL; overwritten to PASS only when all steps complete cleanly.
SETUP_RESULT="FAIL"

# Write running flag immediately so the app knows setup is in progress.
echo "running" > "$INSTALL_DIR/setup_running.flag"
chmod a+r "$INSTALL_DIR/setup_running.flag" 2>/dev/null || true

# On any exit (normal or error) write the completion flag and notify the user.
_on_exit() {
    echo "$SETUP_RESULT" > "$INSTALL_DIR/setup_complete.flag"
    chmod a+r "$INSTALL_DIR/setup_complete.flag" 2>/dev/null || true
    rm -f "$INSTALL_DIR/setup_running.flag"
    if [ "$SETUP_RESULT" = "PASS" ]; then
        osascript -e 'display notification "Setup complete. You can now generate reports." with title "ResilienceScan"' \
            2>/dev/null || true
    else
        osascript -e 'display notification "Setup finished with errors. Check the log for details." with title "ResilienceScan"' \
            2>/dev/null || true
    fi
}
trap _on_exit EXIT

# -- R --

_r_version_ok() {
    [ -f "$RSCRIPT" ] || return 1
    local ver
    ver=$("$RSCRIPT" --version 2>&1 | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1 || echo "0.0.0")
    local major minor patch
    major=$(echo "$ver" | cut -d. -f1)
    minor=$(echo "$ver" | cut -d. -f2)
    patch=$(echo "$ver" | cut -d. -f3)
    # Accept R >= 4.5.1
    [ "$major" -gt 4 ] || \
        ([ "$major" -eq 4 ] && [ "$minor" -gt 5 ]) || \
        ([ "$major" -eq 4 ] && [ "$minor" -eq 5 ] && [ "$patch" -ge 1 ])
}

if ! _r_version_ok; then
    log "Installing R $R_VERSION for macOS ($ARCH)..."
    TMP=$(mktemp -d)
    if [ "$ARCH" = "arm64" ]; then
        R_URL="https://cran.r-project.org/bin/macosx/big-sur-arm64/base/R-${R_VERSION}-arm64.pkg"
    else
        R_URL="https://cran.r-project.org/bin/macosx/big-sur-x86_64/base/R-${R_VERSION}-x86_64.pkg"
    fi
    curl -fsSL -o "$TMP/R.pkg" "$R_URL"
    installer -pkg "$TMP/R.pkg" -target /
    rm -rf "$TMP"
    log "R $R_VERSION installed."
    R_UPGRADED=true
else
    INSTALLED_R=$("$RSCRIPT" --version 2>&1 | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1 || echo "unknown")
    log "R $INSTALLED_R already meets requirements (>= $R_VERSION) -- skipping."
    R_UPGRADED=false
fi

# -- Quarto --

if ! command -v quarto &>/dev/null; then
    log "Installing Quarto $QUARTO_VERSION..."
    TMP=$(mktemp -d)
    curl -fsSL -o "$TMP/quarto.pkg" \
        "https://github.com/quarto-dev/quarto-cli/releases/download/v${QUARTO_VERSION}/quarto-${QUARTO_VERSION}-macos.pkg"
    installer -pkg "$TMP/quarto.pkg" -target /
    rm -rf "$TMP"
    # Quarto installs to /usr/local/bin/quarto
    export PATH="/usr/local/bin:$PATH"
    log "Quarto $QUARTO_VERSION installed."
else
    log "Quarto already present -- skipping."
fi

QUARTO_BIN=$(command -v quarto 2>/dev/null || echo "/usr/local/bin/quarto")

# -- TinyTeX --

if ! command -v tlmgr &>/dev/null; then
    log "Installing TinyTeX via Quarto..."
    "$QUARTO_BIN" install tinytex --no-prompt

    # Locate the TinyTeX bin directory.
    # Quarto 1.4+ installs to ~/Library/Application Support/quarto/tools/tinytex/.
    # When running as root, HOME=/var/root.
    TINYTEX_BIN=""
    for candidate in \
        "${HOME}/Library/Application Support/quarto/tools/tinytex/bin/universal-darwin" \
        "${HOME}/Library/Application Support/quarto/tools/tinytex/bin/aarch64-darwin" \
        "${HOME}/Library/Application Support/quarto/tools/tinytex/bin/x86_64-darwin" \
        "${HOME}/.TinyTeX/bin/universal-darwin" \
        "${HOME}/.TinyTeX/bin/aarch64-darwin" \
        "${HOME}/.TinyTeX/bin/x86_64-darwin" \
        "/usr/local/share/tinytex/bin/universal-darwin" \
        "/usr/local/share/tinytex/bin/aarch64-darwin"; do
        if [ -d "$candidate" ]; then
            TINYTEX_BIN="$candidate"
            break
        fi
    done

    if [ -n "$TINYTEX_BIN" ]; then
        log "Symlinking TinyTeX binaries from $TINYTEX_BIN to /usr/local/bin"
        for bin in tlmgr pdflatex xelatex lualatex luatex tex latex; do
            [ -e "$TINYTEX_BIN/$bin" ] && ln -sf "$TINYTEX_BIN/$bin" "/usr/local/bin/$bin" || true
        done

        # Move TinyTeX to a system-wide location so all users can access it.
        TINYTEX_ROOT="$(dirname "$(dirname "$TINYTEX_BIN")")"
        SYSTEM_TINYTEX="/usr/local/share/tinytex"
        if [ "$TINYTEX_ROOT" != "$SYSTEM_TINYTEX" ]; then
            log "Copying TinyTeX to system location: $SYSTEM_TINYTEX"
            rm -rf "$SYSTEM_TINYTEX" 2>/dev/null || true
            cp -r "$TINYTEX_ROOT" "$SYSTEM_TINYTEX" || true
            # Re-symlink from the system copy
            NEW_BIN="$SYSTEM_TINYTEX/bin/$(basename "$TINYTEX_BIN")"
            if [ -d "$NEW_BIN" ]; then
                for bin in tlmgr pdflatex xelatex lualatex luatex tex latex; do
                    [ -e "$NEW_BIN/$bin" ] && ln -sf "$NEW_BIN/$bin" "/usr/local/bin/$bin" || true
                done
            fi
        fi
    else
        log "WARNING: TinyTeX bin dir not found after install"
    fi
else
    log "TinyTeX already present -- skipping."
fi

# -- LaTeX packages --

TLMGR=$(command -v tlmgr 2>/dev/null || true)
if [ -n "$TLMGR" ]; then
    log "Installing required LaTeX packages..."
    "$TLMGR" install \
        pgf xcolor colortbl booktabs multirow float wrapfig pdflscape geometry \
        preprint graphics tabu threeparttable threeparttablex ulem makecell \
        environ trimspaces caption hyperref setspace fancyhdr microtype lm \
        needspace varwidth mdwtools xstring tools 2>/dev/null || true
    "$TLMGR" generate --rebuild-sys 2>/dev/null || true

    # capt-of stub (not available via tlmgr)
    TEXMF_LOCAL=$("$TLMGR" conf | grep TEXMFLOCAL | head -1 | awk '{print $NF}' || echo "")
    if [ -z "$TEXMF_LOCAL" ]; then
        TEXMF_LOCAL="/usr/local/texlive/texmf-local"
    fi
    CAPT_OF_DIR="$TEXMF_LOCAL/tex/latex/capt-of"
    mkdir -p "$CAPT_OF_DIR"
    if [ ! -f "$CAPT_OF_DIR/capt-of.sty" ]; then
        cat > "$CAPT_OF_DIR/capt-of.sty" << 'STY'
\NeedsTeXFormat{LaTeX2e}
\ProvidesPackage{capt-of}[2023/01/01 capt-of stub]
\RequirePackage{caption}
\newcommand{\captionof}[2][]{\caption{#2}}
STY
        mktexlsr 2>/dev/null || "$TLMGR" generate --rebuild-sys 2>/dev/null || true
    fi
else
    log "WARNING: tlmgr not found -- skipping LaTeX package installation."
fi

# -- R packages --

log "Installing R packages into $R_LIB..."

# If R was upgraded, wipe the old r-library to force a clean rebuild.
if [ "${R_UPGRADED:-false}" = "true" ] && [ -d "$R_LIB" ]; then
    log "R was upgraded -- removing stale r-library: $R_LIB"
    rm -rf "$R_LIB"
fi

mkdir -p "$R_LIB"

if ! touch "$R_LIB/_write_test_" 2>/dev/null; then
    log "WARNING: R library not writable -- attempting to fix permissions..."
    chmod -R u+w "$R_LIB" 2>/dev/null || true
    if ! touch "$R_LIB/_write_test_" 2>/dev/null; then
        log "ERROR: R library still not writable. Package install will fail."
    fi
fi
rm -f "$R_LIB/_write_test_"

NCPUS=$(sysctl -n hw.ncpu 2>/dev/null || echo 2)
R_PKGS="'readr','dplyr','stringr','tidyr','ggplot2','knitr','fmsb','scales','viridis','patchwork','RColorBrewer','gridExtra','png','lubridate','kableExtra','rmarkdown','jsonlite','ggrepel','cowplot'"

"$RSCRIPT" -e "
  pkgs <- c($R_PKGS)
  install.packages(pkgs, lib='$R_LIB', repos='https://cloud.r-project.org', Ncpus=${NCPUS}, quiet=FALSE)
"

log "Verifying R packages..."

if ! [ -f "$RSCRIPT" ]; then
    log "ERROR: Rscript not found at $RSCRIPT -- cannot verify R packages."
    SETUP_RESULT="FAIL"
else
    MISSING=$("$RSCRIPT" --no-save -e "
  .libPaths(c('$R_LIB', .libPaths()))
  pkgs <- c($R_PKGS)
  bad <- pkgs[!sapply(pkgs, requireNamespace, quietly=TRUE)]
  cat(paste(bad, collapse=' '))
" 2>&1) || { log "WARNING: Rscript package check command failed"; MISSING="check_failed"; }

    if [ -n "$MISSING" ]; then
        log "Retrying packages that failed to load: $MISSING"
        for pkg in $MISSING; do
            "$RSCRIPT" -e "install.packages('$pkg', lib='$R_LIB', repos='https://cloud.r-project.org')" 2>&1 || true
        done
        STILL_MISSING=$("$RSCRIPT" --no-save -e "
      .libPaths(c('$R_LIB', .libPaths()))
      pkgs <- c($R_PKGS)
      bad <- pkgs[!sapply(pkgs, requireNamespace, quietly=TRUE)]
      cat(paste(bad, collapse=' '))
" 2>&1) || { log "WARNING: Rscript final check failed"; STILL_MISSING="check_failed"; }
        if [ -n "$STILL_MISSING" ]; then
            log "ERROR: R packages still not loadable after retry: $STILL_MISSING"
            SETUP_RESULT="FAIL"
        else
            log "R package retry succeeded -- all packages installed and loadable."
            SETUP_RESULT="PASS"
        fi
    else
        log "R package verification: all packages installed and loadable."
        SETUP_RESULT="PASS"
    fi
fi

# Ensure the R library is readable by all users
chmod -R a+rX "$R_LIB"

log "Dependency setup complete."
