//! Windows: Serif Affinity は COM/公式 CLI が限定的なため、インストール済みの `.exe` を起動してファイルを開く。
use anyhow::{Context, Result};
use std::path::{Path, PathBuf};
use std::process::Command;

use super::affinity::AffinityApp;

fn env_exe_for(app: &AffinityApp) -> Option<PathBuf> {
    let key = match app {
        AffinityApp::Photo => "AFFINITY_PHOTO_EXE",
        AffinityApp::Designer => "AFFINITY_DESIGNER_EXE",
        AffinityApp::Publisher => "AFFINITY_PUBLISHER_EXE",
    };
    std::env::var_os(key).map(PathBuf::from)
}

fn program_files() -> PathBuf {
    std::env::var_os("ProgramFiles")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(r"C:\Program Files"))
}

fn program_files_x86() -> PathBuf {
    std::env::var_os("ProgramFiles(x86)")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(r"C:\Program Files (x86)"))
}

/// 一般的な MSI インストール（Affinity 2）とレガシー候補を順に試す。
fn candidate_exes(app: &AffinityApp) -> Vec<PathBuf> {
    let pf = program_files();
    let pf86 = program_files_x86();
    match app {
        AffinityApp::Photo => vec![
            pf.join(r"Affinity\Photo 2\Affinity Photo.exe"),
            pf.join(r"Affinity\Affinity Photo\Affinity Photo.exe"),
            pf86.join(r"Affinity\Photo 2\Affinity Photo.exe"),
        ],
        AffinityApp::Designer => vec![
            pf.join(r"Affinity\Designer 2\Affinity Designer.exe"),
            pf.join(r"Affinity\Affinity Designer\Affinity Designer.exe"),
            pf86.join(r"Affinity\Designer 2\Affinity Designer.exe"),
        ],
        AffinityApp::Publisher => vec![
            pf.join(r"Affinity\Publisher 2\Affinity Publisher.exe"),
            pf.join(r"Affinity\Affinity Publisher\Affinity Publisher.exe"),
            pf86.join(r"Affinity\Publisher 2\Affinity Publisher.exe"),
        ],
    }
}

pub fn resolve_affinity_exe(app: &AffinityApp) -> Result<PathBuf> {
    if let Some(p) = env_exe_for(app) {
        if p.is_file() {
            return Ok(p);
        }
    }
    for c in candidate_exes(app) {
        if c.is_file() {
            return Ok(c);
        }
    }
    anyhow::bail!(
        "Affinity の実行ファイルが見つかりません。{} を設定するか、Affinity を既定の場所にインストールしてください。",
        match app {
            AffinityApp::Photo => "AFFINITY_PHOTO_EXE",
            AffinityApp::Designer => "AFFINITY_DESIGNER_EXE",
            AffinityApp::Publisher => "AFFINITY_PUBLISHER_EXE",
        }
    )
}

pub fn app_display_name(app: &AffinityApp) -> &'static str {
    match app {
        AffinityApp::Photo => "Affinity Photo",
        AffinityApp::Designer => "Affinity Designer",
        AffinityApp::Publisher => "Affinity Publisher",
    }
}

/// 指定アプリでファイルを開く（プロセスはデタッチ）。
pub fn open_file_path(path: &str, app: Option<&AffinityApp>) -> Result<String> {
    let path_buf = PathBuf::from(path);
    let canonical = std::fs::canonicalize(&path_buf)
        .with_context(|| format!("パスの正規化に失敗しました: {}", path))?;
    let resolved_app = app.cloned().unwrap_or_else(|| AffinityApp::from_file_path(path));
    let exe = resolve_affinity_exe(&resolved_app)?;
    Command::new(&exe)
        .arg(&canonical)
        .spawn()
        .with_context(|| format!("ファイルを開けませんでした: {}", exe.display()))?;
    Ok(app_display_name(&resolved_app).to_string())
}

/// インストールされている最初のアプリ名（MCP 応答用）。
/// 新規: 実行ファイルのみ起動（空白ドキュメントはアプリの既定動作に依存）。
pub fn launch_app(app: &AffinityApp) -> Result<()> {
    let exe = resolve_affinity_exe(app)?;
    Command::new(&exe)
        .spawn()
        .with_context(|| format!("Affinity を起動できませんでした: {}", exe.display()))?;
    Ok(())
}

pub fn open_path_with_photo_fallback(svg_path: &Path) -> Result<String> {
    let canonical = std::fs::canonicalize(svg_path)
        .unwrap_or_else(|_| svg_path.to_path_buf());
    match resolve_affinity_exe(&AffinityApp::Photo) {
        Ok(exe) => {
            Command::new(&exe)
                .arg(&canonical)
                .spawn()
                .with_context(|| format!("ファイルを開けませんでした: {}", exe.display()))?;
            Ok("Affinity Photo".to_string())
        }
        Err(_) => {
            let exe = resolve_affinity_exe(&AffinityApp::Designer).context("Photo/Designer いずれの exe も見つかりません")?;
            Command::new(&exe)
                .arg(&canonical)
                .spawn()
                .with_context(|| format!("ファイルを開けませんでした: {}", exe.display()))?;
            Ok("Affinity Designer".to_string())
        }
    }
}
