//! Windows: Serif Affinity は COM/公式 CLI が限定的なため、インストール済みの `.exe` を起動してファイルを開く。
//! アクティブドキュメントはメインウィンドウのキャプション（タイトルバー）から推定する。
use anyhow::{Context, Result};
use std::path::Path;
use std::path::PathBuf;
use std::process::Command;

use windows::core::PWSTR;
use windows::Win32::Foundation::{CloseHandle, BOOL, HWND, LPARAM, TRUE};
use windows::Win32::System::Threading::{
    OpenProcess, QueryFullProcessImageNameW, PROCESS_NAME_WIN32, PROCESS_QUERY_LIMITED_INFORMATION,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetClassNameW, GetForegroundWindow, GetWindowTextW, GetWindowThreadProcessId,
    IsWindowVisible,
};

use super::affinity::{ActiveDocumentInfo, AffinityApp};

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

pub fn resolve_affinity_exe(app: &AffinityApp) -> anyhow::Result<PathBuf> {
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
pub fn open_file_path(path: &str, app: Option<&AffinityApp>) -> anyhow::Result<String> {
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

/// 新規: 実行ファイルのみ起動（空白ドキュメントはアプリの既定動作に依存）。
pub fn launch_app(app: &AffinityApp) -> anyhow::Result<()> {
    let exe = resolve_affinity_exe(app)?;
    Command::new(&exe)
        .spawn()
        .with_context(|| format!("Affinity を起動できませんでした: {}", exe.display()))?;
    Ok(())
}

pub fn open_path_with_photo_fallback(svg_path: &Path) -> anyhow::Result<String> {
    let canonical = std::fs::canonicalize(svg_path).unwrap_or_else(|_| svg_path.to_path_buf());
    match resolve_affinity_exe(&AffinityApp::Photo) {
        Ok(exe) => {
            Command::new(&exe)
                .arg(&canonical)
                .spawn()
                .with_context(|| format!("ファイルを開けませんでした: {}", exe.display()))?;
            Ok("Affinity Photo".to_string())
        }
        Err(_) => {
            let exe = resolve_affinity_exe(&AffinityApp::Designer)
                .context("Photo/Designer いずれの exe も見つかりません")?;
            Command::new(&exe)
                .arg(&canonical)
                .spawn()
                .with_context(|| format!("ファイルを開けませんでした: {}", exe.display()))?;
            Ok("Affinity Designer".to_string())
        }
    }
}

// --- Active document (window caption) ---

fn is_affinity_exe_path(path: &str) -> bool {
    let p = path.replace('/', "\\").to_lowercase();
    if !p.ends_with(".exe") {
        return false;
    }
    let file = p.rsplit('\\').next().unwrap_or("");
    matches!(
        file,
        "affinity.exe" | "affinity photo.exe" | "affinity designer.exe" | "affinity publisher.exe"
    )
}

/// Store 版などで exe パスが取れないことがあるので、タイトルバーで判定するフォールバック。
/// Affinity 3 はドック時タイトルが単に `Affinity`、アンドック時はドキュメント名のみ（サフィックスなし）のことがある。
fn title_suggests_affinity_window(title: &str) -> bool {
    let t = title.trim().replace('–', "-").replace('—', "-").to_lowercase();
    if t == "affinity" {
        return true;
    }
    if t.ends_with(".afphoto") || t.ends_with(".afdesign") || t.ends_with(".afpub") {
        return true;
    }
    t.contains(" - affinity photo")
        || t.contains(" - affinity designer")
        || t.contains(" - affinity publisher")
        || t.ends_with(" - affinity")
        || t.contains(" - affinity ")
}

unsafe fn query_exe_path(pid: u32) -> Option<String> {
    let h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid).ok()?;
    let mut buf = vec![0u16; 4096];
    let mut size = buf.len() as u32;
    let r = QueryFullProcessImageNameW(h, PROCESS_NAME_WIN32, PWSTR(buf.as_mut_ptr()), &mut size);
    let _ = CloseHandle(h);
    r.ok()?;
    if size == 0 {
        return None;
    }
    Some(String::from_utf16_lossy(&buf[..size as usize]).to_string())
}

/// タイトルがアプリ名のみ（ドキュメント名がない起動直後など）ならドキュメントなしとみなす。
fn is_idle_app_title(title: &str) -> bool {
    matches!(
        title.trim(),
        "Affinity Photo" | "Affinity Designer" | "Affinity Publisher" | "Affinity"
    )
}

fn parse_affinity_caption(title: &str) -> ActiveDocumentInfo {
    let normalized = title.replace('–', "-").replace('—', "-");
    let title = normalized.trim();

    if is_idle_app_title(title) {
        return ActiveDocumentInfo {
            is_open: false,
            name: None,
            path: None,
        };
    }

    // 長いサフィックスを先に（Affinity 3 は " - Affinity" のみの場合あり）
    const SUFFIXES: &[&str] = &[
        " - Affinity Photo",
        " - Affinity Designer",
        " - Affinity Publisher",
        " - Affinity",
    ];
    for suf in SUFFIXES {
        if let Some(rest) = title.strip_suffix(suf) {
            let rest = rest.trim();
            if rest.is_empty() {
                return ActiveDocumentInfo {
                    is_open: false,
                    name: None,
                    path: None,
                };
            }
            let looks_like_path = rest
                .chars()
                .nth(1)
                .is_some_and(|c| c == ':')
                && rest
                    .chars()
                    .next()
                    .is_some_and(|c| c.is_ascii_alphabetic())
                || rest.starts_with(r"\\");
            if looks_like_path {
                let pb = Path::new(rest);
                let name = pb
                    .file_name()
                    .map(|s| s.to_string_lossy().into_owned());
                return ActiveDocumentInfo {
                    is_open: true,
                    name,
                    path: Some(rest.to_string()),
                };
            }
            return ActiveDocumentInfo {
                is_open: true,
                name: Some(rest.to_string()),
                path: None,
            };
        }
    }

    if !title.is_empty() {
        ActiveDocumentInfo {
            is_open: true,
            name: Some(title.to_owned()),
            path: None,
        }
    } else {
        ActiveDocumentInfo {
            is_open: false,
            name: None,
            path: None,
        }
    }
}

unsafe fn describe_affinity_window(hwnd: HWND) -> Option<ActiveDocumentInfo> {
    if !IsWindowVisible(hwnd).as_bool() {
        return None;
    }

    let mut pid = 0u32;
    GetWindowThreadProcessId(hwnd, Some(&mut pid));
    let exe_is_affinity = query_exe_path(pid)
        .map(|p| is_affinity_exe_path(&p))
        .unwrap_or(false);

    let mut buf = [0u16; 512];
    let n = GetWindowTextW(hwnd, &mut buf);
    let title_opt: Option<String> = if n > 0 {
        let mut s = String::from_utf16_lossy(&buf[..n as usize]);
        if s.ends_with('\0') {
            s.pop();
        }
        let t = s.trim();
        if t.is_empty() {
            None
        } else {
            Some(t.to_string())
        }
    } else {
        None
    };

    let title_hint = title_opt
        .as_deref()
        .map(title_suggests_affinity_window)
        .unwrap_or(false);

    // exe が取れない場合はタイトルパターンのみ。exe が Affinity ならキャプションが空でもウィンドウを採用（WinUI 等で GetWindowText が空のことがある）。
    if !exe_is_affinity {
        if !title_hint {
            return None;
        }
    }

    match title_opt.as_ref() {
        // キャプションが空でも exe が Affinity ならメインウィンドウありとみなす
        None => Some(ActiveDocumentInfo {
            is_open: true,
            name: None,
            path: None,
        }),
        Some(s) => {
            let mut info = parse_affinity_caption(s);
            // Affinity 3 ドック時はタイトルが常に "Affinity" などのみでドキュメント名は出ない。
            // exe で本体と分かったうえで「アイドル」扱いのタイトルなら、セッションは開いているとみなす。
            if exe_is_affinity && !info.is_open && is_idle_app_title(s) {
                info = ActiveDocumentInfo {
                    is_open: true,
                    name: None,
                    path: None,
                };
            }
            Some(info)
        }
    }
}

fn affinity_mcp_debug_enabled() -> bool {
    matches!(
        std::env::var("AFFINITY_MCP_DEBUG").as_deref(),
        Ok("1") | Ok("true") | Ok("TRUE")
    )
}

struct DebugDumpCtx {
    lines: Vec<String>,
}

unsafe extern "system" fn enum_debug_dump(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let ctx = &mut *(lparam.0 as *mut DebugDumpCtx);
    if ctx.lines.len() >= 40 {
        return TRUE;
    }
    if !IsWindowVisible(hwnd).as_bool() {
        return TRUE;
    }
    let mut buf = [0u16; 512];
    let n = GetWindowTextW(hwnd, &mut buf);
    let title = if n > 0 {
        String::from_utf16_lossy(&buf[..n as usize])
    } else {
        String::new()
    };
    let title = title.trim();
    let mut cls = [0u16; 256];
    let cn = GetClassNameW(hwnd, &mut cls);
    let class = if cn > 0 {
        String::from_utf16_lossy(&cls[..cn as usize])
    } else {
        String::new()
    };
    let mut pid = 0u32;
    GetWindowThreadProcessId(hwnd, Some(&mut pid));
    let exe = query_exe_path(pid).unwrap_or_else(|| {
        "<OpenProcess or QueryFullProcessImageName failed>".to_string()
    });
    let exe_ok = is_affinity_exe_path(&exe);
    let title_ok = title_suggests_affinity_window(title);
    let hint = exe_ok
        || title_ok
        || title.to_lowercase().contains("affinity")
        || class.to_lowercase().contains("affinity")
        || exe.to_lowercase().contains("affinity")
        || exe.to_lowercase().contains("canva");
    if hint {
        ctx.lines.push(format!(
            "  hwnd={:?} pid={} title={:?} class={:?} exe_ok={} title_ok={} exe={}",
            hwnd, pid, title, class, exe_ok, title_ok, exe
        ));
    }
    TRUE
}

fn maybe_debug_affinity(fg: HWND, fg_affinity_snap: &Option<ActiveDocumentInfo>, info: &ActiveDocumentInfo) {
    if affinity_mcp_debug_enabled() && !info.is_open {
        unsafe {
            debug_dump_window_candidates(fg, fg_affinity_snap);
        }
    }
}

/// `is_open: false` のとき原因調査用。stderr に候補ウィンドウを出す（stdout は JSON-RPC 専用）。
unsafe fn debug_dump_window_candidates(fg: HWND, fg_affinity: &Option<ActiveDocumentInfo>) {
    eprintln!(
        "[affinity-mcp] AFFINITY_MCP_DEBUG: foreground={:?} describe_affinity={:?}",
        fg, fg_affinity
    );
    eprintln!("[affinity-mcp] visible top-level windows whose title/class/exe hints at Affinity:");

    let mut ctx = DebugDumpCtx {
        lines: Vec::new(),
    };
    let ctx_ptr: *mut DebugDumpCtx = &mut ctx;
    let _ = EnumWindows(Some(enum_debug_dump), LPARAM(ctx_ptr as isize));
    if ctx.lines.is_empty() {
        eprintln!("  (none — Affinity may use empty captions / windows not visible to this API)");
    } else {
        for line in ctx.lines {
            eprintln!("{}", line);
        }
    }
}

struct EnumCtx {
    foreground: HWND,
    matches: Vec<(HWND, ActiveDocumentInfo)>,
}

unsafe extern "system" fn enum_windows_proc(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let ctx = &mut *(lparam.0 as *mut EnumCtx);
    if let Some(info) = describe_affinity_window(hwnd) {
        ctx.matches.push((hwnd, info));
    }
    TRUE
}

/// 前面ウィンドウを優先し、次に Affinity のトップレベルウィンドウを列挙する。
/// Affinity 3 では前面がドックの `Affinity` だけのとき、アンドック済みドキュメントがあるならそちらを優先する。
pub fn get_active_document_blocking() -> Result<ActiveDocumentInfo> {
    unsafe {
        let fg = GetForegroundWindow();
        let fg_affinity = if !fg.is_invalid() {
            describe_affinity_window(fg)
        } else {
            None
        };
        let fg_affinity_snap = fg_affinity.clone();
        // 前面がドキュメント付きなら即返す（タイトルに名前がある場合のみ is_open）。
        if let Some(ref info) = fg_affinity {
            if info.is_open {
                return Ok(info.clone());
            }
        }

        let mut ctx = EnumCtx {
            foreground: fg,
            matches: Vec::new(),
        };
        let ctx_ptr: *mut EnumCtx = &mut ctx;
        EnumWindows(Some(enum_windows_proc), LPARAM(ctx_ptr as isize))
            .map_err(|e| anyhow::anyhow!("EnumWindows: {}", e))?;

        if !ctx.matches.is_empty() {
            // ドック時の `Affinity` とアンドック時のドキュメント名が両方ある場合はドキュメント側を優先する。
            ctx.matches
                .sort_by(|a, b| b.1.is_open.cmp(&a.1.is_open));

            if !ctx.foreground.is_invalid() {
                let fg_open = ctx.matches.iter().find(|(h, info)| {
                    *h == ctx.foreground && info.is_open
                });
                if let Some((_, info)) = fg_open {
                    return Ok(info.clone());
                }
            }

            let info = ctx.matches[0].1.clone();
            maybe_debug_affinity(fg, &fg_affinity_snap, &info);
            return Ok(info);
        }

        if let Some(info) = fg_affinity {
            maybe_debug_affinity(fg, &fg_affinity_snap, &info);
            return Ok(info);
        }

        let empty = ActiveDocumentInfo {
            is_open: false,
            name: None,
            path: None,
        };
        maybe_debug_affinity(fg, &fg_affinity_snap, &empty);
        Ok(empty)
    }
}
