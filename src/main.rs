/**
 * AffinityMCP メインエントリーポイント
 * 
 * 概要:
 *   RustベースのMCPサーバー。STDIO経由でJSON-RPC通信を行い、
 *   Canva連携ツールとAffinityブリッジを提供する。
 * 
 * 主な仕様:
 *   - STDIO経由でJSON-RPCリクエスト/レスポンスを処理
 *   - 環境変数 MCP_NAME でサーバー名を設定可能（デフォルト: affinity-mcp）
 *   - stderr にログを出力（tracing-subscriber）
 *   - MCPプロトコル（initialize、tools/list、tools/call）を実装
 * 
 * エラー処理:
 *   - 詳細なエラーメッセージを出力
 *   - 関数名、引数、パラメータを含む
 */
use std::env;
use tracing::Level;
use anyhow::Context;
use std::io::IsTerminal;

use futures::StreamExt;
use jsonrpc_core::IoHandler;
use tokio::io::{self, AsyncWriteExt};
use tokio_util::codec::{FramedRead, LinesCodec};

mod mcp;
mod tools;

/// `jsonrpc-stdio-server` クレートは応答なしの通知でも stdout に改行だけ送り、
/// Cursor が空行を JSON としてパースして "Unexpected end of JSON input" になる。
/// 応答が空のときは何も書かない。
async fn run_stdio_jsonrpc(io: IoHandler) {
    let mut stdin = FramedRead::new(io::stdin(), LinesCodec::new());
    let mut stdout = io::stdout();

    while let Some(line) = stdin.next().await {
        match line {
            Ok(line) => {
                if line.trim().is_empty() {
                    continue;
                }
                let response = io.handle_request(&line).await;
                if let Some(s) = response {
                    if s.is_empty() {
                        continue;
                    }
                    let mut sanitized = s.replace('\n', "");
                    sanitized.push('\n');
                    if let Err(e) = stdout.write_all(sanitized.as_bytes()).await {
                        tracing::warn!(error = ?e, "stdout write");
                    }
                }
            }
            Err(e) => tracing::warn!(error = ?e, "stdin read"),
        }
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let name = env::var("MCP_NAME").unwrap_or_else(|_| "affinity-mcp".into());

    // stderr ログ（ANSIカラーコードを無効化、環境変数で制御可能）
    let log_level = env::var("RUST_LOG")
        .unwrap_or_else(|_| "WARN".to_string())
        .parse::<Level>()
        .unwrap_or(Level::WARN);
    
    let use_ansi = env::var("TERM").is_ok() && 
                   env::var("NO_COLOR").is_err() &&
                   std::io::stderr().is_terminal();
    
    tracing_subscriber::fmt()
        .with_writer(std::io::stderr)
        .with_max_level(log_level)
        .with_ansi(use_ansi)
        .with_target(false)
        .compact()
        .init();

    tracing::debug!(server = %name, "Starting AffinityMCP server (STDIO).");

    // ツール初期化
    tools::register_all().await?;

    // MCPサーバー構築
    let io = mcp::build_server(name.clone())
        .context("MCPサーバーの構築に失敗しました")?;

    // STDIOサーバー起動
    tracing::debug!(server = %name, "MCP server ready. Listening for JSON-RPC requests on STDIO.");
    
    run_stdio_jsonrpc(io).await;

    tracing::debug!("MCP server shutting down.");
    Ok(())
}

