# Claudeの始め方 ― Windows デスクトップ版を中心に

## はじめに

Anthropicが提供するAIアシスタント「Claude」は、2026年現在、単なるチャットボットから**本格的な作業パートナー**へと進化している。Windowsデスクトップアプリには「Chat」「Cowork」「Code」の3つのモードが搭載され、会話・知的作業の自動化・ソフトウェア開発をひとつのアプリで完結できる。

本記事では、Windowsユーザーがゼロから環境を構築し、Claude Code・Cowork・スキル・プラグインを活用するまでの道筋を解説する。

---

## 1. まずはアカウント作成とプラン選択

### 1-1. アカウント登録

[claude.com](https://claude.com) にアクセスし、メールアドレスまたはGoogleアカウントで登録する。無料プランでもChatモードは利用可能。

### 1-2. プラン比較

| プラン | 月額 | Cowork | Claude Code | 主な用途 |
|--------|------|--------|-------------|----------|
| Free | $0 | × | × | 基本的なチャット |
| Pro | $20 | ○ | ○ | 個人の日常利用 |
| Max | $100〜200 | ○ | ○ | ヘビーユーザー・長時間セッション |
| Team | $30/人 | ○ | ○ | チーム利用 |
| Enterprise | 要問合せ | ○ | ○ | 組織全体導入 |

**Cowork・Claude Codeを使うならPro以上が必須。** 本格的に運用するならMaxプランが安定的に使える。

### 1-3. デスクトップアプリのインストール

[claude.com/download](https://claude.com/download) からWindows版インストーラーをダウンロードし、実行する。

> **注意:** Coworkを使う場合、Hyper-V（仮想化機能）が必要。Windows 11 Homeでは制限がある場合があるため、後述の「Coworkの注意点」を確認すること。

---

## 2. 3つのモードを理解する

デスクトップアプリ上部に3つのタブがある。

### Chat（チャット）
従来通りのClaudeとの対話。質問・要約・翻訳・アイデア出しなど。最も安定しており、まずはここから使い始めるのが良い。

### Cowork（コワーク）
Claude Codeのエージェント的な能力を、**コーディング以外の知的作業**にも開放したモード。ローカルファイルの読み書き・編集、複数ステップのタスク実行、ブラウザ連携などが可能。

### Code（コード）
ターミナルベースのClaude Codeをデスクトップから利用するモード。コードベースの理解・編集・Git操作・テスト実行などをエージェント的に行う。

---

## 3. Claude Code ― 開発者向けの中核機能

### 3-1. Claude Codeとは

Claude Codeは、ターミナル上で動作するエージェント型のコーディングツール。自然言語の指示でコードの読解・生成・修正・テスト実行・Gitワークフローまでを自動化できる。

### 3-2. インストール（ターミナル版）

```bash
# Node.js (v18以上) が必要
npm install -g @anthropic-ai/claude-code

# 起動
claude
```

デスクトップアプリの「Code」タブからも利用できるが、ターミナルから直接使う方がフル機能を発揮できる。

### 3-3. 基本的なワークフロー

```
# プロジェクトディレクトリで起動
cd your-project
claude

# 自然言語で指示
> このプロジェクトの構造を説明して
> src/auth.tsのログイン処理にバリデーションを追加して
> テストを実行して失敗したものを修正して
```

Claude Codeは**コードベース全体を把握**した上で作業するため、ファイル間の依存関係も考慮した修正が可能。

### 3-4. 主なコマンド

| コマンド | 機能 |
|----------|------|
| `/help` | ヘルプ表示 |
| `/commit` | 変更をGitコミット |
| `/install-github-app` | GitHub連携のセットアップ |
| `/review` | コードレビュー |
| `/clear` | コンテキストのクリア |

---

## 4. Cowork ― 知的作業の自動化

### 4-1. Coworkでできること

- **ドキュメント作成・編集：** レポート、提案書、議事録の自動生成
- **ファイル整理：** フォルダ構造の再編成、ファイルのリネーム
- **データ分析：** CSVやExcelファイルの読み込み・集計・グラフ化
- **リサーチ：** 複数ファイルの横断的な分析・比較
- **スケジュールタスク：** 定期的・オンデマンドのタスクを予約実行

### 4-2. Coworkの使い方

1. デスクトップアプリで「Cowork」タブを選択
2. Claudeにアクセスを許可するフォルダを指定
3. 自然言語でタスクを指示

```
例：「Documentsフォルダ内の2025年のレポートをすべて読み、
     売上トレンドの要約を新しいMarkdownファイルにまとめて」
```

Claudeはサブエージェントを複数起動して並列に作業する場合もある。

### 4-3. Windows版Coworkの現実 ― 正直な評価

**結論から言えば、2026年3月時点でWindows版Coworkは安定性に大きな課題がある。**

Mac版と比較して、Windows版では以下のような問題が多数報告されている：

#### 主な問題点

1. **VM（仮想マシン）の不安定さ**
   - CoworkはHyper-V上のVMで動作するが、クラッシュ後にVMプロセスが残留し、再起動しないと復旧できないケースがある
   - VPNやネットワーク構成によっては、VMの内部ネットワーク（172.16.0.0/24）と競合してインターネット接続が失敗する

2. **Windows 11 Homeでの制限**
   - API接続の失敗が頻繁に報告されている
   - Hyper-Vの制限に起因する問題がある可能性

3. **ワークスペースの破損**
   - 特定のタスク失敗後にワークスペースが恒久的に壊れ、すべてのタスクが失敗するケースがある
   - 「キャッシュクリアと再起動」では復旧しない場合も

4. **MCP拡張機能の競合**
   - アプリの自動更新後に、手動設定したMCPサーバーが無効になる問題

5. **インストール時のDeveloper Mode要求**
   - 企業環境ではセキュリティポリシーと衝突する可能性

#### 対処法

- **Windows 11 ProまたはEnterprise**を推奨（Hyper-Vのフルサポート）
- VPN使用時はCowork起動前にVPNを切断してテスト
- 問題発生時は`Get-NetNat`コマンドでNATルールの存在を確認
- 深刻な問題は[GitHub Issues](https://github.com/anthropics/claude-code/issues)で報告・確認

#### 筆者の所感

Mac版では比較的安定して動作するCoworkが、Windows版ではエラーが多く実用には忍耐が必要というのが現状。Anthropicはパッチを随時リリースしているが、完全な安定にはまだ時間がかかるだろう。**Windowsユーザーは、ChatモードとClaude Code（ターミナル版）を主軸に据え、Coworkは補助的に使う**のが現実的な運用方針。

---

## 5. スキル（Skills）― Claudeに手順を教える

### 5-1. スキルとは

スキルは、Claudeに「特定のタスクのやり方」を教えるための仕組み。`.md`（Markdown）ファイルを所定のディレクトリに置くだけで、Claudeが自動的に検出して必要な場面で適用する。

**APIもSDKもビルドステップも不要。** プレーンテキストのMarkdownを書くだけ。

### 5-2. スキルの配置場所

```
~/.claude/skills/       ← グローバルスキル（全プロジェクト共通）
.claude/skills/         ← プロジェクト固有スキル
```

### 5-3. スキルの作成例

ファイル: `~/.claude/skills/git-commit-style.md`

```markdown
---
name: Git Commit Style
description: チーム標準のコミットメッセージ規約
---

# コミットメッセージ規約

コミットメッセージは以下の形式に従うこと：

## フォーマット
<type>(<scope>): <subject>

## Type一覧
- feat: 新機能
- fix: バグ修正
- docs: ドキュメントのみの変更
- refactor: リファクタリング
- test: テストの追加・修正

## 例
feat(auth): ログイン時の2要素認証を追加
fix(api): レスポンスのタイムアウト値を修正
```

このスキルを配置すると、Claude Codeがコミット時に自動的にこのスタイルに従う。

### 5-4. スキルの効率性

スキルは**遅延読み込み（Progressive Disclosure）**で動作する：

1. **メタデータ読み込み**（約30〜50トークン）：利用可能なスキルを走査
2. **全文読み込み**（最大5,000トークン）：該当スキルが適用される場合のみ
3. **リソース読み込み**：実行コード等は必要時のみ

つまり、大量のスキルを登録してもコンテキストを圧迫しない。

### 5-5. コミュニティスキル

[awesome-claude-skills](https://github.com/travisvn/awesome-claude-skills) リポジトリに、コミュニティが作成した便利なスキル集がまとめられている。

---

## 6. プラグイン（Plugins）― スキル・フック・MCPをまとめる

### 6-1. プラグインとは

プラグインは、スキル・フック・MCP設定を**ひとつのパッケージ**にまとめたもの。配布・再利用が容易な形式。

### 6-2. プラグインの構造

```
my-plugin/
├── plugin.json          ← マニフェスト（名前・バージョン・構成定義）
├── skills/
│   └── my-skill.md      ← スキルファイル
├── hooks/
│   └── pre-commit.sh    ← フックスクリプト
└── mcp/
    └── config.json      ← MCP設定
```

### 6-3. plugin.jsonの例

```json
{
  "name": "my-dev-toolkit",
  "version": "1.0.0",
  "author": "your-name",
  "description": "開発ワークフロー用プラグイン",
  "skills": ["skills/my-skill.md"],
  "hooks": {
    "pre-commit": "hooks/pre-commit.sh"
  },
  "mcp": {
    "servers": ["mcp/config.json"]
  }
}
```

### 6-4. プラグインのテスト

```bash
claude --plugin-dir ./my-plugin
```

ローカルのプラグインディレクトリを指定して動作確認できる。

---

## 7. MCP（Model Context Protocol）― 外部サービス連携

### 7-1. MCPとは

MCPは、ClaudeをGitHub・データベース・Slack・Google Calendarなどの外部ツールと接続するためのプロトコル。Claudeに「手足」を与えるような仕組み。

### 7-2. MCP設定例（GitHub連携）

ファイル: `~/.claude/mcp_servers.json`

```json
{
  "github": {
    "command": "npx",
    "args": ["-y", "@modelcontextprotocol/server-github"],
    "env": {
      "GITHUB_PERSONAL_ACCESS_TOKEN": "ghp_xxxxxxxxxxxx"
    }
  }
}
```

### 7-3. デスクトップアプリでの設定

デスクトップアプリの「Customize」セクションから、GUIでコネクタ（MCP）の追加・管理ができる。Chrome拡張機能「Claude in Chrome」もここから有効化する。

### 7-4. トークンコストに注意

MCPサーバーは起動時にツール定義をコンテキストに読み込むため、5つのサーバーで約55,000トークンを消費する場合がある。Claude Codeの**MCP Tool Search**機能を使えば、遅延読み込みにより最大95%のコンテキスト削減が可能。

### 7-5. スキルとMCPの使い分け

| 項目 | スキル | MCP |
|------|--------|-----|
| 目的 | 手順・ワークフローの知識 | 外部ツールとの接続 |
| コンテキストコスト | 30〜50トークン | 数万トークン |
| 適した場面 | 繰り返し作業のパターン化 | リアルタイムのデータ取得 |

両者は補完関係にあり、**スキル内でMCPツールを参照する**こともできる。

---

## 8. GitHubは登録すべきか？

### 結論：登録を強く推奨する

GitHubの登録は無料で、Claudeの活用幅が大きく広がる。

### 登録するメリット

1. **Claude Code + GitHub連携**
   - PRの自動作成・レビュー・マージ
   - Issue起票→自動修正→PR作成の一連の流れを自動化
   - `@claude`メンションでIssueやPRにClaudeを呼び出せる

2. **コミュニティリソースへのアクセス**
   - [awesome-claude-skills](https://github.com/travisvn/awesome-claude-skills)などのスキル集
   - [claude-code-plugins-plus-skills](https://github.com/jeremylongshore/claude-code-plugins-plus-skills)（270以上のプラグイン・739のスキル）
   - MCP サーバーのソースコードと設定例

3. **バグ報告・情報収集**
   - [anthropics/claude-code](https://github.com/anthropics/claude-code)のIssuesで最新の不具合情報と回避策を確認可能
   - Windows版Coworkの問題に遭遇した際、同様の報告を検索できる

4. **GitHub Copilot経由でのClaude利用**
   - GitHub Copilot Pro+またはEnterpriseプランなら、追加料金なしでClaude Codeをコーディングエージェントとして利用可能

5. **バージョン管理の基本**
   - 自分のコードやドキュメントのバージョン管理
   - Claude Codeとの共同作業のベースとなる

### 登録不要なケース

- Claudeをチャット（会話）目的のみで使い、コーディングを一切しない場合
- Coworkでのファイル操作のみで完結する場合

**ただし、GitHubアカウントは無料で作れるため、将来的な拡張性を考えると登録しておいて損はない。**

---

## 9. 推奨する段階的な導入ステップ

### ステップ1：基盤を固める（最初の1週間）
- [ ] Claudeアカウント作成（Pro以上推奨）
- [ ] デスクトップアプリのインストール
- [ ] ChatモードでClaudeの対話に慣れる
- [ ] GitHubアカウント作成（無料）

### ステップ2：Claude Codeを始める（2〜3週目）
- [ ] Node.jsのインストール
- [ ] Claude Code（ターミナル版）のインストール
- [ ] 小さなプロジェクトでClaude Codeを試す
- [ ] `/commit`、`/help`などの基本コマンドを覚える

### ステップ3：拡張機能を追加する（4週目〜）
- [ ] 最初のスキルを作成（コミットメッセージ規約など）
- [ ] MCPサーバーをひとつ追加（GitHub連携推奨）
- [ ] デスクトップアプリの「Customize」でコネクタを探索

### ステップ4：Coworkを試す（安定性を見ながら）
- [ ] 簡単なファイル整理タスクから開始
- [ ] エラーが発生したら[GitHub Issues](https://github.com/anthropics/claude-code/issues)を確認
- [ ] Mac環境があればそちらでCoworkを優先利用

---

## 10. トラブルシューティング

### Claude Codeが起動しない
```bash
node -v   # v18以上であること
npm install -g @anthropic-ai/claude-code  # 再インストール
```

### Coworkが接続エラーを出す（Windows）
1. Hyper-Vが有効か確認（設定 → アプリ → オプション機能 → Windows の機能）
2. VPNを一時的に無効化して再試行
3. PowerShellで`Get-NetNat`を実行し、NATルールの存在を確認
4. アプリの再起動、またはPC自体の再起動

### MCPサーバーが動作しない
- デスクトップアプリの自動更新後に設定が競合する場合がある
- `~/.claude/mcp_servers.json`の設定をバックアップしておくこと

---

## おわりに

Claudeは「チャットAI」の域を超え、**ファイル操作・コーディング・外部サービス連携を自律的に行えるエージェント**へと進化している。

Windowsユーザーにとっての現実的な優先順位は：

1. **Chat** ― 今すぐ使える、安定している
2. **Claude Code（ターミナル）** ― 開発者なら必須、安定性も高い
3. **スキル・プラグイン・MCP** ― 慣れてきたら段階的に導入
4. **Cowork** ― 将来的に期待できるが、Windows版は安定を待ってから本格運用

Mac版Coworkが安定して動作する一方、Windows版はまだ発展途上というのが偽りのない現状評価だ。しかし、Anthropicのアップデート頻度は高く、状況は改善に向かっている。まずはChatとClaude Codeで基盤を固め、Coworkの安定化を待ちながら段階的に活用範囲を広げていくのが賢明だろう。

---

### 参考リンク

- [Claude ダウンロード](https://claude.com/download)
- [Cowork 公式紹介](https://claude.com/product/cowork)
- [Cowork ヘルプセンター](https://support.claude.com/en/articles/13345190-get-started-with-cowork)
- [Claude Code ドキュメント](https://code.claude.com/docs/en/desktop)
- [Claude Code プラグイン作成ガイド](https://code.claude.com/docs/en/plugins)
- [awesome-claude-skills（GitHub）](https://github.com/travisvn/awesome-claude-skills)
- [claude-code-plugins-plus-skills（GitHub）](https://github.com/jeremylongshore/claude-code-plugins-plus-skills)
- [Claude Code Issues（バグ報告）](https://github.com/anthropics/claude-code/issues)
- [GitHub Integration ヘルプ](https://support.claude.com/en/articles/10167454-using-the-github-integration)

---

*最終更新: 2026年3月*
