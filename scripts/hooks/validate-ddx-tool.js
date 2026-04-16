#!/usr/bin/env node
/**
 * PostToolUse Hook: 医学診断ツールHTML バリデーション
 *
 * medical-ddx-tools/ ディレクトリ内の .html ファイルが編集された際に自動実行。
 * CLAUDE.mdのコーディング規約への準拠を検証する。
 *
 * 検出項目:
 * 1. const/let の使用（var を使うべき — iOS Safari互換性）
 * 2. アロー関数 => の使用（function(){} を使うべき）
 * 3. テンプレートリテラル ` の使用（文字列連結を使うべき）
 * 4. SW更新スクリプトの存在チェック
 * 5. manifest.json リンクの存在チェック
 */

'use strict';

var fs = require('fs');
var path = require('path');

var MAX_STDIN = 1024 * 1024;

function log(msg) {
  process.stderr.write(msg + '\n');
}

function isDdxTool(filePath) {
  if (!filePath) return false;
  var normalized = filePath.replace(/\\/g, '/');
  return normalized.indexOf('/medical-ddx-tools/') !== -1
    && normalized.endsWith('.html')
    && normalized.indexOf('index.html') === -1;
}

function findLineNumbers(content, pattern) {
  var lines = content.split('\n');
  var matches = [];
  for (var i = 0; i < lines.length; i++) {
    if (pattern.test(lines[i])) {
      matches.push({ line: i + 1, text: lines[i].trim().substring(0, 80) });
    }
  }
  return matches;
}

function extractScriptContent(htmlContent) {
  // <script>タグ内のJSだけを抽出（CSS内のletやconstを誤検知しない）
  var scripts = [];
  var regex = /<script[^>]*>([\s\S]*?)<\/script>/gi;
  var match;
  while ((match = regex.exec(htmlContent)) !== null) {
    scripts.push(match[1]);
  }
  return scripts.join('\n');
}

function validate(rawInput) {
  try {
    var input = JSON.parse(rawInput);
    var filePath = String(input.tool_input && input.tool_input.file_path || '');

    if (!isDdxTool(filePath)) {
      return rawInput;
    }

    if (!fs.existsSync(filePath)) {
      return rawInput;
    }

    var content = fs.readFileSync(filePath, 'utf8');
    var jsContent = extractScriptContent(content);
    var warnings = 0;

    log('');
    log('[DdxValidator] ========================================');
    log('[DdxValidator] 診断ツールHTML検証: ' + path.basename(filePath));
    log('[DdxValidator] ========================================');

    // const/let チェック（JS部分のみ）
    var jsLines = jsContent.split('\n');
    var constLetLines = [];
    for (var i = 0; i < jsLines.length; i++) {
      // コメント行はスキップ
      var trimmed = jsLines[i].trim();
      if (trimmed.indexOf('//') === 0) continue;
      if (/\b(const|let)\s/.test(jsLines[i])) {
        constLetLines.push((i + 1) + ': ' + trimmed.substring(0, 80));
      }
    }
    if (constLetLines.length > 0) {
      log('[DdxValidator] WARNING [ES6_CONST_LET] const/letが検出されました。iOS Safari互換性のためvarを使用してください。');
      constLetLines.slice(0, 5).forEach(function(l) { log('  ' + l); });
      if (constLetLines.length > 5) {
        log('  ...他 ' + (constLetLines.length - 5) + ' 箇所');
      }
      warnings++;
    }

    // アロー関数チェック（JS部分のみ）
    var arrowLines = [];
    for (var j = 0; j < jsLines.length; j++) {
      var trimmedJ = jsLines[j].trim();
      if (trimmedJ.indexOf('//') === 0) continue;
      if (/=>/.test(jsLines[j])) {
        arrowLines.push((j + 1) + ': ' + trimmedJ.substring(0, 80));
      }
    }
    if (arrowLines.length > 0) {
      log('[DdxValidator] WARNING [ES6_ARROW] アロー関数(=>)が検出されました。function(){}を使用してください。');
      arrowLines.slice(0, 5).forEach(function(l) { log('  ' + l); });
      if (arrowLines.length > 5) {
        log('  ...他 ' + (arrowLines.length - 5) + ' 箇所');
      }
      warnings++;
    }

    // テンプレートリテラルチェック（JS部分のみ、CSSの`は除外）
    var templateLines = [];
    for (var k = 0; k < jsLines.length; k++) {
      var trimmedK = jsLines[k].trim();
      if (trimmedK.indexOf('//') === 0) continue;
      if (/`/.test(jsLines[k])) {
        templateLines.push((k + 1) + ': ' + trimmedK.substring(0, 80));
      }
    }
    if (templateLines.length > 0) {
      log('[DdxValidator] WARNING [ES6_TEMPLATE] テンプレートリテラル(`)が検出されました。文字列連結を使用してください。');
      templateLines.slice(0, 3).forEach(function(l) { log('  ' + l); });
      if (templateLines.length > 3) {
        log('  ...他 ' + (templateLines.length - 3) + ' 箇所');
      }
      warnings++;
    }

    // SW更新スクリプトチェック
    if (content.indexOf('serviceWorker') === -1) {
      log('[DdxValidator] WARNING [NO_SW] Service Worker更新スクリプトが見つかりません。<meta charset>直後に配置してください。');
      warnings++;
    }

    // manifest.jsonリンクチェック
    if (content.indexOf('manifest.json') === -1) {
      log('[DdxValidator] WARNING [NO_MANIFEST] manifest.jsonへのリンクが見つかりません。');
      warnings++;
    }

    if (warnings === 0) {
      log('[DdxValidator] OK 全チェック通過');
    } else {
      log('[DdxValidator] ----------------------------------------');
      log('[DdxValidator] 結果: ' + warnings + ' 警告');
    }
    log('[DdxValidator] ========================================');
    log('');

  } catch (e) {
    // パースエラーはスキップ
  }

  return rawInput;
}

if (require.main === module) {
  var raw = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', function(chunk) {
    if (raw.length < MAX_STDIN) {
      var remaining = MAX_STDIN - raw.length;
      raw += chunk.substring(0, remaining);
    }
  });

  process.stdin.on('end', function() {
    var result = validate(raw);
    process.stdout.write(result);
  });
}

module.exports = { run: validate };
