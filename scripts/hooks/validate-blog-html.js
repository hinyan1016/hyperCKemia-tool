#!/usr/bin/env node
/**
 * PostToolUse Hook: はてなブログHTML バリデーション
 *
 * blog/ ディレクトリ内の *_blog.html ファイルが編集された際に自動実行。
 * はてなブログで問題を起こす要素と、SEO構成順序を検証する。
 *
 * 検出項目:
 * 1. <style> タグ（テーマCSSと競合→空白発生）
 * 2. <iframe> タグ（「見たまま」モードで除去される）
 * 3. <th> タグ（テーマCSSが強制上書き）
 * 4. <div class="..."> パターン（正規化で分割・破壊）
 * 5. SEO構成順序の検証
 */

'use strict';

var fs = require('fs');
var path = require('path');

var MAX_STDIN = 1024 * 1024;

// はてなブログ禁止要素の定義
var FORBIDDEN_PATTERNS = [
  {
    pattern: /<style[\s>]/i,
    id: 'STYLE_TAG',
    severity: 'ERROR',
    message: '<style>タグが検出されました。はてなブログではテーマCSSと競合し空白が発生します。インラインスタイルを使用してください。'
  },
  {
    pattern: /<iframe[\s>]/i,
    id: 'IFRAME_TAG',
    severity: 'ERROR',
    message: '<iframe>タグが検出されました。「見たまま」モードの正規化で除去されます。YouTube動画はURLをそのまま貼り付けてください。'
  },
  {
    pattern: /<th[\s>]/i,
    id: 'TH_TAG',
    severity: 'WARNING',
    message: '<th>タグが検出されました。はてなブログのテーマCSSが背景色を強制上書きします。<td>にインラインスタイルで代用してください。'
  },
  {
    pattern: /<div\s+class="/i,
    id: 'DIV_CLASS',
    severity: 'WARNING',
    message: '<div class="...">が検出されました。「見たまま」モードの正規化で入れ子構造が破壊されます。シンプルな<p><a>タグを使用してください。'
  }
];

// SEO構成順序の検証（期待される順序）
var SEO_STRUCTURE_MARKERS = [
  { id: 'intro', pattern: /^<p>/, label: '導入文' },
  { id: 'toc', pattern: /<h2[^>]*>.*目次/i, label: '目次' },
  { id: 'content', pattern: /<h2[\s>]/i, label: '本文H2セクション' },
  { id: 'faq', pattern: /FAQ|よくある質問/i, label: 'FAQ' },
  { id: 'references', pattern: /参考文献|References/i, label: '参考文献' }
];

function log(msg) {
  process.stderr.write(msg + '\n');
}

function isBlogHtml(filePath) {
  if (!filePath) return false;
  var normalized = filePath.replace(/\\/g, '/');
  return normalized.indexOf('/blog/') !== -1 && normalized.endsWith('_blog.html');
}

function findLineNumbers(content, pattern) {
  var lines = content.split('\n');
  var matches = [];
  for (var i = 0; i < lines.length; i++) {
    if (pattern.test(lines[i])) {
      matches.push(i + 1);
    }
  }
  return matches;
}

function validateForbiddenElements(content, filePath) {
  var errors = 0;
  var warnings = 0;

  FORBIDDEN_PATTERNS.forEach(function(rule) {
    var lineNums = findLineNumbers(content, rule.pattern);
    if (lineNums.length > 0) {
      var prefix = rule.severity === 'ERROR' ? 'ERROR' : 'WARNING';
      log('[BlogValidator] ' + prefix + ' [' + rule.id + '] ' + rule.message);
      log('  ファイル: ' + filePath);
      log('  該当行: ' + lineNums.join(', '));
      if (rule.severity === 'ERROR') {
        errors++;
      } else {
        warnings++;
      }
    }
  });

  return { errors: errors, warnings: warnings };
}

function validateSeoStructure(content, filePath) {
  var warnings = 0;

  // 導入文がファイル先頭（<body>直後または最初の<p>）にあるか
  var firstParagraph = content.indexOf('<p>');
  var firstH2 = content.indexOf('<h2');

  if (firstParagraph === -1) {
    log('[BlogValidator] WARNING [SEO_NO_INTRO] 導入文（<p>タグ）が見つかりません。キーワードリッチな導入文を最初に配置してください。');
    log('  ファイル: ' + filePath);
    warnings++;
  } else if (firstH2 !== -1 && firstH2 < firstParagraph) {
    log('[BlogValidator] WARNING [SEO_INTRO_ORDER] H2見出しが導入文より前に出現しています。導入文を最初に配置してください。');
    log('  ファイル: ' + filePath);
    warnings++;
  }

  // FAQが参考文献より前にあるか
  var faqPos = content.search(/FAQ|よくある質問/i);
  var refPos = content.search(/参考文献|References/i);

  if (faqPos !== -1 && refPos !== -1 && faqPos > refPos) {
    log('[BlogValidator] WARNING [SEO_FAQ_ORDER] FAQセクションが参考文献の後に配置されています。参考文献はFAQの後に配置してください。');
    log('  ファイル: ' + filePath);
    warnings++;
  }

  // 参考文献にDOIリンクがあるか
  if (refPos !== -1) {
    var refSection = content.substring(refPos);
    if (refSection.indexOf('doi.org') === -1 && refSection.indexOf('DOI') === -1) {
      log('[BlogValidator] WARNING [SEO_NO_DOI] 参考文献セクションにDOIリンクが見つかりません。');
      log('  ファイル: ' + filePath);
      warnings++;
    }
  }

  return { warnings: warnings };
}

function validate(rawInput) {
  try {
    var input = JSON.parse(rawInput);
    var filePath = String(input.tool_input && input.tool_input.file_path || '');

    if (!isBlogHtml(filePath)) {
      return rawInput;
    }

    if (!fs.existsSync(filePath)) {
      return rawInput;
    }

    var content = fs.readFileSync(filePath, 'utf8');

    log('');
    log('[BlogValidator] ========================================');
    log('[BlogValidator] はてなブログHTML検証: ' + path.basename(filePath));
    log('[BlogValidator] ========================================');

    var forbidden = validateForbiddenElements(content, filePath);
    var seo = validateSeoStructure(content, filePath);

    var totalErrors = forbidden.errors;
    var totalWarnings = forbidden.warnings + seo.warnings;

    if (totalErrors === 0 && totalWarnings === 0) {
      log('[BlogValidator] OK 全チェック通過');
    } else {
      log('[BlogValidator] ----------------------------------------');
      log('[BlogValidator] 結果: ' + totalErrors + ' エラー, ' + totalWarnings + ' 警告');
      if (totalErrors > 0) {
        log('[BlogValidator] ERRORはブログ投稿前に必ず修正してください！');
      }
    }
    log('[BlogValidator] ========================================');
    log('');

  } catch (e) {
    // パースエラーはスキップ
  }

  return rawInput;
}

// stdin エントリーポイント
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
