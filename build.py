#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт сборки songbook.html из docx/txt файлов песен.
Запуск: python build.py
"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("Установите python-docx: pip install python-docx")
    raise


SONGS_DIR = Path("songs")
OUTPUT_FILE = Path("songbook.html")

# Разделитель песен: строка, состоящая только из 3+ звёздочек
SONG_SEPARATOR_RE = re.compile(r"^\*{3,}$")


def parse_docx(filepath: Path) -> list[dict]:
    """Парсит docx файл и возвращает список песен."""
    doc = Document(str(filepath))
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    return _split_into_songs(paragraphs)


def parse_txt(filepath: Path) -> list[dict]:
    """Парсит txt файл и возвращает список песен."""
    text = filepath.read_text(encoding="utf-8")
    paragraphs = [line.strip() for line in text.splitlines()]
    return _split_into_songs(paragraphs)


def _split_into_songs(paragraphs: list[str]) -> list[dict]:
    """
    Разбивает список параграфов на песни.
    Разделители — строки из 3+ звёздочек.
    """
    songs = []
    current_block = []

    for para in paragraphs:
        if SONG_SEPARATOR_RE.match(para):
            if current_block:
                song = _extract_song(current_block)
                if song:
                    songs.append(song)
                current_block = []
        else:
            current_block.append(para)

    # Последний блок
    if current_block:
        song = _extract_song(current_block)
        if song:
            songs.append(song)

    return songs


def _extract_song(block: list[str]) -> dict | None:
    """Извлекает название и текст из блока параграфов."""
    # Убираем ведущие и trailing пустые строки
    while block and not block[0]:
        block.pop(0)
    while block and not block[-1]:
        block.pop()

    if not block:
        return None

    title = block[0].upper()
    body = block[1:]

    # Убираем пустые строки в начале тела
    while body and not body[0]:
        body.pop(0)

    # Если название выглядит как примечание (в скобках), объединяем со следующей строкой
    if title.startswith("(") and title.endswith(")") and len(body) > 0:
        title = body[0]
        body = body[1:]
        while body and not body[0]:
            body.pop(0)

    # Фильтруем пустые строки из тела, но сохраняем одинарные как разделители куплетов
    # Заменяем последовательности пустых строк на один <br>
    cleaned_body = []
    empty_count = 0
    for line in body:
        if line:
            if empty_count > 0:
                cleaned_body.append("")
            cleaned_body.append(line)
            empty_count = 0
        else:
            empty_count += 1

    text = "\n".join(cleaned_body)

    return {
        "title": title,
        "text": text,
        "source": "",
    }


def _song_sort_key(song: dict) -> tuple[str, str]:
    """Возвращает ключ сортировки песни по названию для стабильного алфавитного порядка."""
    normalized_title = song["title"].strip().casefold().replace("ё", "е")
    normalized_title = re.sub(r"^[^\w]+", "", normalized_title)
    return normalized_title, song["title"]


def collect_songs() -> list[dict]:
    """Собирает все песни из папки songs."""
    all_songs = []
    if not SONGS_DIR.exists():
        print(f"Папка {SONGS_DIR} не найдена. Создайте её и положите туда docx/txt файлы.")
        return all_songs

    files = sorted(SONGS_DIR.iterdir())
    for filepath in files:
        if filepath.name.startswith("~$"):
            continue

        if filepath.suffix.lower() == ".docx":
            print(f"Обработка: {filepath.name}")
            songs = parse_docx(filepath)
            for s in songs:
                s["source"] = filepath.name
            all_songs.extend(songs)
        elif filepath.suffix.lower() == ".txt":
            print(f"Обработка: {filepath.name}")
            songs = parse_txt(filepath)
            for s in songs:
                s["source"] = filepath.name
            all_songs.extend(songs)

    all_songs.sort(key=_song_sort_key)

    return all_songs


def escape_js_string(s: str) -> str:
    """Экранирует строку для вставки в JavaScript."""
    s = s.replace("\\", "\\\\")
    s = s.replace('"', '\\"')
    s = s.replace("\n", "\\n")
    s = s.replace("\r", "")
    return s


def generate_html(songs: list[dict]) -> str:
    """Генерирует единый HTML файл."""
    songs_json = json.dumps(songs, ensure_ascii=False, indent=2)

    # Экранируем для JS вставки
    songs_js = songs_json.replace("<", "\\u003c").replace(">", "\\u003e")

    html = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<title>Песенник</title>
<style>
:root {
  --bg: #fdfbf7;
  --fg: #2c2c2c;
  --muted: #666;
  --accent: #8b4513;
  --accent-light: #d2691e;
  --card-bg: #fff;
  --border: #e0d8cf;
  --shadow: rgba(0,0,0,0.08);
  --toc-bg: #fff;
  --highlight: #fff3cd;
  --safe-bottom: env(safe-area-inset-bottom, 0px);
}

[data-theme="dark"] {
  --bg: #1a1a1a;
  --fg: #e0e0e0;
  --muted: #999;
  --accent: #d2691e;
  --accent-light: #ff8c42;
  --card-bg: #2a2a2a;
  --border: #444;
  --shadow: rgba(0,0,0,0.3);
  --toc-bg: #2a2a2a;
  --highlight: #5c4a1f;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

html, body {
  height: 100%;
  overflow: hidden;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  background: var(--bg);
  color: var(--fg);
  line-height: 1.7;
  transition: background 0.3s, color 0.3s;
  -webkit-text-size-adjust: 100%;
}

button,
input {
  font: inherit;
}

button {
  -webkit-appearance: none;
  appearance: none;
  touch-action: manipulation;
  -webkit-tap-highlight-color: transparent;
}

/* App shell */
#app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  height: 100dvh;
  min-height: -webkit-fill-available;
}

/* Header */
header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  background: var(--card-bg);
  border-bottom: 1px solid var(--border);
  box-shadow: 0 1px 4px var(--shadow);
  z-index: 10;
}

.title-menu-btn {
  flex: 1 1 auto;
  min-width: 0;
  background: none;
  border: none;
  color: var(--fg);
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  border-radius: 12px;
  text-align: left;
  transition: background 0.2s;
}

.title-menu-btn:hover, .title-menu-btn:active {
  background: var(--border);
}

.title-menu-btn::before {
  content: "☰";
  flex: 0 0 auto;
  font-size: 1.1rem;
}

header h1 {
  font-size: 1.1rem;
  font-weight: 600;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  min-width: 0;
}

.header-buttons {
  display: flex;
  gap: 4px;
  flex: 0 0 auto;
}

.icon-btn {
  background: none;
  border: none;
  color: var(--fg);
  font-size: 1.3rem;
  cursor: pointer;
  padding: 6px;
  border-radius: 50%;
  width: 40px;
  height: 40px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: background 0.2s;
}

.icon-btn:hover, .icon-btn:active {
  background: var(--border);
}

/* Content area */
main {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  -webkit-overflow-scrolling: touch;
  padding: 20px 20px calc(20px + var(--safe-bottom));
  overscroll-behavior: contain;
}

.song-text {
  font-size: 1.15rem;
  max-width: 680px;
  margin: 0 auto;
  word-wrap: break-word;
}

.song-text p {
  margin: 0;
}

.song-text .verse-gap {
  height: 1.2em;
}

.song-text.justify p {
  text-align: justify;
  text-align-last: justify;
  -moz-text-align-last: justify;
}

body.concert-mode header {
  transform: translateY(-100%);
  opacity: 0;
  pointer-events: none;
}

body.concert-mode main {
  padding-top: 28px;
}

body.concert-mode .song-text {
  font-size: 28px;
  max-width: 900px;
}

#concert-hint {
  position: fixed;
  top: 12px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 80;
  padding: 8px 12px;
  border-radius: 999px;
  background: var(--card-bg);
  color: var(--muted);
  box-shadow: 0 2px 10px var(--shadow);
  font-size: 0.85rem;
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.2s;
}

body.concert-mode.show-controls header {
  transform: translateY(0);
  opacity: 1;
  pointer-events: auto;
}

body.concert-mode.show-controls #concert-hint {
  opacity: 1;
}

/* TOC Overlay */
#toc-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.4);
  z-index: 100;
  opacity: 0;
  pointer-events: none;
  visibility: hidden;
  transition: opacity 0.25s;
}

#toc-overlay.open {
  opacity: 1;
  pointer-events: auto;
  visibility: visible;
}

#toc-panel {
  position: fixed;
  bottom: 0;
  left: 0;
  right: 0;
  max-height: min(80vh, 80dvh);
  padding-bottom: var(--safe-bottom);
  background: var(--toc-bg);
  border-radius: 20px 20px 0 0;
  box-shadow: 0 -4px 20px var(--shadow);
  z-index: 101;
  transform: translateY(100%);
  transition: transform 0.3s cubic-bezier(0.32,0.72,0,1);
  display: flex;
  flex-direction: column;
  pointer-events: none;
  visibility: hidden;
}

#toc-panel.open {
  transform: translateY(0);
  pointer-events: auto;
  visibility: visible;
}

.toc-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 16px 20px 8px;
  border-bottom: 1px solid var(--border);
}

.toc-header h2 {
  font-size: 1.1rem;
  font-weight: 600;
}

.toc-search {
  padding: 8px 20px 12px;
}

.toc-search input {
  width: 100%;
  padding: 10px 14px;
  font-size: 1rem;
  border: 1px solid var(--border);
  border-radius: 12px;
  background: var(--bg);
  color: var(--fg);
  outline: none;
}

.toc-search input:focus {
  border-color: var(--accent);
}

.toc-filters {
  display: flex;
  gap: 8px;
  padding: 0 20px 12px;
}

.chip-btn {
  border: 1px solid var(--border);
  background: var(--bg);
  color: var(--fg);
  border-radius: 999px;
  padding: 8px 12px;
  cursor: pointer;
}

.chip-btn.active {
  border-color: var(--accent);
  background: var(--accent);
  color: #fff;
}

.toc-list {
  overflow-y: auto;
  -webkit-overflow-scrolling: touch;
  padding: 0 20px 20px;
}

.toc-item {
  padding: 12px 4px;
  border-bottom: 1px solid var(--border);
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: space-between;
  transition: background 0.15s;
}

.toc-item:active {
  background: var(--border);
}

.toc-item-title {
  font-weight: 500;
  font-size: 0.95rem;
}

.toc-empty {
  text-align: center;
  padding: 30px;
  color: var(--muted);
}

/* Search overlay */
#search-overlay {
  position: fixed;
  inset: 0;
  background: var(--bg);
  z-index: 200;
  display: none;
  flex-direction: column;
}

#search-overlay.open {
  display: flex;
}

.search-header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 12px 16px;
  border-bottom: 1px solid var(--border);
  background: var(--card-bg);
}

.search-header input {
  flex: 1;
  padding: 10px 14px;
  font-size: 1rem;
  border: 1px solid var(--border);
  border-radius: 12px;
  background: var(--bg);
  color: var(--fg);
  outline: none;
}

.search-results {
  flex: 1;
  overflow-y: auto;
  padding: 0 16px;
}

.search-result-item {
  padding: 14px 4px;
  border-bottom: 1px solid var(--border);
  cursor: pointer;
}

.search-result-item:active {
  background: var(--border);
}

.search-result-title {
  font-weight: 600;
  margin-bottom: 4px;
}

.search-result-snippet {
  font-size: 0.9rem;
  color: var(--muted);
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

mark {
  background: var(--highlight);
  color: inherit;
  border-radius: 2px;
  padding: 0 2px;
}

/* Settings panel */
#settings-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.3);
  z-index: 150;
  opacity: 0;
  pointer-events: none;
  visibility: hidden;
  transition: opacity 0.2s;
}

#settings-overlay.open {
  opacity: 1;
  pointer-events: auto;
  visibility: visible;
}

#settings-panel {
  position: fixed;
  right: 12px;
  top: 64px;
  width: min(340px, calc(100vw - 24px));
  max-height: calc(100dvh - 88px - var(--safe-bottom));
  overflow-y: auto;
  background: var(--card-bg);
  color: var(--fg);
  border: 1px solid var(--border);
  border-radius: 18px;
  box-shadow: 0 8px 24px var(--shadow);
  z-index: 151;
  padding: 14px;
  display: none;
}

#settings-panel.open {
  display: block;
}

.settings-title {
  font-weight: 700;
  margin-bottom: 10px;
}

.settings-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 0;
  border-top: 1px solid var(--border);
}

.settings-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.settings-row input[type="range"] {
  width: 120px;
}

.settings-value {
  min-width: 42px;
  color: var(--muted);
  font-size: 0.9rem;
  text-align: right;
}

/* Bookmark indicator */
#bookmark-btn.active {
  color: var(--accent-light);
}

/* Font size controls */
.fontsize-controls {
  display: flex;
  align-items: center;
  gap: 4px;
}

/* Empty state */
.empty-state {
  text-align: center;
  padding: 40px 20px;
  color: var(--muted);
}

@media (max-width: 640px) {
  header {
    padding: 8px 10px;
  }

  header h1 {
    font-size: 1rem;
  }

  .icon-btn {
    width: 36px;
    height: 36px;
    font-size: 1.15rem;
  }

  .hide-mobile {
    display: none;
  }
}
#ios-debug {
  position: fixed;
  top: 4px;
  right: 4px;
  background: rgba(255,0,0,0.85);
  color: #fff;
  padding: 6px 10px;
  z-index: 99999;
  font-size: 12px;
  border-radius: 6px;
  display: none;
  max-width: 80vw;
  word-break: break-word;
}
#ios-debug.ok {
  background: rgba(0,128,0,0.85);
}
/* noscript fallback for iOS Quick Look */
noscript {
  display: block;
  position: fixed;
  inset: 0;
  z-index: 99999;
  background: var(--bg);
  color: var(--fg);
}
noscript .noscript-box {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  text-align: center;
  padding: 32px 24px;
  max-width: 340px;
  width: 90%;
}
noscript .noscript-box h2 {
  font-size: 1.4rem;
  margin-bottom: 16px;
  color: var(--accent);
}
noscript .noscript-box p {
  font-size: 1rem;
  margin-bottom: 12px;
  line-height: 1.6;
  color: var(--fg);
}
noscript .noscript-box .hint {
  font-size: 0.9rem;
  color: var(--muted);
  margin-top: 20px;
  padding-top: 16px;
  border-top: 1px solid var(--border);
}
</style>
</head>
<body>
<div id="ios-debug"></div>
<noscript>
  <div class="noscript-box">
    <h2>JavaScript отключен</h2>
    <p>На iPhone файл открывается в режиме предпросмотра, где JavaScript заблокирован системой.</p>
    <p>Откройте этот файл в браузере Safari или через приложение Documents / HTML Viewer.</p>
    <p class="hint">Или загрузите файл на веб-сайт и откройте по ссылке.</p>
  </div>
</noscript>
<div id="app">
  <header>
    <button class="title-menu-btn" id="title-menu-btn" type="button" title="Открыть оглавление" aria-label="Открыть оглавление">
      <h1 id="current-title">Песенник</h1>
    </button>
    <div class="header-buttons">
      <button class="icon-btn" id="bookmark-btn" type="button" title="Закладка">☆</button>
      <button class="icon-btn" id="btn-search" type="button" title="Поиск">🔍</button>
      <button class="icon-btn" id="btn-settings" type="button" title="Настройки">⚙</button>
    </div>
  </header>
  <main id="content">
    <div class="empty-state">Выберите песню из меню ☰</div>
  </main>
</div>
<div id="concert-hint">Тап по тексту скрывает/показывает управление</div>

<!-- TOC -->
<div id="toc-overlay"></div>
<div id="toc-panel">
  <div class="toc-header">
    <h2>Оглавление</h2>
    <button class="icon-btn" id="toc-close" type="button">✕</button>
  </div>
  <div class="toc-search">
    <input type="text" id="toc-filter" placeholder="Фильтр по названию..." autocomplete="off">
  </div>
  <div class="toc-filters">
    <button class="chip-btn active" id="toc-filter-all" type="button">Все</button>
    <button class="chip-btn" id="toc-filter-favorites" type="button">Избранное</button>
  </div>
  <div class="toc-list" id="toc-list"></div>
</div>

<!-- Settings -->
<div id="settings-overlay"></div>
<div id="settings-panel">
  <div class="settings-title">Настройки</div>
  <div class="settings-row">
    <span>Размер текста</span>
    <div class="settings-actions">
      <button class="icon-btn fontsize-controls" id="btn-decrease" type="button" title="Меньше">A-</button>
      <button class="icon-btn fontsize-controls" id="btn-increase" type="button" title="Больше">A+</button>
    </div>
  </div>
  <div class="settings-row">
    <span>Выравнивание</span>
    <button class="icon-btn" id="btn-align" type="button" title="Выравнивание">⬅</button>
  </div>
  <div class="settings-row">
    <span>Тема</span>
    <button class="icon-btn" id="btn-theme" type="button" title="Тема">☀</button>
  </div>
  <div class="settings-row">
    <span>Концертный режим</span>
    <button class="chip-btn" id="btn-concert" type="button">Выкл</button>
  </div>
  <div class="settings-row">
    <span>Автопрокрутка</span>
    <button class="chip-btn" id="btn-autoscroll" type="button">Старт</button>
  </div>
  <div class="settings-row">
    <span>Скорость</span>
    <div class="settings-actions">
      <input type="range" id="autoscroll-speed" min="1" max="10" step="1" value="3">
      <span class="settings-value" id="autoscroll-speed-value">3</span>
    </div>
  </div>
</div>

<!-- Search -->
<div id="search-overlay">
  <div class="search-header">
    <button class="icon-btn" id="search-back" type="button">←</button>
    <input type="text" id="search-input" placeholder="Поиск по тексту..." autocomplete="off">
  </div>
  <div class="search-results" id="search-results"></div>
</div>

<script>
(function() {
  'use strict';

  // === DATA ===
  const SONGS = __SONGS_DATA__;

  // === SAFE STORAGE (fallback for file://) ===
  const _memStore = {};
  const storage = {
    getItem(k) {
      try { return localStorage.getItem(k); } catch(e) { return _memStore[k] || null; }
    },
    setItem(k, v) {
      try { localStorage.setItem(k, v); } catch(e) { _memStore[k] = v; }
    }
  };

  // === STATE ===
  let currentIndex = -1;
  let fontSize = parseInt(storage.getItem('sb_fontSize')) || 18;
  let theme = storage.getItem('sb_theme') || 'light';
  let alignMode = storage.getItem('sb_align') || 'left';
  let bookmarks = JSON.parse(storage.getItem('sb_bookmarks') || '[]');
  let tocFavoritesOnly = false;
  let concertMode = storage.getItem('sb_concert') === '1';
  let autoscrollSpeed = parseInt(storage.getItem('sb_autoscrollSpeed')) || 3;
  let autoscrollTimer = null;
  let controlsTimer = null;
  let touchStartX = 0;
  let touchStartY = 0;

  // === DEBUG ===
  var _debugEl = document.getElementById('ios-debug');
  var _debugTimer = null;
  function _debug(msg) {
    if (!_debugEl) return;
    clearTimeout(_debugTimer);
    _debugEl.classList.remove('ok');
    _debugEl.style.display = 'block';
    _debugEl.textContent = String(msg).slice(0, 200);
  }
  function _debugOk(msg) {
    if (!_debugEl) return;
    clearTimeout(_debugTimer);
    _debugEl.style.display = 'block';
    _debugEl.classList.add('ok');
    _debugEl.textContent = String(msg).slice(0, 200);
    _debugTimer = setTimeout(function() {
      if (_debugEl) _debugEl.style.display = 'none';
    }, 3000);
  }
  window.onerror = function(msg, url, line) {
    _debug('JS Error: ' + msg + ' (line ' + line + ')');
    return false;
  };

  // === DOM ===
  function $(id) {
    var el = document.getElementById(id);
    if (!el) _debug('Missing #' + id);
    return el;
  }
  var content;
  var currentTitle;
  var tocOverlay;
  var tocPanel;
  var tocList;
  var tocFilter;
  var searchOverlay;
  var searchInput;
  var searchResults;
  var bookmarkBtn;
  var settingsOverlay;
  var settingsPanel;
  var autoscrollSpeedInput;
  var autoscrollSpeedValue;

  function _bindClick(el, fn) {
    if (!el) return;
    el.addEventListener('click', fn);
    // iOS Safari: empty touchstart handler forces click generation
    el.addEventListener('touchstart', function() {}, {passive: true});
  }

  // === INIT ===
  function init() {
    try {
      content = $('content');
      currentTitle = $('current-title');
      tocOverlay = $('toc-overlay');
      tocPanel = $('toc-panel');
      tocList = $('toc-list');
      tocFilter = $('toc-filter');
      searchOverlay = $('search-overlay');
      searchInput = $('search-input');
      searchResults = $('search-results');
      bookmarkBtn = $('bookmark-btn');
      settingsOverlay = $('settings-overlay');
      settingsPanel = $('settings-panel');
      autoscrollSpeedInput = $('autoscroll-speed');
      autoscrollSpeedValue = $('autoscroll-speed-value');

      applyTheme();
      applyFontSize();
      applyAlign();
      applyConcertMode(false);
      applyAutoscrollSpeed();
      renderTOC(SONGS);
      restoreLastPosition();

      // Events
      _bindClick($('title-menu-btn'), openTOC);
      _bindClick(tocOverlay, closeTOC);
      _bindClick($('toc-close'), closeTOC);
      if (tocFilter) tocFilter.addEventListener('input', function(e) { filterTOC(e.target.value); });
      _bindClick($('toc-filter-all'), function() { setTOCFilter(false); });
      _bindClick($('toc-filter-favorites'), function() { setTOCFilter(true); });

      _bindClick($('btn-settings'), toggleSettings);
      _bindClick(settingsOverlay, closeSettings);
      _bindClick($('btn-theme'), toggleTheme);
      _bindClick($('btn-align'), toggleAlign);
      _bindClick($('btn-increase'), function() { changeFontSize(+1); });
      _bindClick($('btn-decrease'), function() { changeFontSize(-1); });
      _bindClick($('btn-concert'), toggleConcertMode);
      _bindClick($('btn-autoscroll'), toggleAutoscroll);
      if (autoscrollSpeedInput) autoscrollSpeedInput.addEventListener('input', function(e) { setAutoscrollSpeed(parseInt(e.target.value)); });
      _bindClick(bookmarkBtn, toggleBookmark);

      _bindClick($('btn-search'), openSearch);
      _bindClick($('search-back'), closeSearch);
      if (searchInput) searchInput.addEventListener('input', function(e) { doSearch(e.target.value); });
      if (content) {
        content.addEventListener('touchstart', onTouchStart, {passive: true});
        content.addEventListener('touchend', onTouchEnd, {passive: true});
        content.addEventListener('click', onContentTap);
      }

      // Keyboard shortcuts
      document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
          closeTOC();
          closeSearch();
          closeSettings();
        }
      });

      _debugOk('Init OK');
    } catch(e) {
      _debug('Init failed: ' + e.message);
    }
  }

  // === RENDER ===
  function renderSong(index) {
    if (index < 0 || index >= SONGS.length) return;
    currentIndex = index;
    const song = SONGS[index];
    currentTitle.textContent = song.title;

    // Each line as a paragraph so justify can work per line
    const lines = song.text.split('\n');
    let html = '';
    lines.forEach(function(line) {
      if (line.trim() === '') {
        html += '<p class="verse-gap"></p>';
      } else {
        html += '<p>' + escapeHtml(line) + '</p>';
      }
    });

    content.innerHTML = `<div class="song-text">${html}</div>`;
    updateBookmarkIcon();
    applyAlign();
    applyFontSize();
    applyConcertMode(false);
    savePosition();
    closeTOC();
    closeSearch();
    closeSettings();
  }

  function renderTOC(list) {
    renderTOCItems(getVisibleTOCItems(tocFilter.value), 'Нет песен');
  }

  function getVisibleTOCItems(query) {
    const q = query.toLowerCase().trim();
    const result = [];
    SONGS.forEach(function(song, idx) {
      const favoriteMatch = !tocFavoritesOnly || bookmarks.indexOf(idx) >= 0;
      const textMatch = !q || song.title.toLowerCase().indexOf(q) >= 0 || song.text.toLowerCase().indexOf(q) >= 0;
      if (favoriteMatch && textMatch) result.push({song: song, idx: idx});
    });
    return result;
  }

  function renderTOCItems(items, emptyText) {
    if (items.length === 0) {
      tocList.innerHTML = '<div class="toc-empty">' + emptyText + '</div>';
      return;
    }
    tocList.innerHTML = items.map(function(item) { return `
      <div class="toc-item" data-index="${item.idx}">
        <div>
          <div class="toc-item-title">${escapeHtml(item.song.title)}</div>
        </div>
        <span>${bookmarks.indexOf(item.idx) >= 0 ? '★' : ''}</span>
      </div>
    ` }).join('');

    tocList.querySelectorAll('.toc-item').forEach(function(item) {
      item.addEventListener('click', function() { renderSong(parseInt(item.dataset.index)); });
    });
  }

  function setTOCFilter(favoritesOnly) {
    tocFavoritesOnly = favoritesOnly;
    $('toc-filter-all').classList.toggle('active', !favoritesOnly);
    $('toc-filter-favorites').classList.toggle('active', favoritesOnly);
    filterTOC(tocFilter.value);
  }

  function filterTOC(query) {
    const items = getVisibleTOCItems(query);
    const emptyText = tocFavoritesOnly ? 'В избранном ничего не найдено' : 'Ничего не найдено';
    renderTOCItems(items, emptyText);
  }

  // === SEARCH ===
  function openSearch() {
    searchOverlay.classList.add('open');
    setTimeout(function() { searchInput.focus(); }, 0);
  }

  function closeSearch() {
    searchOverlay.classList.remove('open');
    searchInput.value = '';
    searchResults.innerHTML = '';
  }

  function doSearch(query) {
    const q = query.toLowerCase().trim();
    if (!q) {
      searchResults.innerHTML = '';
      return;
    }
    const results = [];
    SONGS.forEach(function(song, idx) {
      const titleMatch = song.title.toLowerCase().indexOf(q) >= 0;
      const textMatch = song.text.toLowerCase().indexOf(q) >= 0;
      if (titleMatch || textMatch) {
        let snippet = '';
        if (textMatch) {
          const pos = song.text.toLowerCase().indexOf(q);
          const start = Math.max(0, pos - 40);
          const end = Math.min(song.text.length, pos + q.length + 60);
          snippet = song.text.slice(start, end);
          if (start > 0) snippet = '...' + snippet;
          if (end < song.text.length) snippet += '...';
        } else {
          snippet = song.text.slice(0, 100) + (song.text.length > 100 ? '...' : '');
        }
        results.push({idx: idx, song: song, snippet: snippet});
      }
    });

    if (results.length === 0) {
      searchResults.innerHTML = '<div class="toc-empty">Ничего не найдено</div>';
      return;
    }

    searchResults.innerHTML = results.map(function(r) {
      const highlighted = highlightText(r.snippet, query);
      return `
        <div class="search-result-item" data-index="${r.idx}">
          <div class="search-result-title">${escapeHtml(r.song.title)}</div>
          <div class="search-result-snippet">${highlighted}</div>
        </div>
      `;
    }).join('');

    searchResults.querySelectorAll('.search-result-item').forEach(function(item) {
      item.addEventListener('click', function() { renderSong(parseInt(item.dataset.index)); });
    });
  }

  function highlightText(text, query) {
    const re = new RegExp('(' + escapeRegExp(query) + ')', 'gi');
    return escapeHtml(text).replace(re, '<mark>$1</mark>');
  }

  // === BOOKMARKS ===
  function toggleBookmark() {
    if (currentIndex < 0) return;
    const pos = bookmarks.indexOf(currentIndex);
    if (pos >= 0) {
      bookmarks.splice(pos, 1);
    } else {
      bookmarks.push(currentIndex);
      bookmarks.sort(function(a,b) { return a-b; });
    }
    storage.setItem('sb_bookmarks', JSON.stringify(bookmarks));
    updateBookmarkIcon();
    renderTOC(SONGS); // refresh stars
  }

  function updateBookmarkIcon() {
    if (currentIndex >= 0 && bookmarks.indexOf(currentIndex) >= 0) {
      bookmarkBtn.textContent = '★';
      bookmarkBtn.classList.add('active');
    } else {
      bookmarkBtn.textContent = '☆';
      bookmarkBtn.classList.remove('active');
    }
  }

  // === THEME ===
  function toggleTheme() {
    theme = theme === 'light' ? 'dark' : 'light';
    applyTheme();
    storage.setItem('sb_theme', theme);
  }

  function applyTheme() {
    document.documentElement.setAttribute('data-theme', theme);
    $('btn-theme').textContent = theme === 'light' ? '☀' : '☾';
  }

  // === FONT SIZE ===
  function changeFontSize(delta) {
    fontSize = Math.max(12, Math.min(32, fontSize + delta));
    applyFontSize();
    storage.setItem('sb_fontSize', fontSize);
  }

  function applyFontSize() {
    const el = document.querySelector('.song-text');
    if (el) el.style.fontSize = fontSize + 'px';
  }

  // === ALIGN ===
  function toggleAlign() {
    alignMode = alignMode === 'left' ? 'justify' : 'left';
    applyAlign();
    storage.setItem('sb_align', alignMode);
  }

  function applyAlign() {
    const el = document.querySelector('.song-text');
    if (el) {
      if (alignMode === 'justify') {
        el.classList.add('justify');
      } else {
        el.classList.remove('justify');
      }
    }
    $('btn-align').textContent = alignMode === 'left' ? '⬅' : '▤';
  }

  // === SETTINGS ===
  function toggleSettings() {
    if (settingsPanel.classList.contains('open')) {
      closeSettings();
    } else {
      settingsOverlay.classList.add('open');
      settingsPanel.classList.add('open');
    }
  }

  function closeSettings() {
    settingsOverlay.classList.remove('open');
    settingsPanel.classList.remove('open');
  }

  // === NAVIGATION ===
  function goToSong(delta) {
    if (SONGS.length === 0) return;
    const base = currentIndex < 0 ? 0 : currentIndex;
    const next = Math.max(0, Math.min(SONGS.length - 1, base + delta));
    if (next !== currentIndex) renderSong(next);
  }

  function onTouchStart(e) {
    if (!e.changedTouches || e.changedTouches.length === 0) return;
    touchStartX = e.changedTouches[0].clientX;
    touchStartY = e.changedTouches[0].clientY;
  }

  function onTouchEnd(e) {
    if (!e.changedTouches || e.changedTouches.length === 0) return;
    const dx = e.changedTouches[0].clientX - touchStartX;
    const dy = e.changedTouches[0].clientY - touchStartY;
    if (Math.abs(dx) < 70 || Math.abs(dx) < Math.abs(dy) * 1.5) return;
    goToSong(dx < 0 ? 1 : -1);
  }

  // === CONCERT MODE ===
  function toggleConcertMode() {
    concertMode = !concertMode;
    storage.setItem('sb_concert', concertMode ? '1' : '0');
    applyConcertMode(true);
  }

  function applyConcertMode(showControls) {
    document.body.classList.toggle('concert-mode', concertMode);
    $('btn-concert').textContent = concertMode ? 'Вкл' : 'Выкл';
    if (concertMode && showControls) showConcertControls();
    if (!concertMode) document.body.classList.remove('show-controls');
  }

  function onContentTap() {
    if (!concertMode) return;
    if (document.body.classList.contains('show-controls')) {
      hideConcertControls();
    } else {
      showConcertControls();
    }
  }

  function showConcertControls() {
    document.body.classList.add('show-controls');
    clearTimeout(controlsTimer);
    controlsTimer = setTimeout(hideConcertControls, 3500);
  }

  function hideConcertControls() {
    document.body.classList.remove('show-controls');
    clearTimeout(controlsTimer);
  }

  // === AUTOSCROLL ===
  function toggleAutoscroll() {
    if (autoscrollTimer) {
      stopAutoscroll();
    } else {
      startAutoscroll();
    }
  }

  function startAutoscroll() {
    stopAutoscroll();
    $('btn-autoscroll').textContent = 'Стоп';
    autoscrollTimer = setInterval(function() {
      const step = Math.max(1, autoscrollSpeed) / 2;
      content.scrollTop += step;
      if (content.scrollTop + content.clientHeight >= content.scrollHeight - 2) stopAutoscroll();
    }, 50);
  }

  function stopAutoscroll() {
    if (autoscrollTimer) clearInterval(autoscrollTimer);
    autoscrollTimer = null;
    $('btn-autoscroll').textContent = 'Старт';
  }

  function setAutoscrollSpeed(value) {
    autoscrollSpeed = Math.max(1, Math.min(10, value || 3));
    storage.setItem('sb_autoscrollSpeed', autoscrollSpeed);
    applyAutoscrollSpeed();
    if (autoscrollTimer) startAutoscroll();
  }

  function applyAutoscrollSpeed() {
    autoscrollSpeedInput.value = autoscrollSpeed;
    autoscrollSpeedValue.textContent = autoscrollSpeed;
  }

  // === TOC PANEL ===
  function openTOC() {
    tocOverlay.classList.add('open');
    tocPanel.classList.add('open');
    tocFilter.value = '';
    setTOCFilter(false);
    renderTOC(SONGS);
    setTimeout(function() { tocFilter.focus(); }, 250);
  }

  function closeTOC() {
    tocOverlay.classList.remove('open');
    tocPanel.classList.remove('open');
  }

  // === PERSISTENCE ===
  function savePosition() {
    storage.setItem('sb_lastIndex', currentIndex);
  }

  function restoreLastPosition() {
    const saved = parseInt(storage.getItem('sb_lastIndex'));
    if (!isNaN(saved) && saved >= 0 && saved < SONGS.length) {
      renderSong(saved);
    }
  }

  // === UTILS ===
  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
  }

  try {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', init, {once: true});
    } else {
      init();
    }
  } catch(e) {
    _debug('Bootstrap failed: ' + e.message);
  }
})();
</script>
</body>
</html>
'''
    return html.replace('__SONGS_DATA__', songs_js)


def deploy_to_github():
    """Автоматический деплой собранного файла на GitHub Pages."""
    print("\n=== Деплой на GitHub Pages ===")

    # Проверяем, есть ли git
    try:
        subprocess.run(["git", "--version"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("❌ Git не найден. Установите Git: https://git-scm.com/")
        return False

    repo_path = Path.cwd()
    git_dir = repo_path / ".git"

    # Проверяем, является ли текущая папка git-репозиторием
    if not git_dir.exists():
        print("❌ В текущей папке нет git-репозитория.")
        print("   Сначала создайте репозиторий на GitHub и свяжите его:")
        print("   1. Создайте репозиторий на https://github.com/new")
        print("   2. Выполните в этой папке:")
        print(f"      git init")
        print(f"      git remote add origin https://github.com/ВАШ_ЛОГИН/ВАШ_РЕПО.git")
        print("   3. Запустите сборку снова: python build.py --deploy")
        return False

    # Проверяем, есть ли remote origin
    try:
        result = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            capture_output=True, text=True, check=True
        )
        remote_url = result.stdout.strip()
        print(f"✓ Remote: {remote_url}")
    except subprocess.CalledProcessError:
        print("❌ Не найден remote 'origin'.")
        print("   Добавьте remote:")
        print(f"   git remote add origin https://github.com/ВАШ_ЛОГИН/ВАШ_РЕПО.git")
        return False

    # Добавляем файл, коммитим, пушим
    try:
        subprocess.run(["git", "add", "songbook.html"], check=True, capture_output=True)

        # Проверяем, есть ли изменения
        status = subprocess.run(
            ["git", "status", "--porcelain", "songbook.html"],
            capture_output=True, text=True, check=True
        )
        if not status.stdout.strip():
            print("✓ Изменений в songbook.html нет (файл уже актуален)")
            return True

        subprocess.run(
            ["git", "commit", "-m", "Обновление песенника"],
            check=True, capture_output=True
        )
        print("✓ Коммит создан")

        # Определяем текущую ветку
        branch_result = subprocess.run(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            capture_output=True, text=True, check=True
        )
        branch = branch_result.stdout.strip()

        subprocess.run(
            ["git", "push", "origin", branch],
            check=True, capture_output=True
        )
        print(f"✓ Отправлено на GitHub (ветка: {branch})")

        # Извлекаем имя репозитория из URL
        if "github.com" in remote_url:
            # Обработка HTTPS и SSH URL
            clean_url = remote_url.replace(".git", "").replace("git@github.com:", "https://github.com/")
            pages_url = clean_url.replace("https://github.com/", "https://").replace("/", ".", 1)
            print(f"\n📱 GitHub Pages будет доступен через 1-2 минуты:")
            print(f"   {pages_url}/songbook.html")
            print(f"\n   ⚠️  Убедитесь, что в настройках репозитория включен GitHub Pages:")
            print(f"      Settings → Pages → Source → Deploy from a branch → {branch} / root")

        return True

    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка при деплое: {e}")
        print(f"   stderr: {e.stderr.decode('utf-8', errors='replace') if e.stderr else '---'}")
        return False


def main():
    deploy = "--deploy" in sys.argv

    print("=== Сборка песенника ===")
    songs = collect_songs()
    print(f"Всего песен найдено: {len(songs)}")
    if not songs:
        print("Нет песен для сборки. Выход.")
        return

    html = generate_html(songs)
    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"Готово: {OUTPUT_FILE.absolute()}")
    print(f"Размер: {OUTPUT_FILE.stat().st_size / 1024:.1f} КБ")

    if deploy:
        deploy_to_github()
    else:
        print("\n💡 Для автодеплоя на GitHub Pages запустите:")
        print("   python build.py --deploy")


if __name__ == "__main__":
    main()
