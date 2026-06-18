# Copyright Notice:
# Copyright 2017-2025 Distributed Management Task Force, Inc. All rights reserved.
# License: BSD 3-Clause License. For full text see link: https://github.com/DMTF/Redfish-Use-Case-Checkers/blob/main/LICENSE.md

import html as html_mod
import json
from datetime import datetime

import openpyxl
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter

from redfish_use_case_checkers import redfish_logo
from redfish_use_case_checkers.system_under_test import SystemUnderTest

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Redfish Use Case Checkers &mdash; Test Report</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html, body {{ height: 100%; }}
    body {{
      font-family: "Segoe UI", system-ui, Arial, sans-serif;
      font-size: 13px;
      background: #f0f2f5;
      color: #1a1a2e;
      display: flex;
      flex-direction: column;
    }}

    /* ════════════════════════════
       TOP NAVBAR
       ════════════════════════════ */
    .top-nav {{
      position: fixed;
      top: 0; left: 0; right: 0;
      height: 56px;
      background: linear-gradient(90deg, #0d1b2a 0%, #1a3a5c 60%, #1565c0 100%);
      display: flex;
      align-items: center;
      padding: 0 16px;
      gap: 10px;
      z-index: 1000;
      box-shadow: 0 2px 10px rgba(0,0,0,0.35);
    }}
    .top-nav img {{ height: 34px; border-radius: 4px; flex-shrink: 0; }}
    .nav-brand {{ flex-shrink: 0; }}
    .top-nav-title {{
      color: #fff;
      font-size: 15px;
      font-weight: 700;
      letter-spacing: 0.3px;
      white-space: nowrap;
    }}
    .top-nav-sub {{
      color: rgba(255,255,255,0.55);
      font-size: 11px;
      white-space: nowrap;
    }}
    .top-nav-search-wrap {{
      position: absolute;
      left: 50%; transform: translateX(-50%);
      display: flex; flex-direction: column; align-items: center;
      width: 380px;
      pointer-events: auto;
    }}
    .nav-filter-wrap {{ position: relative; width: 100%; }}
    .nav-filter-wrap .fi {{
      position: absolute; left: 10px; top: 50%;
      transform: translateY(-50%);
      color: rgba(255,255,255,0.5); font-size: 14px;
      pointer-events: none;
    }}
    #uriFilter {{
      width: 100%;
      padding: 6px 30px 6px 30px;
      border: 1px solid rgba(255,255,255,0.25);
      border-radius: 20px;
      font-size: 12px; outline: none;
      transition: border-color 0.2s, box-shadow 0.2s;
      font-family: "Cascadia Code","Consolas",monospace;
      color: #fff;
      background: rgba(255,255,255,0.12);
    }}
    #uriFilter::placeholder {{ color: rgba(255,255,255,0.45); }}
    #uriFilter:focus {{
      border-color: #90caf9;
      box-shadow: 0 0 0 3px rgba(144,202,249,0.2);
      background: rgba(255,255,255,0.2);
    }}
    #filterClear {{
      position: absolute; right: 9px; top: 50%;
      transform: translateY(-50%);
      background: none; border: none; cursor: pointer;
      color: rgba(255,255,255,0.5); font-size: 13px; line-height: 1;
      display: none;
    }}
    #filterClear:hover {{ color: #ef5350; }}
    .nav-filter-meta {{
      font-size: 10px; color: rgba(255,255,255,0.5);
      margin-top: 2px; text-align: center;
    }}
    .nav-filter-meta b {{ color: #90caf9; }}
    .nav-hamburger {{
      display: none;
      align-items: center; justify-content: center;
      width: 36px; height: 36px;
      background: rgba(255,255,255,0.1);
      border: 1px solid rgba(255,255,255,0.2);
      border-radius: 6px; color: #fff; font-size: 18px;
      cursor: pointer; flex-shrink: 0;
    }}
    .nav-hamburger:hover {{ background: rgba(255,255,255,0.2); }}
    .top-nav-meta {{
      color: rgba(255,255,255,0.7);
      font-size: 11px;
      text-align: right;
      line-height: 1.5;
      margin-left: auto;
      flex-shrink: 0;
    }}
    .top-nav-meta a {{ color: #90caf9; text-decoration: none; }}
    .top-nav-meta a:hover {{ text-decoration: underline; }}

    /* ════════════════════════════
       APP SHELL  (sidebar + main)
       ════════════════════════════ */
    .app-shell {{
      display: flex;
      margin-top: 56px;
      height: calc(100vh - 56px);
      overflow: hidden;
    }}
    .sidebar-overlay {{
      display: none;
      position: fixed;
      inset: 56px 0 0 0;
      background: rgba(0,0,0,0.45);
      z-index: 99;
    }}
    .sidebar-overlay.show {{ display: block; }}
    .sidebar {{
      width: 300px;
      min-width: 300px;
      background: #fff;
      border-right: 1px solid #dde3ec;
      height: 100%;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      box-shadow: 2px 0 8px rgba(0,0,0,0.04);
      z-index: 100;
      transition: left 0.3s ease;
    }}
    .sidebar::-webkit-scrollbar {{ width: 5px; }}
    .sidebar::-webkit-scrollbar-thumb {{ background: #c8d3e0; border-radius: 3px; }}
    .main-content {{
      flex: 1;
      min-width: 0;
      padding: 20px 24px 60px;
      overflow-y: auto;
      overflow-x: hidden;
      height: 100%;
    }}
    .main-content::-webkit-scrollbar {{ width: 6px; }}
    .main-content::-webkit-scrollbar-thumb {{ background: #c8d3e0; border-radius: 3px; }}
    a {{ color: #0d6efd; }}

    /* ── Sidebar sections ── */
    .sidebar-section {{
      padding: 16px 18px 10px;
      border-bottom: 1px solid #edf0f5;
    }}
    .sidebar-section:last-child {{ border-bottom: none; }}
    .sidebar-label {{
      font-size: 10px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #8a97aa;
      margin-bottom: 10px;
    }}

    /* System info table in sidebar */
    .sys-table {{ width: 100%; border-collapse: collapse; }}
    .sys-table td {{
      padding: 4px 0;
      font-size: 12px;
      border: none;
      background: transparent;
      vertical-align: top;
      color: #333;
      word-break: break-all;
    }}
    .sys-table td:first-child {{
      font-weight: 600;
      color: #6c757d;
      white-space: nowrap;
      padding-right: 10px;
      min-width: 90px;
    }}

    /* Score tiles in sidebar */
    .score-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }}
    .score-tile {{
      border-radius: 8px;
      padding: 10px 12px;
      color: #fff;
      text-align: center;
      font-weight: 700;
    }}
    .score-tile .snum {{ font-size: 26px; line-height: 1.1; }}
    .score-tile .slbl {{ font-size: 10px; text-transform: uppercase; letter-spacing: 0.7px; margin-top: 2px; opacity: 0.9; }}
    .st-pass {{ background: linear-gradient(135deg, #27ae60, #2ecc71); }}
    .st-warn {{ background: linear-gradient(135deg, #d68910, #f39c12); }}
    .st-fail {{ background: linear-gradient(135deg, #c0392b, #e74c3c); }}
    .st-skip {{ background: linear-gradient(135deg, #5d6d7e, #85929e); }}

    /* ── Section ── */
    .section {{ margin-bottom: 1.5rem; }}
    .section-header {{
      background: #f7f9fc;
      color: #0d1b2a;
      padding: 10px 16px;
      border-radius: 8px 8px 0 0;
      font-weight: 600;
      font-size: 13px;
      letter-spacing: .01em;
      display: flex;
      align-items: center;
      gap: .55rem;
      cursor: pointer;
      user-select: none;
      border-bottom: 1px solid #dde3ec;
    }}
    .section-arrow {{ font-size: .7rem; color: #6c757d; transition: transform .2s; }}
    .section-header.collapsed .section-arrow {{ transform: rotate(-90deg); }}
    .section-body {{
      border: 1px solid #dde3ec;
      border-top: none;
      border-radius: 0 0 8px 8px;
      overflow: hidden;
      box-shadow: 0 1px 3px rgba(0,0,0,.05);
      margin-bottom: 8px;
    }}

    /* ── Test block ── */
    .test-block {{ border-bottom: 1px solid #dde3ec; }}
    .test-block:last-child {{ border-bottom: none; }}
    .test-heading {{
      background: #fff;
      padding: 10px 16px;
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 1rem;
    }}
    .test-heading:hover {{ background: #f8fafc; }}
    .test-name   {{ font-weight: 700; font-size: 13px; color: #0d1b2a; font-family: "Cascadia Code", "Consolas", monospace; }}
    .test-desc   {{ font-size: 11px; color: #6c757d; margin-top: .12rem; font-style: italic; }}
    .test-detail {{ font-size: 11px; color: #6c757d; margin-top: .08rem; font-style: italic; }}
    .test-toggle {{ flex-shrink: 0; font-size: .65rem; color: #6c757d; margin-top: .3rem; transition: transform .2s; }}
    .test-body         {{ display: block; }}
    .test-body.hidden  {{ display: none; }}

    /* ── Results table ── */
    .test-results-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
    .test-results-table thead tr {{ background: #f0f4f8; }}
    .test-results-table th {{
      padding: 7px 12px;
      text-align: left;
      font-weight: 700;
      font-size: 11px;
      color: #2c3e50;
      border-bottom: 2px solid #dde3ec;
    }}
    .test-results-table td {{
      padding: 5px 12px;
      border-bottom: 1px solid #f0f2f5;
      vertical-align: top;
      color: #333;
      background: #fff;
    }}
    .test-results-table tr:last-child td {{ border-bottom: none; }}
    .test-results-table tr:hover td {{ background: #f8fafc; }}
    .col-op     {{ width: 38%; }}
    .col-result {{ width:  9%; white-space: nowrap; }}
    .col-msg    {{ width: 53%; }}

    /* ── Badges ── */
    .badge {{
      display: inline-block;
      padding: .15rem .55rem;
      border-radius: 9999px;
      font-size: .69rem;
      font-weight: 700;
      letter-spacing: .04em;
      text-transform: uppercase;
    }}
    .badge-pass {{ background: #d4edda; color: #145a32; }}
    .badge-warn {{ background: #fff3cd; color: #7d6008; }}
    .badge-fail {{ background: #f8d7da; color: #7b241c; }}
    .badge-skip {{ background: #f5f7fa; color: #7f8c8d; }}

    /* ── Toolbar ── */
    .toolbar {{
      display: flex;
      align-items: center;
      gap: .6rem;
      margin-bottom: 1rem;
    }}
    .toolbar-btn {{
      background: #0d6efd;
      color: #fff;
      border: none;
      border-radius: 5px;
      padding: 5px 13px;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      letter-spacing: .02em;
      transition: background .15s;
      box-shadow: 0 2px 5px rgba(13,110,253,.3);
    }}
    .toolbar-btn:hover {{ background: #0b5ed7; }}

    /* ── Scroll-to-top ── */
    #scrollTopBtn {{
      position: fixed;
      bottom: 24px; right: 24px;
      width: 42px; height: 42px;
      border-radius: 50%;
      background: #0d6efd; color: #fff;
      border: none; cursor: pointer;
      font-size: 20px; line-height: 42px;
      text-align: center;
      box-shadow: 0 4px 14px rgba(13,110,253,0.45);
      display: none; z-index: 1100;
      transition: background 0.2s, transform 0.2s;
    }}
    #scrollTopBtn:hover {{ background: #0b5ed7; transform: translateY(-2px); }}

    /* ── Inline result summary chips (at most 4 per test) ── */
    .result-chips {{ display: flex; gap: .4rem; margin-top: .45rem; flex-wrap: wrap; }}
    .rchip {{
      display: inline-flex; align-items: center; gap: .3rem;
      padding: .22rem .75rem; border-radius: 9999px;
      font-size: .72rem; font-weight: 700; white-space: nowrap; cursor: default;
    }}
    .rchip-pass {{ background: #c6f6d5; color: #276749; border: 1px solid #38a169; }}
    .rchip-warn {{ background: #fefcbf; color: #744210; border: 1px solid #d69e2e; }}
    .rchip-fail {{ background: #fed7d7; color: #742a2a; border: 1px solid #e53e3e; }}
    .rchip-skip {{ background: #e2e8f0; color: #4a5568; border: 1px solid #a0aec0; }}

    /* ── Section count badges ── */
    .section-counts {{ display: flex; gap: .35rem; margin-left: auto; flex-shrink: 0; align-items: center; }}
    .scnt {{ padding: 2px 8px; border-radius: 20px; font-size: 11px; font-weight: 700; letter-spacing: .01em; }}
    .scnt-pass {{ background: #d4edda; color: #145a32; }}
    .scnt-warn {{ background: #fff3cd; color: #7d6008; }}
    .scnt-fail {{ background: #f8d7da; color: #7b241c; }}
    .scnt-skip {{ background: #f5f7fa; color: #7f8c8d; }}

    /* ── Test block status left border ── */
    .test-block.tb-fail {{ border-left: 4px solid #e53e3e; }}
    .test-block.tb-warn {{ border-left: 4px solid #d69e2e; }}
    .test-block.tb-pass {{ border-left: 4px solid #38a169; }}

    /* Configuration toggle */
    .btn-config {{
      display: inline-flex; align-items: center; gap: 4px;
      padding: 4px 12px; border-radius: 5px; cursor: pointer;
      font-size: 11px; font-weight: 600; border: none;
      background: #0d6efd; color: #fff;
      transition: background 0.15s; user-select: none;
      margin-bottom: 6px;
    }}
    .btn-config:hover {{ background: #0b5ed7; }}
    .config-panel {{ display: none; }}
    .config-panel.open {{ display: block; }}
    .config-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      margin-top: 6px;
    }}
    .config-table tr:nth-child(even) td {{ background: #f5f7fa; }}
    .config-table tr:nth-child(odd) td {{ background: #ffffff; }}
    .config-table td {{
      padding: 5px 8px;
      border: 1px solid #dde3ec;
      vertical-align: top;
      word-break: break-all;
      color: #333;
    }}
    .config-table td:first-child {{
      font-weight: 700;
      color: #1a3a5c;
      white-space: nowrap;
      width: 45%;
      background: #eef4fb;
    }}
    .config-table td:last-child {{
      font-weight: 400;
      color: #444;
    }}

    @media (max-width: 860px) {{
      .nav-hamburger {{ display: flex; }}
      .top-nav-search-wrap {{ width: 240px; }}
      .top-nav-meta {{ display: none; }}
      .sidebar {{
        position: fixed;
        left: -300px;
        top: 56px;
        height: calc(100vh - 56px);
      }}
      .sidebar.open {{ left: 0; }}
      .main-content {{ width: 100%; }}
    }}
    @media (max-width: 540px) {{
      .top-nav {{ padding: 0 10px; gap: 8px; }}
      .top-nav img {{ height: 28px; }}
      .nav-brand {{ display: none; }}
      .top-nav-search-wrap {{ width: 180px; }}
    }}
  </style>
</head>
<body>

<nav class="top-nav">
  <button class="nav-hamburger" id="hamburgerBtn" onclick="toggleSidebar()" title="Menu">&#9776;</button>
  <img alt="DMTF Redfish Logo" src="data:image/gif;base64,{}"/>
  <div class="nav-brand">
    <div class="top-nav-title">Redfish Use Case Checkers</div>
    <div class="top-nav-sub">Test Report</div>
  </div>
  <div class="top-nav-search-wrap">
    <div class="nav-filter-wrap">
      <span class="fi">&#8260;</span>
      <input id="uriFilter" type="text" placeholder="Filter by Use Case&hellip;" autocomplete="off"/>
      <button id="filterClear" onclick="clearFilter()" title="Clear">&#10005;</button>
    </div>
    <div class="nav-filter-meta">Showing <b id="filterCount">&#8230;</b> test blocks</div>
  </div>
  <div class="top-nav-meta">
    Version: {} &nbsp;|&nbsp; Generated: {}<br/>
    <a href="https://github.com/DMTF/Redfish-Use-Case-Checkers" target="_blank">DMTF/Redfish-Use-Case-Checkers</a>
  </div>
</nav>

<div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>
<div class="app-shell">

  <aside class="sidebar">

    <!-- System Info -->
    <div class="sidebar-section">
      <div class="sidebar-label">System Under Test</div>
      <table class="sys-table">
        <tr><td>Host</td><td>{}</td></tr>
        <tr><td>User</td><td>{}</td></tr>
        <tr><td>Password</td><td>{}</td></tr>
        <tr><td>Product</td><td>{}</td></tr>
        <tr><td>Manufacturer</td><td>{}</td></tr>
        <tr><td>Model</td><td>{}</td></tr>
        <tr><td>Firmware</td><td>{}</td></tr>
      </table>
    </div>

    <!-- Score tiles -->
    <div class="sidebar-section">
      <div class="sidebar-label">Results Summary</div>
      <div class="score-grid">
        <div class="score-tile st-pass"><div class="snum">{}</div><div class="slbl">&#10003; Pass</div></div>
        <div class="score-tile st-warn"><div class="snum">{}</div><div class="slbl">&#9888; Warning</div></div>
        <div class="score-tile st-fail"><div class="snum">{}</div><div class="slbl">&#10007; Fail</div></div>
        <div class="score-tile st-skip"><div class="snum">{}</div><div class="slbl">&#8212; Not Tested</div></div>
      </div>
    </div>

    <!-- Configuration -->
    <div class="sidebar-section">
      <div class="sidebar-label">Configuration</div>
      <span class="btn-config" id="btnConfig"
        onclick="(function(){{var p=document.getElementById('configPanel'),b=document.getElementById('btnConfig');if(p.classList.contains('open')){{p.classList.remove('open');b.innerHTML='&#9881; Show Configuration';}}else{{p.classList.add('open');b.innerHTML='&#9650; Hide Configuration';}}}})()">
        &#9881; Show Configuration
      </span>
      <div class="config-panel" id="configPanel">
        <table class="config-table">
          {}
        </table>
      </div>
    </div>

  </aside>

  <div class="main-content">
    <div class="toolbar">
      <button class="toolbar-btn" onclick="expandAll()">&#9660;&nbsp; Expand All</button>
      <button class="toolbar-btn" onclick="collapseAll()">&#9654;&nbsp; Collapse All</button>
    </div>

    {}

  </div>

</div><!-- /app-shell -->

<button id="scrollTopBtn" title="Back to top"
  onclick="document.querySelector('.main-content').scrollTo({{top:0,behavior:'smooth'}})">
  &#8679;
</button>

<script>
  /* Section collapse/expand */
  document.querySelectorAll('.section-header').forEach(function(hdr) {{
    hdr.addEventListener('click', function() {{
      this.classList.toggle('collapsed');
      var body = this.nextElementSibling;
      body.style.display = (body.style.display === 'none') ? '' : 'none';
    }});
  }});

  /* Test row collapse/expand */
  document.querySelectorAll('.test-heading').forEach(function(hdr) {{
    hdr.addEventListener('click', function() {{
      var body = this.nextElementSibling;
      var arrow = this.querySelector('.test-toggle');
      if (body && body.classList.contains('test-body')) {{
        var hidden = body.classList.toggle('hidden');
        if (arrow) arrow.style.transform = hidden ? 'rotate(-90deg)' : '';
      }}
    }});
  }});

  /* Expand all */
  function expandAll() {{
    document.querySelectorAll('.section-header').forEach(function(hdr) {{
      hdr.classList.remove('collapsed');
      hdr.nextElementSibling.style.display = '';
    }});
    document.querySelectorAll('.test-body').forEach(function(b) {{
      b.classList.remove('hidden');
    }});
    document.querySelectorAll('.test-toggle').forEach(function(a) {{
      a.style.transform = '';
    }});
  }}

  /* Collapse all */
  function collapseAll() {{
    document.querySelectorAll('.section-header').forEach(function(hdr) {{
      hdr.classList.add('collapsed');
      hdr.nextElementSibling.style.display = 'none';
    }});
    document.querySelectorAll('.test-body').forEach(function(b) {{
      b.classList.add('hidden');
    }});
    document.querySelectorAll('.test-toggle').forEach(function(a) {{
      a.style.transform = 'rotate(-90deg)';
    }});
  }}

  /* Scroll-to-top button */
  document.querySelector('.main-content').addEventListener('scroll', function() {{
    document.getElementById('scrollTopBtn').style.display =
      this.scrollTop > 200 ? 'block' : 'none';
  }});

  /* URI filter */
  (function() {{
    var filter = document.getElementById('uriFilter');
    var clearBtn = document.getElementById('filterClear');
    var countEl = document.getElementById('filterCount');
    if (!filter) return;

    function updateCount() {{
      var allTests = document.querySelectorAll('.test-block').length;
      var visibleTests = 0;
      document.querySelectorAll('.test-block').forEach(function(tb) {{
        if (tb.style.display !== 'none') visibleTests += 1;
      }});
      if (countEl) countEl.innerHTML = '<b>' + visibleTests + '</b> / ' + allTests;
      if (clearBtn) clearBtn.style.display = filter.value.trim() ? 'block' : 'none';
    }}

    filter.addEventListener('input', function() {{
      var q = (this.value || '').toLowerCase().trim();
      var sections = document.querySelectorAll('.section');

      sections.forEach(function(section) {{
        var sectionHdr = section.querySelector('.section-header');
        var testBlocks = section.querySelectorAll('.test-block');
        var sectionMatches = false;

        testBlocks.forEach(function(tb) {{
          var heading = tb.querySelector('.test-heading');
          var txt = heading ? heading.textContent.toLowerCase() : '';
          var match = !q || txt.indexOf(q) !== -1;
          tb.style.display = match ? '' : 'none';
          if (match) sectionMatches = true;
        }});

        if (sectionHdr && !sectionMatches && q) {{
          var secTxt = sectionHdr.textContent.toLowerCase();
          if (secTxt.indexOf(q) !== -1) {{
            sectionMatches = true;
            testBlocks.forEach(function(tb) {{ tb.style.display = ''; }});
          }}
        }}

        section.style.display = sectionMatches || !q ? '' : 'none';
      }});

      updateCount();
    }});

    updateCount();
  }})();

  /* Mobile sidebar toggle */
  function toggleSidebar() {{
    var sb = document.querySelector('.sidebar');
    var ov = document.getElementById('sidebarOverlay');
    var open = sb.classList.toggle('open');
    ov.classList.toggle('show', open);
  }}

  function clearFilter() {{
    document.getElementById('uriFilter').value = '';
    document.getElementById('uriFilter').dispatchEvent(new Event('input'));
  }}
</script>

</body>
</html>
"""


def html_report(sut: SystemUnderTest, report_dir, time, tool_version, args=None):
    """
    Creates the HTML report for the system under test

    Args:
        sut: The system under test
        report_dir: The directory for the report
        time: The time the tests finished
        tool_version: The version of the tool

    Returns:
        The path to the HTML report
    """

    file = report_dir / datetime.strftime(time, "RedfishUseCaseCheckersReport_%m_%d_%Y_%H%M%S.html")

    html = ""
    for test_category in sut._results:
        # Aggregate counts for section header badges
        sec_pass = sum(1 for t in test_category["Tests"] for r in t["Results"] if r["Result"] == "PASS")
        sec_warn = sum(1 for t in test_category["Tests"] for r in t["Results"] if r["Result"] == "WARN")
        sec_fail = sum(1 for t in test_category["Tests"] for r in t["Results"] if r["Result"] == "FAIL")
        sec_skip = sum(1 for t in test_category["Tests"] for r in t["Results"] if not r["Result"])
        sec_cnt = '<div class="section-counts">'
        if sec_pass: sec_cnt += '<span class="scnt scnt-pass">&#10003;&nbsp;{}</span>'.format(sec_pass)
        if sec_warn: sec_cnt += '<span class="scnt scnt-warn">&#9888;&nbsp;{}</span>'.format(sec_warn)
        if sec_fail: sec_cnt += '<span class="scnt scnt-fail">&#10007;&nbsp;{}</span>'.format(sec_fail)
        if sec_skip: sec_cnt += '<span class="scnt scnt-skip">&ndash;&nbsp;{}</span>'.format(sec_skip)
        sec_cnt += '</div>'
        html += '<div class="section">'
        html += '<div class="section-header"><span class="section-arrow">&#9660;</span>{}{}</div>'.format(
            html_mod.escape(test_category["Category"]), sec_cnt
        )
        html += '<div class="section-body">'

        for test in test_category["Tests"]:
            # Determine worst result for status left-border
            _rvals = [r["Result"] for r in test["Results"]]
            if "FAIL" in _rvals:
                tb_cls = " tb-fail"
            elif "WARN" in _rvals:
                tb_cls = " tb-warn"
            elif "PASS" in _rvals:
                tb_cls = " tb-pass"
            else:
                tb_cls = ""
            html += '<div class="test-block{}">' .format(tb_cls)
            html += '<div class="test-heading">'
            html += '<div class="test-heading-info">'
            html += '<div class="test-name">{}</div>'.format(html_mod.escape(test["Name"]))
            html += '<div class="test-desc">{}</div>'.format(html_mod.escape(test["Description"]))
            if test.get("Details"):
                html += '<div class="test-detail">{}</div>'.format(html_mod.escape(test["Details"]))
            # Inline result summary chips — aggregate counts (max 4 chips per test)
            t_pass = sum(1 for _r in test["Results"] if _r["Result"] == "PASS")
            t_warn = sum(1 for _r in test["Results"] if _r["Result"] == "WARN")
            t_fail = sum(1 for _r in test["Results"] if _r["Result"] == "FAIL")
            t_skip = sum(1 for _r in test["Results"] if not _r["Result"])
            html += '<div class="result-chips">'
            if t_pass: html += '<span class="rchip rchip-pass">&#10003; {} Pass</span>'.format(t_pass)
            if t_warn: html += '<span class="rchip rchip-warn">&#9888; {} Warn</span>'.format(t_warn)
            if t_fail: html += '<span class="rchip rchip-fail">&#10007; {} Fail</span>'.format(t_fail)
            if t_skip: html += '<span class="rchip rchip-skip">&ndash; {} N/A</span>'.format(t_skip)
            html += '</div>'
            html += '</div>'  # test-heading-info
            html += '<span class="test-toggle">&#9660;</span>'
            html += '</div>'  # test-heading

            html += '<div class="test-body">'
            html += '<table class="test-results-table">'
            html += (
                '<thead><tr>'
                '<th class="col-op">Operation</th>'
                '<th class="col-result">Result</th>'
                '<th class="col-msg">Message</th>'
                '</tr></thead><tbody>'
            )
            for result in test["Results"]:
                if result["Result"] == "PASS":
                    badge = "badge-pass"
                elif result["Result"] == "WARN":
                    badge = "badge-warn"
                elif result["Result"] == "FAIL":
                    badge = "badge-fail"
                else:
                    badge = "badge-skip"
                operation = result["Operation"] if result["Operation"] else "No testing performed"
                html += '<tr><td>{}</td><td><span class="badge {}">{}</span></td><td>{}</td></tr>'.format(
                    html_mod.escape(operation),
                    badge,
                    html_mod.escape(result["Result"]) if result["Result"] else "N/A",
                    html_mod.escape(result["Message"]),
                )
            html += '</tbody></table>'
            html += '</div>'  # test-body
            html += '</div>'  # test-block

        html += '</div>'  # section-body
        html += '</div>'  # section

    config_rows_html = ""
    if args:
        for key, val in sorted(args.items()):
            if val is None:
                val = ""
            elif isinstance(val, list):
                val = " ".join(str(v) for v in val)
            else:
                val = str(val)
            if key in ("password",):
                val = "********" if val else ""
            config_rows_html += "<tr><td>{}</td><td>{}</td></tr>".format(
                html_mod.escape(key), html_mod.escape(val)
            )

    with open(str(file), "w", encoding="utf-8") as fd:
        fd.write(
            html_template.format(
                redfish_logo.logo,
                tool_version,
                time.strftime("%c"),
                sut.rhost,
                sut.username,
                "********",
                sut.product,
                sut.manufacturer,
                sut.model,
                sut.firmware_version,
                sut.pass_count,
                sut.warn_count,
                sut.fail_count,
                sut.skip_count,
                config_rows_html,
                html,
            )
        )
    return file


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _thin_border():
    side = Side(style="thin", color="BFBFBF")
    return Border(left=side, right=side, top=side, bottom=side)


def _cell(ws, row, col, value="", bold=False, color=None, fill_hex=None,
          align="left", wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Segoe UI", size=10, bold=bold, color=color or "1A1A2E")
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = _thin_border()
    if fill_hex:
        c.fill = PatternFill("solid", fgColor=fill_hex)
    return c


def xlsx_report(sut: SystemUnderTest, report_dir, time, tool_version):
    """
    Creates the Excel (xlsx) report for the system under test.

    Args:
        sut: The system under test
        report_dir: The directory for the report
        time: The time the tests finished
        tool_version: The version of the tool

    Returns:
        The path to the xlsx report
    """

    file = report_dir / datetime.strftime(time, "RedfishUseCaseCheckersReport_%m_%d_%Y_%H%M%S.xlsx")

    wb = openpyxl.Workbook()

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"

    HDR_FILL   = "0D1B2A"
    HDR_FONT   = "FFFFFF"
    SUB_FILL   = "1A3A5C"
    ACCENT_FILL= "1565C0"
    META_FILL  = "F0F2F5"
    PASS_FILL  = "D4EDDA"; PASS_FONT  = "145A32"
    WARN_FILL  = "FFF3CD"; WARN_FONT  = "7D6008"
    FAIL_FILL  = "F8D7DA"; FAIL_FONT  = "7B241C"
    SKIP_FILL  = "F5F7FA"; SKIP_FONT  = "7F8C8D"
    CAT_FILL   = "1565C0"
    TEST_FILL  = "F7F9FC"
    SCORE_PASS = "27AE60"
    SCORE_WARN = "D68910"
    SCORE_FAIL = "C0392B"
    SCORE_SKIP = "5D6D7E"

    # Title row  — spans A:D (4 equal columns)
    ws_sum.merge_cells("A1:D1")
    t = ws_sum["A1"]
    t.value = "Redfish Use Case Checkers — Compliance Report"
    t.font = Font(name="Segoe UI", size=14, bold=True, color=HDR_FONT)
    t.fill = PatternFill("solid", fgColor=HDR_FILL)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    # Sub-title
    ws_sum.merge_cells("A2:D2")
    s = ws_sum["A2"]
    s.value = "Version: {}    |    Generated: {}".format(tool_version, time.strftime("%c"))
    s.font = Font(name="Segoe UI", size=10, color=HDR_FONT)
    s.fill = PatternFill("solid", fgColor=SUB_FILL)
    s.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 18

    # Meta info — label in A, value merged B:D
    meta = [
        ("Target System", sut.rhost),
        ("User", sut.username),
        ("Product", sut.product),
        ("Manufacturer", sut.manufacturer),
        ("Model", sut.model),
        ("Firmware", sut.firmware_version),
    ]
    for i, (label, value) in enumerate(meta, start=3):
        _cell(ws_sum, i, 1, label, bold=True, color="6C757D", fill_hex=META_FILL, align="right")
        ws_sum.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        _cell(ws_sum, i, 2, value, fill_hex="FFFFFF")
        # Apply borders to every cell in the merged range so all edges are visible
        for col in range(2, 5):
            ws_sum.cell(row=i, column=col).border = _thin_border()

    # Blank row, then summary counters spanning A:D
    r = 3 + len(meta) + 1
    score_labels = [("✓  PASS", SCORE_PASS), ("⚠  WARN", SCORE_WARN),
                    ("✗  FAIL", SCORE_FAIL), ("–  NOT TESTED", SCORE_SKIP)]
    for col, (label, fill) in enumerate(score_labels, start=1):
        _cell(ws_sum, r, col, label, bold=True, color=HDR_FONT,
              fill_hex=fill, align="center")
        ws_sum.row_dimensions[r].height = 22
    r += 1
    score_data = [
        (sut.pass_count, PASS_FILL, PASS_FONT),
        (sut.warn_count, WARN_FILL, WARN_FONT),
        (sut.fail_count, FAIL_FILL, FAIL_FONT),
        (sut.skip_count, SKIP_FILL, SKIP_FONT),
    ]
    for col, (count, fill, fnt) in enumerate(score_data, start=1):
        _cell(ws_sum, r, col, count, bold=True, color=fnt,
              fill_hex=fill, align="center")
        ws_sum.row_dimensions[r].height = 28

    # Column widths — A (label) narrower, B-D equal data columns
    ws_sum.column_dimensions["A"].width = 20
    for col_letter in ["B", "C", "D"]:
        ws_sum.column_dimensions[col_letter].width = 24

    # Hide grid lines on summary sheet
    ws_sum.sheet_view.showGridLines = False

    # ── Results sheet ────────────────────────────────────────────────────────
    ws = wb.create_sheet("Results")

    # Column definitions: Category | Test Name | Description | Operation | Result | Message
    col_headers = ["Category", "Test Name", "Description", "Operation", "Result", "Message"]
    col_widths   = [22, 28, 40, 40, 12, 55]

    for col, (hdr, width) in enumerate(zip(col_headers, col_widths), start=1):
        _cell(ws, 1, col, hdr, bold=True, color=HDR_FONT,
              fill_hex=SUB_FILL, align="center")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    # Hide grid lines on results sheet
    ws.sheet_view.showGridLines = False

    row = 2
    result_style = {
        "PASS": (PASS_FILL, PASS_FONT),
        "WARN": (WARN_FILL, WARN_FONT),
        "FAIL": (FAIL_FILL, FAIL_FONT),
    }

    for test_category in sut._results:
        category = test_category["Category"]
        for test in test_category["Tests"]:
            test_name = test["Name"]
            description = test["Description"]
            details = test.get("Details", "")
            for result in test["Results"]:
                operation = result["Operation"] or "No testing performed"
                result_val = result["Result"] or "N/A"
                message = result["Message"]

                fill, fnt = result_style.get(result_val, (SKIP_FILL, SKIP_FONT))

                _cell(ws, row, 1, category,    fill_hex=CAT_FILL,  color=HDR_FONT, wrap=True)
                _cell(ws, row, 2, test_name,   fill_hex=TEST_FILL, bold=True, wrap=True)
                _cell(ws, row, 3, description, fill_hex=TEST_FILL, wrap=True)
                _cell(ws, row, 4, operation,   wrap=True)
                _cell(ws, row, 5, result_val,  fill_hex=fill, color=fnt,
                      bold=True, align="center")
                _cell(ws, row, 6, message,     wrap=True)
                ws.row_dimensions[row].height = 30
                row += 1

    wb.save(str(file))
    return file

