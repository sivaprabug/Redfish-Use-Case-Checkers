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
  <title>Redfish Use Case Checkers Report</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    :root {{
      --primary:       #1a3a5c;
      --primary-light: #2c5282;
      --accent:        #3182ce;
      --pass-color:  #276749; --pass-bg:  #c6f6d5; --pass-border:  #38a169;
      --warn-color:  #744210; --warn-bg:  #fefcbf; --warn-border:  #d69e2e;
      --fail-color:  #742a2a; --fail-bg:  #fed7d7; --fail-border:  #e53e3e;
      --skip-color:  #4a5568; --skip-bg:  #e2e8f0; --skip-border:  #a0aec0;
      --border:   #e2e8f0;
      --text:     #2d3748;
      --text-sub: #718096;
      --bg:       #f7fafc;
      --card-bg:  #ffffff;
    }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }}

    /* ── Header ── */
    .header {{
      background: var(--primary);
      color: #fff;
      padding: 0 2rem;
      height: 64px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      box-shadow: 0 2px 8px rgba(0,0,0,.3);
      position: sticky;
      top: 0;
      z-index: 100;
    }}
    .header-brand    {{ display: flex; align-items: center; gap: .9rem; }}
    .header-logo img {{ height: 40px; }}
    .header-title    {{ font-size: 1.05rem; font-weight: 600; letter-spacing: .02em; }}
    .header-subtitle {{ font-size: .73rem; opacity: .75; }}
    .header-meta     {{ text-align: right; font-size: .78rem; opacity: .85; line-height: 1.7; }}
    .header-meta a   {{ color: rgba(255,255,255,.7); font-size: .7rem; text-decoration: none; }}

    /* ── Page layout ── */
    .page-wrap {{ display: flex; align-items: flex-start; min-height: calc(100vh - 64px); }}
    .sidebar {{
      width: 240px;
      flex-shrink: 0;
      background: #ffffff;
      border-right: 1px solid #dde3ec;
      box-shadow: 2px 0 8px rgba(0,0,0,.05);
      padding: 1.4rem 1rem 2rem;
      position: sticky;
      top: 64px;
      height: calc(100vh - 64px);
      overflow-y: auto;
    }}
    .main-content {{ flex: 1; padding: 1.5rem; min-width: 0; }}

    /* ── Sidebar section headings ── */
    .sb-heading {{
      font-size: .6rem;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: .12em;
      color: var(--primary);
      padding: .3rem .5rem;
      background: #eef3fa;
      border-left: 3px solid var(--accent);
      border-radius: 0 4px 4px 0;
      margin-bottom: .65rem;
    }}
    .sb-stats + .sb-heading {{ margin-top: 1.4rem; }}

    /* ── Sidebar info rows ── */
    .sb-row {{
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: .5rem;
      padding: .32rem 0;
      border-bottom: 1px solid #f0f3f7;
      font-size: .78rem;
    }}
    .sb-row:last-of-type {{ border-bottom: none; }}
    .sb-key {{
      color: #64748b;
      font-size: .67rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .05em;
      white-space: nowrap;
      flex-shrink: 0;
    }}
    .sb-val {{
      font-weight: 600;
      color: #1e293b;
      word-break: break-all;
      text-align: right;
      font-size: .78rem;
    }}

    /* ── Sidebar summary stat blocks — 2×2 grid ── */
    .sb-stats {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: .5rem;
      margin-top: .4rem;
    }}
    .sb-stat {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: .7rem .4rem .55rem;
      border-radius: 8px;
      font-weight: 700;
      color: #fff;
      text-align: center;
      min-height: 66px;
    }}
    .sb-stat-count {{ font-size: 1.7rem; line-height: 1; letter-spacing: -.02em; }}
    .sb-stat-label {{ font-size: .58rem; text-transform: uppercase; letter-spacing: .08em; margin-top: .25rem; opacity: .92; line-height: 1.2; }}
    .sb-stat.pass {{ background: #2d9e55; }}
    .sb-stat.warn {{ background: #c98209; }}
    .sb-stat.fail {{ background: #c0392b; }}
    .sb-stat.skip {{ background: #5a6b7d; }}

    /* ── Section ── */
    .section {{ margin-bottom: 1.5rem; }}
    .section-header {{
      background: var(--primary-light);
      color: #fff;
      padding: .75rem 1.25rem;
      border-radius: 8px 8px 0 0;
      font-weight: 600;
      font-size: .88rem;
      letter-spacing: .03em;
      display: flex;
      align-items: center;
      gap: .55rem;
      cursor: pointer;
      user-select: none;
    }}
    .section-arrow {{ font-size: .65rem; transition: transform .2s; }}
    .section-header.collapsed .section-arrow {{ transform: rotate(-90deg); }}
    .section-body {{
      border: 1px solid var(--border);
      border-top: none;
      border-radius: 0 0 8px 8px;
      overflow: hidden;
    }}

    /* ── Test block ── */
    .test-block {{ border-bottom: 1px solid var(--border); }}
    .test-block:last-child {{ border-bottom: none; }}
    .test-heading {{
      background: #f8f9fa;
      padding: .7rem 1.25rem;
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 1rem;
    }}
    .test-heading:hover {{ background: #eef2f7; }}
    .test-name   {{ font-weight: 600; font-size: .87rem; color: var(--primary); }}
    .test-desc   {{ font-size: .79rem; color: var(--text-sub); margin-top: .12rem; }}
    .test-detail {{ font-size: .76rem; color: var(--text-sub); margin-top: .08rem; font-style: italic; }}
    .test-toggle {{ flex-shrink: 0; font-size: .65rem; color: var(--text-sub); margin-top: .3rem; transition: transform .2s; }}
    .test-body         {{ display: block; }}
    .test-body.hidden  {{ display: none; }}

    /* ── Results table ── */
    .test-results-table {{ width: 100%; border-collapse: collapse; font-size: .83rem; }}
    .test-results-table thead tr {{ background: #edf2f7; }}
    .test-results-table th {{
      padding: .45rem 1.25rem;
      text-align: left;
      font-weight: 600;
      font-size: .7rem;
      text-transform: uppercase;
      letter-spacing: .05em;
      color: var(--text-sub);
      border-bottom: 1px solid var(--border);
    }}
    .test-results-table td {{
      padding: .5rem 1.25rem;
      border-bottom: 1px solid #f0f0f0;
      vertical-align: top;
    }}
    .test-results-table tr:last-child td {{ border-bottom: none; }}
    .test-results-table tr:hover td {{ background: #f7fafc; }}
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
    .badge-pass {{ background: var(--pass-bg); color: var(--pass-color); }}
    .badge-warn {{ background: var(--warn-bg); color: var(--warn-color); }}
    .badge-fail {{ background: var(--fail-bg); color: var(--fail-color); }}
    .badge-skip {{ background: var(--skip-bg); color: var(--skip-color); }}

    /* ── Toolbar ── */
    .toolbar {{
      display: flex;
      align-items: center;
      gap: .6rem;
      margin-bottom: 1rem;
    }}
    .toolbar-btn {{
      background: var(--primary-light);
      color: #fff;
      border: none;
      border-radius: 6px;
      padding: .4rem .9rem;
      font-size: .78rem;
      font-weight: 600;
      cursor: pointer;
      letter-spacing: .02em;
      transition: background .15s;
    }}
    .toolbar-btn:hover {{ background: var(--primary); }}

    /* ── Scroll-to-top ── */
    #scroll-top {{
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      width: 42px;
      height: 42px;
      border-radius: 50%;
      background: var(--primary-light);
      color: #fff;
      border: none;
      font-size: 1.1rem;
      cursor: pointer;
      box-shadow: 0 3px 10px rgba(0,0,0,.25);
      display: none;
      align-items: center;
      justify-content: center;
      z-index: 200;
      transition: background .15s, opacity .2s;
    }}
    #scroll-top:hover {{ background: var(--primary); }}
    #scroll-top.visible {{ display: flex; }}

    /* ── Footer ── */
    .footer {{
      text-align: center;
      font-size: .77rem;
      color: var(--text-sub);
      padding: 2rem 1rem;
      border-top: 1px solid var(--border);
      margin-top: 2rem;
    }}
    .footer a {{ color: var(--accent); text-decoration: none; }}
    .footer a:hover {{ text-decoration: underline; }}

    /* ── Inline result summary chips (at most 4 per test) ── */
    .result-chips {{ display: flex; gap: .4rem; margin-top: .45rem; flex-wrap: wrap; }}
    .rchip {{
      display: inline-flex; align-items: center; gap: .3rem;
      padding: .22rem .75rem; border-radius: 9999px;
      font-size: .72rem; font-weight: 700; white-space: nowrap; cursor: default;
    }}
    .rchip-pass {{ background: var(--pass-bg); color: var(--pass-color); border: 1px solid var(--pass-border); }}
    .rchip-warn {{ background: var(--warn-bg); color: var(--warn-color); border: 1px solid var(--warn-border); }}
    .rchip-fail {{ background: var(--fail-bg); color: var(--fail-color); border: 1px solid var(--fail-border); }}
    .rchip-skip {{ background: var(--skip-bg); color: var(--skip-color); border: 1px solid var(--skip-border); }}

    /* ── Section count badges ── */
    .section-counts {{ display: flex; gap: .35rem; margin-left: auto; flex-shrink: 0; align-items: center; }}
    .scnt {{ padding: .12rem .58rem; border-radius: 9999px; font-size: .7rem; font-weight: 700; letter-spacing: .03em; }}
    .scnt-pass {{ background: rgba(56,161,105,.35);   color: #c6f6d5; }}
    .scnt-warn {{ background: rgba(214,158,46,.35);  color: #fefcbf; }}
    .scnt-fail {{ background: rgba(229,62,62,.35);    color: #fed7d7; }}
    .scnt-skip {{ background: rgba(160,174,192,.3);  color: #e2e8f0; }}

    /* ── Test block status left border ── */
    .test-block.tb-fail {{ border-left: 4px solid var(--fail-border); }}
    .test-block.tb-warn {{ border-left: 4px solid var(--warn-border); }}
    .test-block.tb-pass {{ border-left: 4px solid var(--pass-border); }}

    /* ── Sidebar progress bar ── */
    .sb-pass-rate {{
      text-align: center;
      margin: .5rem 0 .65rem;
    }}
    .sb-pass-rate-num {{
      font-size: 1.9rem;
      font-weight: 800;
      color: #2d9e55;
      letter-spacing: -.03em;
      line-height: 1;
    }}
    .sb-pass-rate-sub {{
      font-size: .62rem;
      color: #64748b;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .07em;
      margin-top: .18rem;
    }}
    .sb-prog-bar {{
      height: 12px;
      border-radius: 6px;
      overflow: hidden;
      display: flex;
      background: #f0f3f7;
      margin-bottom: .3rem;
    }}
    .sb-prog-seg {{
      height: 100%;
      width: 0%;
      transition: width 1.2s cubic-bezier(.4,0,.2,1);
    }}
    .sb-prog-seg.pass {{ background: #2d9e55; }}
    .sb-prog-seg.warn {{ background: #c98209; }}
    .sb-prog-seg.fail {{ background: #c0392b; }}
    .sb-prog-seg.skip {{ background: #cbd5e1; }}
    .sb-prog-labels {{ display: flex; }}
    .sb-prog-lbl {{
      font-size: .62rem;
      font-weight: 700;
      text-align: center;
      overflow: hidden;
      white-space: nowrap;
      width: 0%;
      transition: width 1.2s cubic-bezier(.4,0,.2,1);
    }}
    .sb-prog-lbl.pass {{ color: #2d9e55; }}
    .sb-prog-lbl.warn {{ color: #c98209; }}
    .sb-prog-lbl.fail {{ color: #c0392b; }}
    .sb-prog-lbl.skip {{ color: #94a3b8; }}
  </style>
</head>
<body>

<header class="header">
  <div class="header-brand">
    <div class="header-logo">
      <img src="data:image/gif;base64,{}" alt="DMTF Redfish Logo">
    </div>
    <div>
      <div class="header-title">Redfish Use Case Checkers</div>
      <div class="header-subtitle">Automated Compliance Report &mdash; v{}</div>
    </div>
  </div>
  <div class="header-meta">
    Generated: {}<br>
    <a href="https://github.com/DMTF/Redfish-Use-Case-Checkers">github.com/DMTF/Redfish-Use-Case-Checkers</a>
  </div>
</header>

<div class="page-wrap">

  <aside class="sidebar">
    <div class="sb-heading">System Under Test</div>
    <div class="sb-row"><span class="sb-key">Host</span><span class="sb-val">{}/redfish/v1/</span></div>
    <div class="sb-row"><span class="sb-key">User</span><span class="sb-val"><strong>{}</strong></span></div>
    <div class="sb-row"><span class="sb-key">Password</span><span class="sb-val">{}</span></div>
    <div class="sb-row"><span class="sb-key">Product</span><span class="sb-val">{}</span></div>
    <div class="sb-row"><span class="sb-key">Manufacturer</span><span class="sb-val">{}</span></div>
    <div class="sb-row"><span class="sb-key">Model</span><span class="sb-val">{}</span></div>
    <div class="sb-row"><span class="sb-key">Firmware</span><span class="sb-val">{}</span></div>
    <div class="sb-heading" style="margin-top:1.25rem">Results Summary</div>
    <div class="sb-stats">
      <div class="sb-stat pass">
        <span class="sb-stat-count">{}</span>
        <span class="sb-stat-label">&#10003;<br>Pass</span>
      </div>
      <div class="sb-stat warn">
        <span class="sb-stat-count">{}</span>
        <span class="sb-stat-label">&#9888;<br>Warning</span>
      </div>
      <div class="sb-stat fail">
        <span class="sb-stat-count">{}</span>
        <span class="sb-stat-label">&#10007;<br>Fail</span>
      </div>
      <div class="sb-stat skip">
        <span class="sb-stat-count">{}</span>
        <span class="sb-stat-label">&ndash;<br>Not Tested</span>
      </div>
    </div>

    <div class="sb-heading" style="margin-top:1.25rem">Pass Rate</div>
    <div class="sb-pass-rate">
      <div class="sb-pass-rate-num" id="pass-rate-pct">—</div>
      <div class="sb-pass-rate-sub">of all checks passed</div>
    </div>
    <div class="sb-prog-bar">
      <div class="sb-prog-seg pass" id="pseg-pass"></div>
      <div class="sb-prog-seg warn" id="pseg-warn"></div>
      <div class="sb-prog-seg fail" id="pseg-fail"></div>
      <div class="sb-prog-seg skip" id="pseg-skip"></div>
    </div>
    <div class="sb-prog-labels">
      <div class="sb-prog-lbl pass" id="plbl-pass"></div>
      <div class="sb-prog-lbl warn" id="plbl-warn"></div>
      <div class="sb-prog-lbl fail" id="plbl-fail"></div>
      <div class="sb-prog-lbl skip" id="plbl-skip"></div>
    </div>
  </aside>

  <div class="main-content">
    <div class="toolbar">
      <button class="toolbar-btn" onclick="expandAll()">&#9660;&nbsp; Expand All</button>
      <button class="toolbar-btn" onclick="collapseAll()">&#9654;&nbsp; Collapse All</button>
    </div>

    {}

  </div>

</div>

<button id="scroll-top" title="Back to top" onclick="window.scrollTo({{top:0,behavior:'smooth'}})">
  &#8679;
</button>

<footer class="footer">
  This report was generated by the
  <a href="https://github.com/DMTF/Redfish-Use-Case-Checkers">DMTF Redfish Use Case Checkers</a>.
  For feedback, please
  <a href="https://github.com/DMTF/Redfish-Use-Case-Checkers/issues">open an issue</a>.
</footer>

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
  var scrollBtn = document.getElementById('scroll-top');
  window.addEventListener('scroll', function() {{
    if (window.scrollY > 400) {{
      scrollBtn.classList.add('visible');
    }} else {{
      scrollBtn.classList.remove('visible');
    }}
  }});

  /* ── Pass-rate progress bar animation ── */
  (function() {{
    var keys = ['pass', 'warn', 'fail', 'skip'];
    var counts = {{}};
    var total = 0;
    keys.forEach(function(k) {{
      var el = document.querySelector('.sb-stat.' + k + ' .sb-stat-count');
      counts[k] = el ? (parseInt(el.textContent, 10) || 0) : 0;
      total += counts[k];
    }});
    if (!total) return;

    /* Show overall pass-rate number immediately */
    var pctPass = (counts.pass / total * 100);
    var pctEl = document.getElementById('pass-rate-pct');
    if (pctEl) pctEl.textContent = pctPass.toFixed(1) + '%';

    /* Animate bars after a short paint delay */
    setTimeout(function() {{
      keys.forEach(function(k) {{
        var pct = (counts[k] / total * 100);
        var seg = document.getElementById('pseg-' + k);
        var lbl = document.getElementById('plbl-' + k);
        if (seg) seg.style.width = pct.toFixed(3) + '%';
        if (lbl) {{
          lbl.style.width = pct.toFixed(3) + '%';
          lbl.textContent = counts[k] > 0 ? pct.toFixed(1) + '%' : '';
        }}
      }});
    }}, 120);
  }})();
</script>

</body>
</html>
"""


def html_report(sut: SystemUnderTest, report_dir, time, tool_version):
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
    c.font = Font(name="Calibri", size=10, bold=bold, color=color or "000000")
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

    HDR_FILL   = "1A3A5C"
    HDR_FONT   = "FFFFFF"
    SUB_FILL   = "2C5282"
    META_FILL  = "EDF2F7"
    PASS_FILL  = "C6F6D5"; PASS_FONT  = "276749"
    WARN_FILL  = "FEFCBF"; WARN_FONT  = "744210"
    FAIL_FILL  = "FED7D7"; FAIL_FONT  = "742A2A"
    SKIP_FILL  = "E2E8F0"; SKIP_FONT  = "4A5568"
    CAT_FILL   = "2C5282"
    TEST_FILL  = "EDF2F7"

    # Title row  — spans A:D (4 equal columns)
    ws_sum.merge_cells("A1:D1")
    t = ws_sum["A1"]
    t.value = "Redfish Use Case Checkers — Compliance Report"
    t.font = Font(name="Calibri", size=14, bold=True, color=HDR_FONT)
    t.fill = PatternFill("solid", fgColor=HDR_FILL)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 30

    # Sub-title
    ws_sum.merge_cells("A2:D2")
    s = ws_sum["A2"]
    s.value = "Tool Version: {}    |    Generated: {}".format(tool_version, time.strftime("%c"))
    s.font = Font(name="Calibri", size=10, color=HDR_FONT)
    s.fill = PatternFill("solid", fgColor=SUB_FILL)
    s.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 18

    # Meta info — label in A, value merged B:D
    meta = [
        ("Target System", "{}/redfish/v1/".format(sut.rhost)),
        ("User", sut.username),
        ("Product", sut.product),
        ("Manufacturer", sut.manufacturer),
        ("Model", sut.model),
        ("Firmware", sut.firmware_version),
    ]
    for i, (label, value) in enumerate(meta, start=3):
        _cell(ws_sum, i, 1, label, bold=True, fill_hex=META_FILL, align="right")
        ws_sum.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        _cell(ws_sum, i, 2, value, fill_hex="FFFFFF")
        # Apply borders to every cell in the merged range so all edges are visible
        for col in range(2, 5):
            ws_sum.cell(row=i, column=col).border = _thin_border()

    # Blank row, then summary counters spanning A:D
    r = 3 + len(meta) + 1
    for col, label in enumerate(["PASS", "WARN", "FAIL", "NOT TESTED"], start=1):
        _cell(ws_sum, r, col, label, bold=True, color=HDR_FONT,
              fill_hex=HDR_FILL, align="center")
    r += 1
    fills  = [PASS_FILL,  WARN_FILL,  FAIL_FILL,  SKIP_FILL]
    fonts  = [PASS_FONT,  WARN_FONT,  FAIL_FONT,  SKIP_FONT]
    counts = [sut.pass_count, sut.warn_count, sut.fail_count, sut.skip_count]
    for col, (fill, fnt, count) in enumerate(zip(fills, fonts, counts), start=1):
        _cell(ws_sum, r, col, count, bold=True, color=fnt,
              fill_hex=fill, align="center")

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
              fill_hex=HDR_FILL, align="center")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 20
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

