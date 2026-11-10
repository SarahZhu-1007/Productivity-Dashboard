[Uploading suzhou_efficiency_dashboard_v10 (1).html…]()
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>宜家苏州商场人员效率看板</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- ECharts -->
  <script src="https://g.alicdn.com/code/lib/echarts/5.4.3/echarts.min.js"></script>
  <!-- SheetJS for local Excel/CSV reading & writing -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif; }
    .table-input { width: 100%; min-width: 90px; border: 1px solid #cbd5e1; border-radius: .5rem; padding: .45rem .6rem; outline: none; background: #fff; transition: all .2s; }
    .table-input:focus { border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,.15); }
    .fixed-cell { background: #f8fafc; color: #334155; }
    .badge { display: inline-flex; align-items: center; padding: 0.25rem 0.65rem; border-radius: 9999px; font-size: 0.825rem; font-weight: 600; }
    .nav-tab { transition: all 0.2s ease-in-out; }
    .nav-tab.active { background-color: #ffffff; color: #1e3a8a; font-weight: 700; shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-bottom: 3px solid #f59e0b; }
    .nav-tab:not(.active) { color: #e0e7ff; background-color: rgba(255, 255, 255, 0.1); }
    .nav-tab:not(.active):hover { background-color: rgba(255, 255, 255, 0.2); color: #ffffff; }
  </style>
</head>
<body class="min-h-screen bg-slate-100 text-slate-800 flex flex-col justify-between">

  <div>
    <!-- Top Header & IKEA Brand Banner -->
    <header class="bg-gradient-to-r from-blue-900 via-blue-800 to-indigo-900 text-white shadow-xl">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 pt-6 pb-2">
        <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4 pb-4">
          <div>
            <div class="inline-flex items-center gap-2 rounded-full bg-white/15 px-3 py-1 text-xs font-semibold tracking-wider uppercase mb-2">
              <span class="w-2 h-2 rounded-full bg-amber-400 animate-pulse"></span>
              IKEA SUZHOU · Store Operations Efficiency System
            </div>
            <h1 class="text-3xl md:text-4xl font-extrabold tracking-tight flex items-center gap-3">
              <span>宜家苏州商场人员效率看板</span>
            </h1>
            <p class="mt-1 text-blue-200 text-xs sm:text-sm">支持双表格独立拆分上传（工时/关键驱动）、动态日历汇总与效率指标分析</p>
          </div>
          <div class="flex items-center gap-3 bg-white/10 px-4 py-2.5 rounded-2xl backdrop-blur text-xs">
            <span class="text-blue-200">当前商场：</span>
            <span class="font-bold text-amber-300">苏州商场 (484 - Suzhou Store)</span>
          </div>
        </div>

        <!-- Pagination / Module Tabs Navigation -->
        <nav class="flex flex-wrap items-center gap-2 pt-2 border-t border-white/10" aria-label="Modules">
          <button id="tab-btn-module1" onclick="switchPage('module1')" class="nav-tab active px-5 py-3 rounded-t-xl text-sm flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
            <span>模块 1：员工工时总体情况</span>
          </button>
          <button id="tab-btn-module2" onclick="switchPage('module2')" class="nav-tab px-5 py-3 rounded-t-xl text-sm flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z"/></svg>
            <span>模块 2：人员效率</span>
          </button>
          <button id="tab-btn-module3" onclick="switchPage('module3')" class="nav-tab px-5 py-3 rounded-t-xl text-sm flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
            <span>模块 3：周内工时占比 😊</span>
          </button>
          <button id="tab-btn-module4" onclick="switchPage('module4')" class="nav-tab px-5 py-3 rounded-t-xl text-sm flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/></svg>
            <span>模块 4：待配置</span>
          </button>
        </nav>
      </div>
    </header>

    <!-- Main Content Area: Each Module displays as a separate page view -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 py-8">

      <!-- ======================================================== -->
      <!-- PAGE 1: 模块 1 员工工时总体情况                           -->
      <!-- ======================================================== -->
      <div id="page-module1" class="space-y-8">

        <!-- Breadcrumb / Header -->
        <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md bg-blue-100 text-blue-800 text-xs font-bold">模块 1</span>
              <h2 class="text-xl sm:text-2xl font-bold text-slate-900">员工工时总体情况</h2>
            </div>
            <p class="text-xs sm:text-sm text-slate-500 mt-1">包含月度汇总与年度工时对比，输入数据实时暂存，公式自动计算完成率与信号灯。</p>
          </div>
          <div class="text-xs text-slate-400 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200">
            切换上方分页标签可浏览不同模块，当前数据已保存在本地
          </div>
        </div>

        <!-- Part A: 月度汇总表 -->
        <div class="rounded-3xl bg-white shadow-md border border-slate-200 overflow-hidden">
          <div class="p-5 sm:p-6 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 bg-slate-50/70">
            <div>
              <h3 class="text-lg sm:text-xl font-bold text-slate-900 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
                月度汇总表
              </h3>
              <p class="text-slate-500 text-xs mt-1">可选择 2026年9月 至 2027年8月，空白内容可在线填写与修改，各月独立保存。</p>
            </div>
            <div class="flex items-center gap-3">
              <label class="text-xs font-bold text-slate-700 whitespace-nowrap">选择月份：</label>
              <select id="monthSelect" class="table-input w-48 text-xs font-bold" onchange="switchMonth()"></select>
            </div>
          </div>

          <div class="p-5 sm:p-6">
            <div class="mb-4 flex flex-wrap items-center gap-3">
              <button onclick="saveMonthlyData(true)" class="rounded-xl bg-blue-600 px-4 py-2 text-white font-bold text-xs shadow hover:bg-blue-700 transition flex items-center gap-1.5">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>
                保存本月数据
              </button>
              <button onclick="clearMonthlyData()" class="rounded-xl bg-slate-200 px-4 py-2 text-slate-700 font-bold text-xs hover:bg-slate-300 transition">
                清空本月数据
              </button>
              <span id="monthlySaveStatus" class="text-xs text-slate-500 font-medium">请选择月份并填写数据。</span>
            </div>

            <div class="overflow-x-auto rounded-xl border border-slate-200">
              <table class="w-full border-collapse text-sm">
                <thead>
                  <tr class="bg-slate-900 text-white text-xs sm:text-sm">
                    <th class="px-4 py-3 text-left font-semibold">项目</th>
                    <th class="px-4 py-3 text-center font-semibold">月度销售目标</th>
                    <th class="px-4 py-3 text-center font-semibold">月度实际销售</th>
                    <th class="px-4 py-3 text-center font-semibold text-blue-300">完成率</th>
                    <th class="px-4 py-3 text-center font-semibold">月度工时目标</th>
                    <th class="px-4 py-3 text-center font-semibold">月度实际工时</th>
                    <th class="px-4 py-3 text-center font-semibold text-indigo-300">完成率</th>
                  </tr>
                </thead>
                <tbody id="monthlyBody" class="divide-y divide-slate-200 bg-white"></tbody>
              </table>
            </div>
            <div class="mt-3 text-xs text-slate-500">
              <strong>计算逻辑：</strong>月度销售完成率 = 月度实际销售 ÷ 月度销售目标；月度工时使用率 = 月度实际工时 ÷ 月度工时目标。
            </div>
          </div>
        </div>

        <!-- Part B: 年度员工工时总体情况 -->
        <div class="rounded-3xl bg-white shadow-md border border-slate-200 overflow-hidden">
          <div class="p-5 sm:p-6 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 bg-slate-50/70">
            <div>
              <h3 class="text-lg sm:text-xl font-bold text-slate-900 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-indigo-600"></span>
                年度工时与销售总体情况
              </h3>
              <p class="text-slate-500 text-xs mt-1">固定年度目标，“更新时间”、“实际销售”、“实际工时”保留空白可填。</p>
            </div>
            <div class="flex items-center gap-3">
              <label class="text-xs font-bold text-slate-700 whitespace-nowrap">更新时间：</label>
              <input id="updateTime" class="table-input w-52 text-xs" placeholder="请输入更新时间" oninput="autoSave()" />
            </div>
          </div>

          <div class="p-5 sm:p-6">
            <div class="mb-4 flex flex-wrap items-center gap-3">
              <button onclick="saveData(true)" class="rounded-xl bg-blue-600 px-4 py-2 text-white font-bold text-xs shadow hover:bg-blue-700 transition flex items-center gap-1.5">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>
                保存输入数据
              </button>
              <button onclick="clearSavedData()" class="rounded-xl bg-slate-200 px-4 py-2 text-slate-700 font-bold text-xs hover:bg-slate-300 transition">
                清空已保存数据
              </button>
              <span id="saveStatus" class="text-xs text-slate-500 font-medium">数据会自动暂存，也可手动保存。</span>
            </div>

            <div class="overflow-x-auto rounded-xl border border-slate-200">
              <table class="w-full border-collapse text-sm">
                <thead>
                  <tr class="bg-slate-900 text-white text-xs sm:text-sm">
                    <th class="px-4 py-3 text-left font-semibold">项目</th>
                    <th class="px-4 py-3 text-right font-semibold">销售目标</th>
                    <th class="px-4 py-3 text-center font-semibold">实际销售</th>
                    <th class="px-4 py-3 text-center font-semibold text-blue-300">完成率</th>
                    <th class="px-4 py-3 text-right font-semibold">年度工时目标</th>
                    <th class="px-4 py-3 text-center font-semibold">实际工时</th>
                    <th class="px-4 py-3 text-center font-semibold text-indigo-300">完成率</th>
                  </tr>
                </thead>
                <tbody id="dataBody" class="divide-y divide-slate-200 bg-white"></tbody>
              </table>
            </div>

            <!-- Visual Charts: 2 Compact Bar Charts -->
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-6">
              <div class="rounded-2xl bg-slate-50 border border-slate-200 p-5">
                <div class="flex items-start justify-between gap-3 mb-2">
                  <div>
                    <h4 class="text-sm sm:text-base font-bold text-slate-900">线下（商场）销售完成率 vs 工时使用率</h4>
                    <p class="text-xs text-slate-500 mt-0.5">工时使用率 ÷ 销售完成率（&lt;102%绿灯, 102%-108%黄灯, &gt;108%红灯）</p>
                  </div>
                  <span id="badge-0" class="shrink-0 rounded-full px-2.5 py-1 text-xs font-bold bg-slate-100 text-slate-600">待输入</span>
                </div>
                <div id="chart0" class="h-52 w-full"></div>
              </div>

              <div class="rounded-2xl bg-slate-50 border border-slate-200 p-5">
                <div class="flex items-start justify-between gap-3 mb-2">
                  <div>
                    <h4 class="text-sm sm:text-base font-bold text-slate-900">餐饮销售完成率 vs 工时使用率</h4>
                    <p class="text-xs text-slate-500 mt-0.5">工时使用率 ÷ 销售完成率（&lt;102%绿灯, 102%-108%黄灯, &gt;108%红灯）</p>
                  </div>
                  <span id="badge-1" class="shrink-0 rounded-full px-2.5 py-1 text-xs font-bold bg-slate-100 text-slate-600">待输入</span>
                </div>
                <div id="chart1" class="h-52 w-full"></div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- ======================================================== -->
      <!-- PAGE 2: 模块 2 人员效率（支持双表独立拆分上传识别）          -->
      <!-- ======================================================== -->
      <div id="page-module2" class="space-y-8 hidden">

        <!-- Breadcrumb / Header -->
        <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md bg-indigo-100 text-indigo-800 text-xs font-bold">模块 2</span>
              <h2 class="text-xl sm:text-2xl font-bold text-slate-900">人员效率</h2>
            </div>
            <p class="text-xs sm:text-sm text-slate-500 mt-1">
              拆分支持【工时表】与【关键驱动指标表】独立上传，系统自动按日期关联并累加，公式计算实际效率。
            </p>
          </div>
          <div class="text-xs text-slate-500 bg-indigo-50/70 border border-indigo-100 px-3 py-1.5 rounded-lg">
            计算公式：<span class="font-bold text-indigo-700">实际效率 = 关键驱动指标 ÷ 工时</span>
          </div>
        </div>

        <div class="bg-white rounded-3xl p-6 shadow-md border border-slate-200 space-y-6">

          <!-- DUAL UPLOAD PANELS: 工时表格 & 关键驱动指标表格 -->
          <div class="grid grid-cols-1 md:grid-cols-2 gap-5">

            <!-- Panel 1: 人员效率工时.xlsx 上传 -->
            <div class="bg-gradient-to-br from-blue-50/70 to-slate-50 border-2 border-blue-200/80 rounded-2xl p-5 flex flex-col justify-between space-y-4">
              <div class="flex items-start justify-between gap-3">
                <div class="flex items-center gap-3">
                  <div class="w-11 h-11 rounded-xl bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                    <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
                  </div>
                  <div>
                    <div class="font-bold text-slate-900 text-base">① 上传【工时表格】</div>
                    <div class="text-xs text-slate-500 mt-0.5">识别文件如：<code>人员效率工时.xlsx</code></div>
                  </div>
                </div>
                <span id="hoursTag" class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-blue-100 text-blue-800">已就绪</span>
              </div>

              <div class="bg-white/80 p-3 rounded-xl border border-blue-100 text-xs text-slate-600 flex items-center justify-between">
                <span id="hoursStatus">已加载示范数据（包含商场、销售、餐饮等各部门）</span>
                <span id="hoursDaysCount" class="font-bold text-blue-700"></span>
              </div>

              <div class="flex items-center gap-2 pt-1">
                <input type="file" id="hoursFileInput" accept=".xlsx,.xls,.csv" class="hidden" onchange="handleHoursUpload(event)" />
                <button onclick="document.getElementById('hoursFileInput').click()" class="flex-1 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-semibold text-xs rounded-xl shadow transition flex items-center justify-center gap-1.5">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
                  上传工时表格
                </button>
                <button onclick="downloadHoursTemplate()" class="px-3.5 py-2.5 bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 font-semibold text-xs rounded-xl shadow-sm transition flex items-center justify-center gap-1.5">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                  下载工时模板
                </button>
              </div>
            </div>

            <!-- Panel 2: 人员效率关键驱动.xlsx 上传 -->
            <div class="bg-gradient-to-br from-indigo-50/70 to-slate-50 border-2 border-indigo-200/80 rounded-2xl p-5 flex flex-col justify-between space-y-4">
              <div class="flex items-start justify-between gap-3">
                <div class="flex items-center gap-3">
                  <div class="w-11 h-11 rounded-xl bg-indigo-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                    <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"/></svg>
                  </div>
                  <div>
                    <div class="font-bold text-slate-900 text-base">② 上传【关键驱动指标表格】</div>
                    <div class="text-xs text-slate-500 mt-0.5">识别文件如：<code>人员效率关键驱动.xlsx</code></div>
                  </div>
                </div>
                <span id="driverTag" class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-indigo-100 text-indigo-800">已就绪</span>
              </div>

              <div class="bg-white/80 p-3 rounded-xl border border-indigo-100 text-xs text-slate-600 flex items-center justify-between">
                <span id="driverStatus">已加载示范数据（包含商场、销售、餐饮等各部门）</span>
                <span id="driverDaysCount" class="font-bold text-indigo-700"></span>
              </div>

              <div class="flex items-center gap-2 pt-1">
                <input type="file" id="driverFileInput" accept=".xlsx,.xls,.csv" class="hidden" onchange="handleDriverUpload(event)" /> <button onclick="clearAllEfficiencyData()" class="px-3 py-2.5 bg-rose-50 border border-rose-200 text-rose-600 hover:bg-rose-100 font-semibold text-xs rounded-xl shadow-sm transition">清空数据</button>
                <button onclick="document.getElementById('driverFileInput').click()" class="flex-1 px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs rounded-xl shadow transition flex items-center justify-center gap-1.5">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
                  上传关键驱动指标
                </button>
                <button onclick="downloadDriverTemplate()" class="px-3.5 py-2.5 bg-white hover:bg-indigo-50 text-indigo-700 border border-indigo-200 font-semibold text-xs rounded-xl shadow-sm transition flex items-center justify-center gap-1.5">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                  下载驱动模板
                </button>
              </div>
            </div>

          </div>

          <!-- Calendar Date Filter -->
          <div class="bg-slate-50 rounded-2xl p-4 sm:p-5 border border-slate-200 space-y-3">
            <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
              <span class="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-2">
                <svg class="w-4 h-4 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
                日历筛选（选择统计时间段，数据自动按区间累加计算）
              </span>
              <div class="flex items-center gap-1.5">
                <button onclick="setDateRangeQuick('all')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">全量区间</button>
                <button onclick="setDateRangeQuick('latest30')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">近30天</button>
                <button onclick="setDateRangeQuick('latest7')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">近7天</button>
                <button onclick="resetModule2Defaults()" class="px-2.5 py-1 text-xs font-semibold bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition" title="重置为示范数据">重置</button>
              </div>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-12 gap-3 items-center">
              <div class="lg:col-span-4">
                <label class="text-[11px] text-slate-500 font-medium block mb-1">开始日期</label>
                <input type="date" id="effStartDate" class="table-input py-1.5 text-xs" onchange="applyDateFilter()" />
              </div>
              <div class="lg:col-span-4">
                <label class="text-[11px] text-slate-500 font-medium block mb-1">结束日期</label>
                <input type="date" id="effEndDate" class="table-input py-1.5 text-xs" onchange="applyDateFilter()" />
              </div>
              <div class="lg:col-span-4 self-end">
                <div class="bg-indigo-50/70 border border-indigo-100 rounded-xl px-3 py-2 text-xs text-indigo-950 font-medium" id="effStatusText">
                  统计中...
                </div>
              </div>
            </div>
          </div>

          <!-- Module 2 Main Table -->
          <div class="overflow-x-auto rounded-2xl border border-slate-200 shadow-sm">
            <table class="w-full border-collapse text-sm">
              <thead>
                <tr class="bg-indigo-950 text-white text-xs sm:text-sm">
                  <th class="px-5 py-3.5 text-left font-bold">部门</th>
                  <th class="px-5 py-3.5 text-center font-bold">效率目标</th>
                  <th class="px-5 py-3.5 text-right font-bold text-indigo-200">关键驱动指标 (时间段累加)</th>
                  <th class="px-5 py-3.5 text-right font-bold text-blue-200">工时 (时间段累加)</th>
                  <th class="px-5 py-3.5 text-right font-bold text-amber-300">实际效率</th>
                  <th class="px-5 py-3.5 text-center font-bold">目标对比 / 达成率</th>
                </tr>
              </thead>
              <tbody id="effTableBody" class="divide-y divide-slate-200 bg-white">
                <!-- Dynamically populated -->
              </tbody>
            </table>
          </div>

          <!-- ======================================================== -->
          <!-- MODULE 2: 多时段效率比对分析 (数值比对 + 曲线图)            -->
          <!-- ======================================================== -->
          <div class="mt-8 pt-6 border-t-2 border-indigo-100 space-y-6">

            <!-- Comparison Header & Department Dropdown -->
            <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4 bg-indigo-50/60 p-5 rounded-2xl border border-indigo-100">
              <div>
                <div class="flex items-center gap-2">
                  <span class="w-3 h-3 rounded-full bg-indigo-600 animate-pulse"></span>
                  <h3 class="text-lg sm:text-xl font-bold text-slate-900">人员效率多时段对比分析</h3>
                </div>
                <p class="text-xs sm:text-sm text-slate-500 mt-1">
                  选择任意部门与两个不同时间段，对比综合效率变动数值及每日走势曲线（公式：效率 = 关键驱动指标 ÷ 工时）
                </p>
              </div>

              <!-- Department Select -->
              <div class="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-indigo-200 shadow-sm">
                <label class="text-xs font-bold text-slate-700 whitespace-nowrap">对比部门：</label>
                <select id="compareDeptSelect" class="table-input w-40 text-xs font-bold py-1 bg-white border-0 focus:ring-0" onchange="updateEfficiencyComparison()"><option value="商场OMNI" selected>商场OMNI</option>
<option value="销售部(线下)">销售部(线下)</option>
<option value="顾客关系部(线下)">顾客关系部(线下)</option>
<option value="物流部">物流部</option>
<option value="餐饮部">餐饮部</option></select>
              </div>
            </div>

            <!-- Two Time Period Selectors (Side-by-Side Cards) -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <!-- Period 1 -->
              <div class="bg-blue-50/60 border border-blue-200 rounded-2xl p-4 space-y-3">
                <div class="flex items-center justify-between">
                  <span class="text-xs font-bold text-blue-900 flex items-center gap-1.5">
                    <span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
                    时间段 1 (基准期)
                  </span>
                  <span id="period1DaysTag" class="text-[11px] font-semibold text-blue-700 bg-blue-100 px-2.5 py-0.5 rounded-md">-- 天数据</span>
                </div>
                <div class="grid grid-cols-2 gap-2">
                  <div>
                    <label class="text-[11px] text-slate-500 font-medium block mb-1">起始日期 1</label>
                    <input type="date" id="compareStart1" class="table-input py-1.5 text-xs bg-white" onchange="updateEfficiencyComparison()" />
                  </div>
                  <div>
                    <label class="text-[11px] text-slate-500 font-medium block mb-1">截止日期 1</label>
                    <input type="date" id="compareEnd1" class="table-input py-1.5 text-xs bg-white" onchange="updateEfficiencyComparison()" />
                  </div>
                </div>
              </div>

              <!-- Period 2 -->
              <div class="bg-amber-50/60 border border-amber-200 rounded-2xl p-4 space-y-3">
                <div class="flex items-center justify-between">
                  <span class="text-xs font-bold text-amber-900 flex items-center gap-1.5">
                    <span class="w-2.5 h-2.5 rounded-full bg-amber-600"></span>
                    时间段 2 (对比期)
                  </span>
                  <span id="period2DaysTag" class="text-[11px] font-semibold text-amber-800 bg-amber-100 px-2.5 py-0.5 rounded-md">-- 天数据</span>
                </div>
                <div class="grid grid-cols-2 gap-2">
                  <div>
                    <label class="text-[11px] text-slate-500 font-medium block mb-1">起始日期 2</label>
                    <input type="date" id="compareStart2" class="table-input py-1.5 text-xs bg-white" onchange="updateEfficiencyComparison()" />
                  </div>
                  <div>
                    <label class="text-[11px] text-slate-500 font-medium block mb-1">截止日期 2</label>
                    <input type="date" id="compareEnd2" class="table-input py-1.5 text-xs bg-white" onchange="updateEfficiencyComparison()" />
                  </div>
                </div>
              </div>
            </div>

            <!-- Numeric Comparison Stat Cards (数值比对) -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <!-- Period 1 Eff -->
              <div class="bg-white p-4 rounded-2xl border border-blue-200 shadow-sm">
                <div class="text-xs text-blue-600 font-semibold flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-blue-600"></span>
                  时间段 1 综合效率
                </div>
                <div id="compEffVal1" class="text-2xl font-black text-slate-900 mt-1">--</div>
                <div class="text-[11px] text-slate-400 mt-1" id="compDetail1">驱动: -- | 工时: --</div>
              </div>

              <!-- Period 2 Eff -->
              <div class="bg-white p-4 rounded-2xl border border-amber-200 shadow-sm">
                <div class="text-xs text-amber-700 font-semibold flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-amber-600"></span>
                  时间段 2 综合效率
                </div>
                <div id="compEffVal2" class="text-2xl font-black text-slate-900 mt-1">--</div>
                <div class="text-[11px] text-slate-400 mt-1" id="compDetail2">驱动: -- | 工时: --</div>
              </div>

              <!-- Efficiency Diff (Absolute) -->
              <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                <div class="text-xs text-slate-500 font-medium">效率差值 (时段2 - 时段1)</div>
                <div id="compDiffVal" class="text-2xl font-black text-slate-900 mt-1">--</div>
                <div class="text-[11px] text-slate-400 mt-1" id="compDiffDesc">绝对数值变动</div>
              </div>

              <!-- Efficiency Change Rate (%) -->
              <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                <div class="text-xs text-slate-500 font-medium">效率增幅 (变化百分比)</div>
                <div id="compRateVal" class="text-2xl font-black mt-1">--</div>
                <div class="text-[11px] text-slate-400 mt-1" id="compRateBadge">对比基准期</div>
              </div>
            </div>

            <!-- Efficiency Comparison Line Chart (走势曲线图) -->
            <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
              <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-3">
                <div>
                  <h4 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                    <svg class="w-4 h-4 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z"/></svg>
                    每日效率走势对比曲线
                  </h4>
                  <p class="text-xs text-slate-400 mt-0.5">横坐标按天序同步对比双时段效率走势，平滑曲线呈现效率波动趋势</p>
                </div>
                <div class="flex items-center gap-3 text-xs">
                  <span class="flex items-center gap-1.5"><span class="w-3 h-1 bg-blue-600 rounded"></span><span class="text-slate-600">时段 1 效率</span></span>
                  <span class="flex items-center gap-1.5"><span class="w-3 h-1 bg-amber-500 rounded"></span><span class="text-slate-600">时段 2 效率</span></span>
                  <span id="targetLegend" class="flex items-center gap-1.5"><span class="w-3 h-0.5 border-b-2 border-dashed border-rose-500"></span><span class="text-slate-600">效率目标线</span></span>
                </div>
              </div>
              <div id="chartEfficiencyTrend" class="h-64 sm:h-72 w-full"></div>
            </div>

          </div>

          <!-- ======================================================== -->
          <!-- MODULE 2: 每周效率走势与波动分析 (按周数累加 + 波动曲线图)    -->
          <!-- ======================================================== -->
          <div class="mt-8 pt-6 border-t-2 border-indigo-100 space-y-6">

            <!-- Section Header & Department Selector -->
            <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4 bg-gradient-to-r from-indigo-50/80 via-sky-50/60 to-slate-50 p-5 rounded-2xl border border-indigo-100">
              <div>
                <div class="flex items-center gap-2">
                  <span class="w-3 h-3 rounded-full bg-emerald-600 animate-pulse"></span>
                  <h3 class="text-lg sm:text-xl font-bold text-slate-900">每周人员效率走势与波动分析</h3>
                </div>
                <p class="text-xs sm:text-sm text-slate-500 mt-1">
                  自动抓取工时表与关键驱动指标表中相同【周数】进行独立累加，计算每周效率（关键驱动累加 ÷ 工时累加），监测上下波动幅度。
                </p>
              </div>

              <!-- Department Selector -->
              <div class="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-indigo-200 shadow-sm">
                <label class="text-xs font-bold text-slate-700 whitespace-nowrap">查看部门：</label>
                <select id="weeklyDeptSelect" class="table-input w-40 text-xs font-bold py-1 bg-white border-0 focus:ring-0" onchange="updateWeeklyEfficiencyAnalysis()"><option value="商场OMNI" selected>商场OMNI</option>
<option value="销售部(线下)">销售部(线下)</option>
<option value="顾客关系部(线下)">顾客关系部(线下)</option>
<option value="物流部">物流部</option>
<option value="餐饮部">餐饮部</option></select>
              </div>
            </div>

            <!-- Fluctuation Amplitude KPI Summary Cards -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <!-- Average Weekly Efficiency -->
              <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                <div class="text-xs text-slate-500 font-medium">全期平均周效率</div>
                <div id="weeklyAvgEff" class="text-2xl font-black text-slate-900 mt-1">--</div>
                <div class="text-[11px] text-slate-400 mt-1" id="weeklyTotalWeeksTag">共覆盖 -- 个自然周</div>
              </div>

              <!-- Peak Efficiency -->
              <div class="bg-white p-4 rounded-2xl border border-emerald-200 shadow-sm bg-gradient-to-b from-white to-emerald-50/30">
                <div class="text-xs text-emerald-700 font-semibold flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
                  最高周效率 (波峰)
                </div>
                <div id="weeklyMaxEff" class="text-2xl font-black text-emerald-700 mt-1">--</div>
                <div class="text-[11px] text-emerald-600 font-medium mt-1" id="weeklyMaxWeekDesc">发生于 第 -- 周</div>
              </div>

              <!-- Lowest Efficiency -->
              <div class="bg-white p-4 rounded-2xl border border-rose-200 shadow-sm bg-gradient-to-b from-white to-rose-50/30">
                <div class="text-xs text-rose-700 font-semibold flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-rose-500"></span>
                  最低周效率 (波谷)
                </div>
                <div id="weeklyMinEff" class="text-2xl font-black text-rose-700 mt-1">--</div>
                <div class="text-[11px] text-rose-600 font-medium mt-1" id="weeklyMinWeekDesc">发生于 第 -- 周</div>
              </div>

              <!-- Fluctuation Amplitude -->
              <div class="bg-white p-4 rounded-2xl border border-indigo-200 shadow-sm bg-gradient-to-b from-white to-indigo-50/30">
                <div class="text-xs text-indigo-700 font-semibold flex items-center gap-1.5">
                  <span class="w-2 h-2 rounded-full bg-indigo-600"></span>
                  上下波动幅度 (极差 / 振幅)
                </div>
                <div id="weeklyFluctuationRange" class="text-2xl font-black text-indigo-900 mt-1">--</div>
                <div class="text-[11px] text-indigo-600 font-medium mt-1" id="weeklyFluctuationRate">波动率：--%</div>
              </div>
            </div>

            <!-- Weekly Efficiency Curve Chart -->
            <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-3">
              <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                <div>
                  <h4 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                    <svg class="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"/></svg>
                    每周效率波动曲线图
                  </h4>
                  <p class="text-xs text-slate-400 mt-0.5">平滑曲线追踪各周效率起伏，自动标注极值点、目标线与均值基准线</p>
                </div>
                <div class="flex flex-wrap items-center gap-3 text-xs">
                  <span class="flex items-center gap-1.5"><span class="w-3 h-1 bg-emerald-500 rounded"></span><span class="text-slate-600">周实际效率</span></span>
                  <span class="flex items-center gap-1.5"><span class="w-3 h-0.5 border-b-2 border-dashed border-blue-500"></span><span class="text-slate-600">平均效率线</span></span>
                  <span id="weeklyTargetLegend" class="flex items-center gap-1.5"><span class="w-3 h-0.5 border-b-2 border-dashed border-rose-500"></span><span class="text-slate-600">效率目标线</span></span>
                </div>
              </div>
              <div id="chartWeeklyTrend" class="h-64 sm:h-72 w-full"></div>
            </div>

            <!-- Weekly Data Breakdown Grid / Table -->
            <div class="overflow-x-auto rounded-2xl border border-slate-200 shadow-sm">
              <table class="w-full border-collapse text-xs sm:text-sm">
                <thead>
                  <tr class="bg-slate-900 text-white">
                    <th class="px-4 py-3 text-left font-bold">周数</th>
                    <th class="px-4 py-3 text-center font-bold">包含天数</th>
                    <th class="px-4 py-3 text-right font-bold text-indigo-200">关键驱动指标 (周累加)</th>
                    <th class="px-4 py-3 text-right font-bold text-blue-200">工时 (周累加)</th>
                    <th class="px-4 py-3 text-right font-bold text-amber-300">周实际效率</th>
                    <th class="px-4 py-3 text-center font-bold">较上周环比波动</th>
                  </tr>
                </thead>
                <tbody id="weeklyTableBody" class="divide-y divide-slate-200 bg-white">
                  <!-- Dynamically rendered -->
                </tbody>
              </table>
            </div>

          </div>

          

        </div>
      </div>

      <!-- ======================================================== -->
      <!-- PAGE 3: 模块 3 周内工时占比 😊                            -->
      <!-- ======================================================== -->
      <div id="page-module3" class="space-y-8 hidden">

        <!-- Breadcrumb / Header -->
        <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2.5 py-1 rounded-md bg-amber-100 text-amber-800 text-xs font-bold">模块 3</span>
              <h2 class="text-xl sm:text-2xl font-bold text-slate-900 flex items-center gap-2">
                <span>周内工时占比 😊</span>
              </h2>
            </div>
            <p class="text-xs sm:text-sm text-slate-500 mt-1">
              直接使用模块二上传的【工时表格】数据，按选定时间段对同星期求平均值，计算周日开始至周六结束的一周工时占比分布。
            </p>
          </div>
          <div class="text-xs text-amber-800 bg-amber-50/80 border border-amber-200 px-3.5 py-2 rounded-xl flex items-center gap-2 shadow-sm">
            <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
            <span class="font-semibold">已联动模块二工时数据源</span>
          </div>
        </div>

        <div class="bg-white rounded-3xl p-6 shadow-md border border-slate-200 space-y-6">

          <!-- Section Header & Controls -->
          <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4 bg-gradient-to-r from-amber-50/90 via-orange-50/60 to-slate-50 p-5 rounded-2xl border border-amber-200">
            <div>
              <div class="flex items-center gap-2">
                <span class="w-3 h-3 rounded-full bg-amber-500"></span>
                <h3 class="text-lg font-bold text-slate-900">一周七天工时占比分析 (周日~周六)</h3>
              </div>
              <p class="text-xs sm:text-sm text-slate-500 mt-0.5">
                先求同星期日均值以排除天数不均，再计算 7 天占比（周日开始，周六结束）
              </p>
            </div>

            <!-- Department Selector -->
            <div class="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-amber-200 shadow-sm">
              <label class="text-xs font-bold text-slate-700 whitespace-nowrap">分析部门：</label>
              <select id="weekdayDeptSelect" class="table-input w-40 text-xs font-bold py-1 bg-white border-0 focus:ring-0" onchange="updateWeekdayDistributionAnalysis()"><option value="商场OMNI" selected>商场OMNI</option>
<option value="销售部(线下)">销售部(线下)</option>
<option value="顾客关系部(线下)">顾客关系部(线下)</option>
<option value="物流部">物流部</option>
<option value="餐饮部">餐饮部</option></select>
            </div>
          </div>

          <!-- Date Range Filter Bar for Weekday Analysis -->
          <div class="bg-slate-50 rounded-2xl p-4 sm:p-5 border border-slate-200 space-y-3">
            <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
              <span class="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-2">
                <svg class="w-4 h-4 text-amber-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
                时间段筛选（选定任意周期，系统自动匹配并计算每个星期的平均工时）
              </span>
              <div class="flex items-center gap-1.5">
                <button onclick="setWeekdayRangeQuick('all')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">全量区间</button>
                <button onclick="setWeekdayRangeQuick('latest30')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">近30天</button>
                <button onclick="setWeekdayRangeQuick('latest14')" class="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition">近两周(14天)</button>
              </div>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-12 gap-3 items-center">
              <div class="lg:col-span-4">
                <label class="text-[11px] text-slate-500 font-medium block mb-1">起始日期</label>
                <input type="date" id="weekdayStartDate" class="table-input py-1.5 text-xs bg-white" onchange="updateWeekdayDistributionAnalysis()" />
              </div>
              <div class="lg:col-span-4">
                <label class="text-[11px] text-slate-500 font-medium block mb-1">截止日期</label>
                <input type="date" id="weekdayEndDate" class="table-input py-1.5 text-xs bg-white" onchange="updateWeekdayDistributionAnalysis()" />
              </div>
              <div class="lg:col-span-4 self-end">
                <div class="bg-amber-50/70 border border-amber-200 rounded-xl px-3 py-2 text-xs text-amber-950 font-medium" id="weekdayRangeStatus">
                  统计分析中...
                </div>
              </div>
            </div>
          </div>

          <!-- KPI Cards: Weekday Summary -->
          <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <!-- Total of Daily Averages -->
            <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
              <div class="text-xs text-slate-500 font-medium">周日均工时总和</div>
              <div id="weekdaySumAvgHours" class="text-2xl font-black text-slate-900 mt-1">-- 小时</div>
              <div class="text-[11px] text-slate-400 mt-1" id="weekdayCoverageTag">7天日均之和 (占比分母)</div>
            </div>

            <!-- Peak Day -->
            <div class="bg-white p-4 rounded-2xl border border-amber-200 shadow-sm bg-gradient-to-b from-white to-amber-50/30">
              <div class="text-xs text-amber-700 font-semibold flex items-center gap-1.5">
                <span class="w-2 h-2 rounded-full bg-amber-500"></span>
                工时占比最高日
              </div>
              <div id="weekdayPeakDay" class="text-2xl font-black text-amber-800 mt-1">--</div>
              <div class="text-[11px] text-amber-600 font-medium mt-1" id="weekdayPeakDetail">占比：--% | 日均：-- 小时</div>
            </div>

            <!-- Lowest Day -->
            <div class="bg-white p-4 rounded-2xl border border-sky-200 shadow-sm bg-gradient-to-b from-white to-sky-50/30">
              <div class="text-xs text-sky-700 font-semibold flex items-center gap-1.5">
                <span class="w-2 h-2 rounded-full bg-sky-500"></span>
                工时占比最低日
              </div>
              <div id="weekdayValleyDay" class="text-2xl font-black text-sky-800 mt-1">--</div>
              <div class="text-[11px] text-sky-600 font-medium mt-1" id="weekdayValleyDetail">占比：--% | 日均：-- 小时</div>
            </div>

            <!-- Weekend vs Weekday Ratio -->
            <div class="bg-white p-4 rounded-2xl border border-indigo-200 shadow-sm bg-gradient-to-b from-white to-indigo-50/30">
              <div class="text-xs text-indigo-700 font-semibold flex items-center gap-1.5">
                <span class="w-2 h-2 rounded-full bg-indigo-600"></span>
                周末 (周六+周日) 占比
              </div>
              <div id="weekendRatioVal" class="text-2xl font-black text-indigo-900 mt-1">--%</div>
              <div class="text-[11px] text-indigo-600 font-medium mt-1" id="weekdayRatioDetail">工作日(周一~周五)：--%</div>
            </div>
          </div>

          <!-- Line Chart: 7 Days of Week Hours Distribution (Sun ~ Sat) -->
          <div class="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-3">
            <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
              <div>
                <h4 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <svg class="w-4 h-4 text-amber-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z"/></svg>
                  一周七天每天工时占比折线图 (周日开始，周六结束)
                </h4>
                <p class="text-xs text-slate-400 mt-0.5">严格按“周日、周一、周二、周三、周四、周五、周六”顺序呈现各天日均工时占比</p>
              </div>
              <div class="flex flex-wrap items-center gap-3 text-xs">
                <span class="flex items-center gap-1.5"><span class="w-3 h-1 bg-amber-500 rounded"></span><span class="text-slate-600">工时占比 (%)</span></span>
              </div>
            </div>
            <div id="chartWeekdayDistribution" class="h-64 sm:h-72 w-full"></div>
          </div>

          <!-- 7 Days Breakdown Table -->
          <div class="overflow-x-auto rounded-2xl border border-slate-200 shadow-sm">
            <table class="w-full border-collapse text-xs sm:text-sm">
              <thead>
                <tr class="bg-slate-900 text-white">
                  <th class="px-4 py-3 text-left font-bold">星期 (周日~周六)</th>
                  <th class="px-4 py-3 text-center font-bold">匹配天数</th>
                  <th class="px-4 py-3 text-right font-bold text-blue-200">区间总工时 (小时)</th>
                  <th class="px-4 py-3 text-right font-bold text-emerald-200">日均工时 (平均值)</th>
                  <th class="px-4 py-3 text-right font-bold text-amber-300">工时占比 (%)</th>
                </tr>
              </thead>
              <tbody id="weekdayTableBody" class="divide-y divide-slate-200 bg-white">
                <!-- Dynamically rendered -->
              </tbody>
            </table>
          </div>

        </div>
      </div>

      <!-- ======================================================== -->
      <!-- PAGE 4: 模块 4 (预留)                                     -->
      <!-- ======================================================== -->
      <div id="page-module4" class="space-y-6 hidden">
        <div class="bg-white rounded-3xl p-12 shadow-md border border-slate-200 text-center space-y-4">
          <div class="w-16 h-16 rounded-2xl bg-purple-50 text-purple-600 mx-auto flex items-center justify-center text-2xl font-bold">
            4
          </div>
          <h2 class="text-2xl font-bold text-slate-900">模块 4：待配置</h2>
          <p class="text-slate-500 max-w-md mx-auto text-sm">
            本模块已为您预留，作为第4个独立功能页，支持灵活扩展与数据集成。
          </p>
        </div>
      </div>

    </main>
  </div>

  <!-- Footer -->
  <footer class="bg-white border-t border-slate-200 py-4 text-center text-xs text-slate-400">
    宜家苏州商场 · 人员效率看板 · 4模块分页系统
  </footer>

  <script>
  // Smart Storage Adapter
  const EfficiencyStorage = {
    isGenie: () => typeof window.genie !== 'undefined' && typeof window.genie.storage !== 'undefined',

    async set(key, value) {
      if (this.isGenie()) {
        try {
          await window.genie.storage.command('set', key, value);
          return;
        } catch (e) { console.warn("Genie storage set failed, falling back", e); }
      }
      localStorage.setItem(key, JSON.stringify(value));
    },

    async get(key) {
      if (this.isGenie()) {
        try {
          const val = await window.genie.storage.command('get', key);
          if (val !== null) return val;
        } catch (e) { console.warn("Genie storage get failed, falling back", e); }
      }
      const local = localStorage.getItem(key);
      return local ? JSON.parse(local) : null;
    },

    async del(key) {
      if (this.isGenie()) {
        try {
          await window.genie.storage.command('del', key);
        } catch (e) {}
      }
      localStorage.removeItem(key);
    }
  };

  // ==========================================
  // TAB PAGINATION SWITCHING
  // ==========================================
  const PAGE_STORAGE_KEY = 'suzhou_eff_active_page_v1';
  let currentPage = 'module1';

  function switchPage(pageId) {
    currentPage = pageId;
    localStorage.setItem(PAGE_STORAGE_KEY, pageId);

    // Toggle pages
    ['module1', 'module2', 'module3', 'module4'].forEach(id => {
      const pageEl = document.getElementById('page-' + id);
      const tabBtn = document.getElementById('tab-btn-' + id);
      if (pageEl && tabBtn) {
        if (id === pageId) {
          pageEl.classList.remove('hidden');
          tabBtn.classList.add('active');
        } else {
          pageEl.classList.add('hidden');
          tabBtn.classList.remove('active');
        }
      }
    });

    // Trigger chart resize after rendering
    if (pageId === 'module1') {
      setTimeout(() => {
        charts.forEach(c => c && c.resize());
      }, 50);
    } else if (pageId === 'module2') {
      setTimeout(() => {
        if (compareChart) compareChart.resize();
        if (weeklyChart) weeklyChart.resize();
        updateEfficiencyComparison();
        updateWeeklyEfficiencyAnalysis();
      }, 50);
    } else if (pageId === 'module3') {
      setTimeout(() => {
        if (weekdayChart) weekdayChart.resize();
        updateWeekdayDistributionAnalysis();
      }, 50);
    }
  }

  // ==========================================
  // MODULE 2: 人员效率 LOGIC (DUAL MATRIX UPLOAD)
  // ==========================================
  const EFF_DEPTS = [{"dept": "商场OMNI", "target": "764", "numericTarget": 764}, {"dept": "销售部(线下)", "target": "1732", "numericTarget": 1732}, {"dept": "顾客关系部(线下)", "target": "8.36", "numericTarget": 8.36}, {"dept": "物流部", "target": "0.59", "numericTarget": 0.59}, {"dept": "餐饮部", "target": "8.45", "numericTarget": 8.45}];
  const DEFAULT_HOURS_DATA = {"2026-09-01": {"商场OMNI": 96.4, "销售部(线下)": 41.8, "餐饮部": 152.2, "顾客关系部(线下)": 51.5, "物流部": 80.8}, "2026-09-02": {"商场OMNI": 92.6, "销售部(线下)": 41.6, "餐饮部": 157.5, "顾客关系部(线下)": 49.2, "物流部": 80.4}, "2026-09-03": {"商场OMNI": 96.9, "销售部(线下)": 41.2, "餐饮部": 149.2, "顾客关系部(线下)": 51.2, "物流部": 77.6}, "2026-09-04": {"商场OMNI": 90.2, "销售部(线下)": 42.0, "餐饮部": 151.5, "顾客关系部(线下)": 49.8, "物流部": 76.4}, "2026-09-05": {"商场OMNI": 96.0, "销售部(线下)": 39.2, "餐饮部": 144.8, "顾客关系部(线下)": 50.7, "物流部": 76.4}, "2026-09-06": {"商场OMNI": 97.4, "销售部(线下)": 41.4, "餐饮部": 148.5, "顾客关系部(线下)": 49.0, "物流部": 78.8}, "2026-09-07": {"商场OMNI": 93.1, "销售部(线下)": 39.6, "餐饮部": 155.2, "顾客关系部(线下)": 49.8, "物流部": 81.6}, "2026-09-08": {"商场OMNI": 99.8, "销售部(线下)": 41.2, "餐饮部": 154.5, "顾客关系部(线下)": 48.8, "物流部": 80.4}, "2026-09-09": {"商场OMNI": 90.2, "销售部(线下)": 41.2, "餐饮部": 149.2, "顾客关系部(线下)": 49.2, "物流部": 81.6}, "2026-09-10": {"商场OMNI": 91.2, "销售部(线下)": 38.8, "餐饮部": 150.0, "顾客关系部(线下)": 49.0, "物流部": 81.2}, "2026-09-11": {"商场OMNI": 94.0, "销售部(线下)": 39.8, "餐饮部": 142.5, "顾客关系部(线下)": 51.5, "物流部": 78.0}, "2026-09-12": {"商场OMNI": 95.5, "销售部(线下)": 41.6, "餐饮部": 156.0, "顾客关系部(线下)": 52.5, "物流部": 84.0}, "2026-09-13": {"商场OMNI": 94.5, "销售部(线下)": 38.4, "餐饮部": 152.2, "顾客关系部(线下)": 49.0, "物流部": 77.6}, "2026-09-14": {"商场OMNI": 94.0, "销售部(线下)": 40.2, "餐饮部": 145.5, "顾客关系部(线下)": 50.5, "物流部": 77.6}, "2026-09-15": {"商场OMNI": 90.2, "销售部(线下)": 42.0, "餐饮部": 148.5, "顾客关系部(线下)": 50.5, "物流部": 83.6}, "2026-09-16": {"商场OMNI": 99.3, "销售部(线下)": 41.0, "餐饮部": 148.5, "顾客关系部(线下)": 47.8, "物流部": 79.6}, "2026-09-17": {"商场OMNI": 99.3, "销售部(线下)": 38.2, "餐饮部": 151.5, "顾客关系部(线下)": 48.8, "物流部": 78.8}, "2026-09-18": {"商场OMNI": 96.4, "销售部(线下)": 41.6, "餐饮部": 144.8, "顾客关系部(线下)": 51.0, "物流部": 80.0}, "2026-09-19": {"商场OMNI": 95.0, "销售部(线下)": 38.0, "餐饮部": 153.8, "顾客关系部(线下)": 49.5, "物流部": 84.0}, "2026-09-20": {"商场OMNI": 97.4, "销售部(线下)": 39.6, "餐饮部": 142.5, "顾客关系部(线下)": 51.0, "物流部": 83.2}, "2026-09-21": {"商场OMNI": 96.9, "销售部(线下)": 40.4, "餐饮部": 148.5, "顾客关系部(线下)": 49.0, "物流部": 76.8}, "2026-09-22": {"商场OMNI": 92.6, "销售部(线下)": 40.2, "餐饮部": 148.5, "顾客关系部(线下)": 48.8, "物流部": 76.0}, "2026-09-23": {"商场OMNI": 91.2, "销售部(线下)": 40.4, "餐饮部": 147.8, "顾客关系部(线下)": 47.5, "物流部": 81.6}, "2026-09-24": {"商场OMNI": 90.2, "销售部(线下)": 41.0, "餐饮部": 150.0, "顾客关系部(线下)": 52.5, "物流部": 84.0}, "2026-09-25": {"商场OMNI": 96.4, "销售部(线下)": 40.4, "餐饮部": 147.0, "顾客关系部(线下)": 48.8, "物流部": 76.0}, "2026-09-26": {"商场OMNI": 95.5, "销售部(线下)": 38.6, "餐饮部": 152.2, "顾客关系部(线下)": 52.2, "物流部": 81.6}, "2026-09-27": {"商场OMNI": 92.1, "销售部(线下)": 40.8, "餐饮部": 143.2, "顾客关系部(线下)": 48.0, "物流部": 77.6}, "2026-09-28": {"商场OMNI": 96.0, "销售部(线下)": 40.0, "餐饮部": 156.0, "顾客关系部(线下)": 49.2, "物流部": 82.0}, "2026-09-29": {"商场OMNI": 96.4, "销售部(线下)": 38.2, "餐饮部": 154.5, "顾客关系部(线下)": 47.8, "物流部": 82.8}, "2026-09-30": {"商场OMNI": 95.0, "销售部(线下)": 40.0, "餐饮部": 147.0, "顾客关系部(线下)": 49.0, "物流部": 78.8}};
  const DEFAULT_DRIVER_DATA = {"2026-09-01": {"商场OMNI": 113300.0, "销售部(线下)": 109000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1166.0, "物流部": 3570.0}, "2026-09-02": {"商场OMNI": 104500.0, "销售部(线下)": 108000.0, "餐饮部": 2420.0, "顾客关系部(线下)": 1067.0, "物流部": 3535.0}, "2026-09-03": {"商场OMNI": 114400.0, "销售部(线下)": 106000.0, "餐饮部": 2178.0, "顾客关系部(线下)": 1155.0, "物流部": 3290.0}, "2026-09-04": {"商场OMNI": 99000.0, "销售部(线下)": 110000.0, "餐饮部": 2244.0, "顾客关系部(线下)": 1089.0, "物流部": 3185.0}, "2026-09-05": {"商场OMNI": 112200.0, "销售部(线下)": 96000.0, "餐饮部": 2046.0, "顾客关系部(线下)": 1133.0, "物流部": 3185.0}, "2026-09-06": {"商场OMNI": 115500.0, "销售部(线下)": 107000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1056.0, "物流部": 3395.0}, "2026-09-07": {"商场OMNI": 105600.0, "销售部(线下)": 98000.0, "餐饮部": 2354.0, "顾客关系部(线下)": 1089.0, "物流部": 3640.0}, "2026-09-08": {"商场OMNI": 121000.0, "销售部(线下)": 106000.0, "餐饮部": 2332.0, "顾客关系部(线下)": 1045.0, "物流部": 3535.0}, "2026-09-09": {"商场OMNI": 99000.0, "销售部(线下)": 106000.0, "餐饮部": 2178.0, "顾客关系部(线下)": 1067.0, "物流部": 3640.0}, "2026-09-10": {"商场OMNI": 101200.0, "销售部(线下)": 94000.0, "餐饮部": 2200.0, "顾客关系部(线下)": 1056.0, "物流部": 3605.0}, "2026-09-11": {"商场OMNI": 107800.0, "销售部(线下)": 99000.0, "餐饮部": 1980.0, "顾客关系部(线下)": 1166.0, "物流部": 3325.0}, "2026-09-12": {"商场OMNI": 111100.0, "销售部(线下)": 108000.0, "餐饮部": 2376.0, "顾客关系部(线下)": 1210.0, "物流部": 3850.0}, "2026-09-13": {"商场OMNI": 108900.0, "销售部(线下)": 92000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1056.0, "物流部": 3290.0}, "2026-09-14": {"商场OMNI": 107800.0, "销售部(线下)": 101000.0, "餐饮部": 2068.0, "顾客关系部(线下)": 1122.0, "物流部": 3290.0}, "2026-09-15": {"商场OMNI": 99000.0, "销售部(线下)": 110000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1122.0, "物流部": 3815.0}, "2026-09-16": {"商场OMNI": 119900.0, "销售部(线下)": 105000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1001.0, "物流部": 3465.0}, "2026-09-17": {"商场OMNI": 119900.0, "销售部(线下)": 91000.0, "餐饮部": 2244.0, "顾客关系部(线下)": 1045.0, "物流部": 3395.0}, "2026-09-18": {"商场OMNI": 113300.0, "销售部(线下)": 108000.0, "餐饮部": 2046.0, "顾客关系部(线下)": 1144.0, "物流部": 3500.0}, "2026-09-19": {"商场OMNI": 110000.0, "销售部(线下)": 90000.0, "餐饮部": 2310.0, "顾客关系部(线下)": 1078.0, "物流部": 3850.0}, "2026-09-20": {"商场OMNI": 115500.0, "销售部(线下)": 98000.0, "餐饮部": 1980.0, "顾客关系部(线下)": 1144.0, "物流部": 3780.0}, "2026-09-21": {"商场OMNI": 114400.0, "销售部(线下)": 102000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1056.0, "物流部": 3220.0}, "2026-09-22": {"商场OMNI": 104500.0, "销售部(线下)": 101000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1045.0, "物流部": 3150.0}, "2026-09-23": {"商场OMNI": 101200.0, "销售部(线下)": 102000.0, "餐饮部": 2134.0, "顾客关系部(线下)": 990.0, "物流部": 3640.0}, "2026-09-24": {"商场OMNI": 99000.0, "销售部(线下)": 105000.0, "餐饮部": 2200.0, "顾客关系部(线下)": 1210.0, "物流部": 3850.0}, "2026-09-25": {"商场OMNI": 113300.0, "销售部(线下)": 102000.0, "餐饮部": 2112.0, "顾客关系部(线下)": 1045.0, "物流部": 3150.0}, "2026-09-26": {"商场OMNI": 111100.0, "销售部(线下)": 93000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1199.0, "物流部": 3640.0}, "2026-09-27": {"商场OMNI": 103400.0, "销售部(线下)": 104000.0, "餐饮部": 2002.0, "顾客关系部(线下)": 1012.0, "物流部": 3290.0}, "2026-09-28": {"商场OMNI": 112200.0, "销售部(线下)": 100000.0, "餐饮部": 2376.0, "顾客关系部(线下)": 1067.0, "物流部": 3675.0}, "2026-09-29": {"商场OMNI": 113300.0, "销售部(线下)": 91000.0, "餐饮部": 2332.0, "顾客关系部(线下)": 1001.0, "物流部": 3745.0}, "2026-09-30": {"商场OMNI": 110000.0, "销售部(线下)": 100000.0, "餐饮部": 2112.0, "顾客关系部(线下)": 1056.0, "物流部": 3395.0}};
  const DEFAULT_HOURS_ROWS = [{"日期": "2026-09-01", "周数": 36, "星期": "Tue", "商场OMNI": 96.4, "销售部(线下)": 41.8, "餐饮部": 152.2, "顾客关系部(线下)": 51.5, "物流部": 80.8}, {"日期": "2026-09-02", "周数": 36, "星期": "Wed", "商场OMNI": 92.6, "销售部(线下)": 41.6, "餐饮部": 157.5, "顾客关系部(线下)": 49.2, "物流部": 80.4}, {"日期": "2026-09-03", "周数": 36, "星期": "Thu", "商场OMNI": 96.9, "销售部(线下)": 41.2, "餐饮部": 149.2, "顾客关系部(线下)": 51.2, "物流部": 77.6}, {"日期": "2026-09-04", "周数": 36, "星期": "Fri", "商场OMNI": 90.2, "销售部(线下)": 42.0, "餐饮部": 151.5, "顾客关系部(线下)": 49.8, "物流部": 76.4}, {"日期": "2026-09-05", "周数": 36, "星期": "Sat", "商场OMNI": 96.0, "销售部(线下)": 39.2, "餐饮部": 144.8, "顾客关系部(线下)": 50.7, "物流部": 76.4}, {"日期": "2026-09-06", "周数": 36, "星期": "Sun", "商场OMNI": 97.4, "销售部(线下)": 41.4, "餐饮部": 148.5, "顾客关系部(线下)": 49.0, "物流部": 78.8}, {"日期": "2026-09-07", "周数": 37, "星期": "Mon", "商场OMNI": 93.1, "销售部(线下)": 39.6, "餐饮部": 155.2, "顾客关系部(线下)": 49.8, "物流部": 81.6}, {"日期": "2026-09-08", "周数": 37, "星期": "Tue", "商场OMNI": 99.8, "销售部(线下)": 41.2, "餐饮部": 154.5, "顾客关系部(线下)": 48.8, "物流部": 80.4}, {"日期": "2026-09-09", "周数": 37, "星期": "Wed", "商场OMNI": 90.2, "销售部(线下)": 41.2, "餐饮部": 149.2, "顾客关系部(线下)": 49.2, "物流部": 81.6}, {"日期": "2026-09-10", "周数": 37, "星期": "Thu", "商场OMNI": 91.2, "销售部(线下)": 38.8, "餐饮部": 150.0, "顾客关系部(线下)": 49.0, "物流部": 81.2}, {"日期": "2026-09-11", "周数": 37, "星期": "Fri", "商场OMNI": 94.0, "销售部(线下)": 39.8, "餐饮部": 142.5, "顾客关系部(线下)": 51.5, "物流部": 78.0}, {"日期": "2026-09-12", "周数": 37, "星期": "Sat", "商场OMNI": 95.5, "销售部(线下)": 41.6, "餐饮部": 156.0, "顾客关系部(线下)": 52.5, "物流部": 84.0}, {"日期": "2026-09-13", "周数": 37, "星期": "Sun", "商场OMNI": 94.5, "销售部(线下)": 38.4, "餐饮部": 152.2, "顾客关系部(线下)": 49.0, "物流部": 77.6}, {"日期": "2026-09-14", "周数": 38, "星期": "Mon", "商场OMNI": 94.0, "销售部(线下)": 40.2, "餐饮部": 145.5, "顾客关系部(线下)": 50.5, "物流部": 77.6}, {"日期": "2026-09-15", "周数": 38, "星期": "Tue", "商场OMNI": 90.2, "销售部(线下)": 42.0, "餐饮部": 148.5, "顾客关系部(线下)": 50.5, "物流部": 83.6}, {"日期": "2026-09-16", "周数": 38, "星期": "Wed", "商场OMNI": 99.3, "销售部(线下)": 41.0, "餐饮部": 148.5, "顾客关系部(线下)": 47.8, "物流部": 79.6}, {"日期": "2026-09-17", "周数": 38, "星期": "Thu", "商场OMNI": 99.3, "销售部(线下)": 38.2, "餐饮部": 151.5, "顾客关系部(线下)": 48.8, "物流部": 78.8}, {"日期": "2026-09-18", "周数": 38, "星期": "Fri", "商场OMNI": 96.4, "销售部(线下)": 41.6, "餐饮部": 144.8, "顾客关系部(线下)": 51.0, "物流部": 80.0}, {"日期": "2026-09-19", "周数": 38, "星期": "Sat", "商场OMNI": 95.0, "销售部(线下)": 38.0, "餐饮部": 153.8, "顾客关系部(线下)": 49.5, "物流部": 84.0}, {"日期": "2026-09-20", "周数": 38, "星期": "Sun", "商场OMNI": 97.4, "销售部(线下)": 39.6, "餐饮部": 142.5, "顾客关系部(线下)": 51.0, "物流部": 83.2}, {"日期": "2026-09-21", "周数": 39, "星期": "Mon", "商场OMNI": 96.9, "销售部(线下)": 40.4, "餐饮部": 148.5, "顾客关系部(线下)": 49.0, "物流部": 76.8}, {"日期": "2026-09-22", "周数": 39, "星期": "Tue", "商场OMNI": 92.6, "销售部(线下)": 40.2, "餐饮部": 148.5, "顾客关系部(线下)": 48.8, "物流部": 76.0}, {"日期": "2026-09-23", "周数": 39, "星期": "Wed", "商场OMNI": 91.2, "销售部(线下)": 40.4, "餐饮部": 147.8, "顾客关系部(线下)": 47.5, "物流部": 81.6}, {"日期": "2026-09-24", "周数": 39, "星期": "Thu", "商场OMNI": 90.2, "销售部(线下)": 41.0, "餐饮部": 150.0, "顾客关系部(线下)": 52.5, "物流部": 84.0}, {"日期": "2026-09-25", "周数": 39, "星期": "Fri", "商场OMNI": 96.4, "销售部(线下)": 40.4, "餐饮部": 147.0, "顾客关系部(线下)": 48.8, "物流部": 76.0}, {"日期": "2026-09-26", "周数": 39, "星期": "Sat", "商场OMNI": 95.5, "销售部(线下)": 38.6, "餐饮部": 152.2, "顾客关系部(线下)": 52.2, "物流部": 81.6}, {"日期": "2026-09-27", "周数": 39, "星期": "Sun", "商场OMNI": 92.1, "销售部(线下)": 40.8, "餐饮部": 143.2, "顾客关系部(线下)": 48.0, "物流部": 77.6}, {"日期": "2026-09-28", "周数": 40, "星期": "Mon", "商场OMNI": 96.0, "销售部(线下)": 40.0, "餐饮部": 156.0, "顾客关系部(线下)": 49.2, "物流部": 82.0}, {"日期": "2026-09-29", "周数": 40, "星期": "Tue", "商场OMNI": 96.4, "销售部(线下)": 38.2, "餐饮部": 154.5, "顾客关系部(线下)": 47.8, "物流部": 82.8}, {"日期": "2026-09-30", "周数": 40, "星期": "Wed", "商场OMNI": 95.0, "销售部(线下)": 40.0, "餐饮部": 147.0, "顾客关系部(线下)": 49.0, "物流部": 78.8}];
  const DEFAULT_DRIVER_ROWS = [{"日期": "2026-09-01", "周数": 36, "星期": "Tue", "商场OMNI": 113300.0, "销售部(线下)": 109000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1166.0, "物流部": 3570.0}, {"日期": "2026-09-02", "周数": 36, "星期": "Wed", "商场OMNI": 104500.0, "销售部(线下)": 108000.0, "餐饮部": 2420.0, "顾客关系部(线下)": 1067.0, "物流部": 3535.0}, {"日期": "2026-09-03", "周数": 36, "星期": "Thu", "商场OMNI": 114400.0, "销售部(线下)": 106000.0, "餐饮部": 2178.0, "顾客关系部(线下)": 1155.0, "物流部": 3290.0}, {"日期": "2026-09-04", "周数": 36, "星期": "Fri", "商场OMNI": 99000.0, "销售部(线下)": 110000.0, "餐饮部": 2244.0, "顾客关系部(线下)": 1089.0, "物流部": 3185.0}, {"日期": "2026-09-05", "周数": 36, "星期": "Sat", "商场OMNI": 112200.0, "销售部(线下)": 96000.0, "餐饮部": 2046.0, "顾客关系部(线下)": 1133.0, "物流部": 3185.0}, {"日期": "2026-09-06", "周数": 36, "星期": "Sun", "商场OMNI": 115500.0, "销售部(线下)": 107000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1056.0, "物流部": 3395.0}, {"日期": "2026-09-07", "周数": 37, "星期": "Mon", "商场OMNI": 105600.0, "销售部(线下)": 98000.0, "餐饮部": 2354.0, "顾客关系部(线下)": 1089.0, "物流部": 3640.0}, {"日期": "2026-09-08", "周数": 37, "星期": "Tue", "商场OMNI": 121000.0, "销售部(线下)": 106000.0, "餐饮部": 2332.0, "顾客关系部(线下)": 1045.0, "物流部": 3535.0}, {"日期": "2026-09-09", "周数": 37, "星期": "Wed", "商场OMNI": 99000.0, "销售部(线下)": 106000.0, "餐饮部": 2178.0, "顾客关系部(线下)": 1067.0, "物流部": 3640.0}, {"日期": "2026-09-10", "周数": 37, "星期": "Thu", "商场OMNI": 101200.0, "销售部(线下)": 94000.0, "餐饮部": 2200.0, "顾客关系部(线下)": 1056.0, "物流部": 3605.0}, {"日期": "2026-09-11", "周数": 37, "星期": "Fri", "商场OMNI": 107800.0, "销售部(线下)": 99000.0, "餐饮部": 1980.0, "顾客关系部(线下)": 1166.0, "物流部": 3325.0}, {"日期": "2026-09-12", "周数": 37, "星期": "Sat", "商场OMNI": 111100.0, "销售部(线下)": 108000.0, "餐饮部": 2376.0, "顾客关系部(线下)": 1210.0, "物流部": 3850.0}, {"日期": "2026-09-13", "周数": 37, "星期": "Sun", "商场OMNI": 108900.0, "销售部(线下)": 92000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1056.0, "物流部": 3290.0}, {"日期": "2026-09-14", "周数": 38, "星期": "Mon", "商场OMNI": 107800.0, "销售部(线下)": 101000.0, "餐饮部": 2068.0, "顾客关系部(线下)": 1122.0, "物流部": 3290.0}, {"日期": "2026-09-15", "周数": 38, "星期": "Tue", "商场OMNI": 99000.0, "销售部(线下)": 110000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1122.0, "物流部": 3815.0}, {"日期": "2026-09-16", "周数": 38, "星期": "Wed", "商场OMNI": 119900.0, "销售部(线下)": 105000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1001.0, "物流部": 3465.0}, {"日期": "2026-09-17", "周数": 38, "星期": "Thu", "商场OMNI": 119900.0, "销售部(线下)": 91000.0, "餐饮部": 2244.0, "顾客关系部(线下)": 1045.0, "物流部": 3395.0}, {"日期": "2026-09-18", "周数": 38, "星期": "Fri", "商场OMNI": 113300.0, "销售部(线下)": 108000.0, "餐饮部": 2046.0, "顾客关系部(线下)": 1144.0, "物流部": 3500.0}, {"日期": "2026-09-19", "周数": 38, "星期": "Sat", "商场OMNI": 110000.0, "销售部(线下)": 90000.0, "餐饮部": 2310.0, "顾客关系部(线下)": 1078.0, "物流部": 3850.0}, {"日期": "2026-09-20", "周数": 38, "星期": "Sun", "商场OMNI": 115500.0, "销售部(线下)": 98000.0, "餐饮部": 1980.0, "顾客关系部(线下)": 1144.0, "物流部": 3780.0}, {"日期": "2026-09-21", "周数": 39, "星期": "Mon", "商场OMNI": 114400.0, "销售部(线下)": 102000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1056.0, "物流部": 3220.0}, {"日期": "2026-09-22", "周数": 39, "星期": "Tue", "商场OMNI": 104500.0, "销售部(线下)": 101000.0, "餐饮部": 2156.0, "顾客关系部(线下)": 1045.0, "物流部": 3150.0}, {"日期": "2026-09-23", "周数": 39, "星期": "Wed", "商场OMNI": 101200.0, "销售部(线下)": 102000.0, "餐饮部": 2134.0, "顾客关系部(线下)": 990.0, "物流部": 3640.0}, {"日期": "2026-09-24", "周数": 39, "星期": "Thu", "商场OMNI": 99000.0, "销售部(线下)": 105000.0, "餐饮部": 2200.0, "顾客关系部(线下)": 1210.0, "物流部": 3850.0}, {"日期": "2026-09-25", "周数": 39, "星期": "Fri", "商场OMNI": 113300.0, "销售部(线下)": 102000.0, "餐饮部": 2112.0, "顾客关系部(线下)": 1045.0, "物流部": 3150.0}, {"日期": "2026-09-26", "周数": 39, "星期": "Sat", "商场OMNI": 111100.0, "销售部(线下)": 93000.0, "餐饮部": 2266.0, "顾客关系部(线下)": 1199.0, "物流部": 3640.0}, {"日期": "2026-09-27", "周数": 39, "星期": "Sun", "商场OMNI": 103400.0, "销售部(线下)": 104000.0, "餐饮部": 2002.0, "顾客关系部(线下)": 1012.0, "物流部": 3290.0}, {"日期": "2026-09-28", "周数": 40, "星期": "Mon", "商场OMNI": 112200.0, "销售部(线下)": 100000.0, "餐饮部": 2376.0, "顾客关系部(线下)": 1067.0, "物流部": 3675.0}, {"日期": "2026-09-29", "周数": 40, "星期": "Tue", "商场OMNI": 113300.0, "销售部(线下)": 91000.0, "餐饮部": 2332.0, "顾客关系部(线下)": 1001.0, "物流部": 3745.0}, {"日期": "2026-09-30", "周数": 40, "星期": "Wed", "商场OMNI": 110000.0, "销售部(线下)": 100000.0, "餐饮部": 2112.0, "顾客关系部(线下)": 1056.0, "物流部": 3395.0}];

  const STORAGE_KEY_HOURS = 'suzhou_eff_module2_hours_data_v2';
  const STORAGE_KEY_DRIVER = 'suzhou_eff_module2_driver_data_v2';

  // State: objects keyed by 'YYYY-MM-DD' => { '商场': val, '销售部': val, ... }
  let hoursData = {};
  let driverData = {};

  // Date parsing helper
  function normalizeDate(val) {
    if (!val) return null;
    // 1. If Date object: add 12 hours to safely snap to the intended calendar day
    // This perfectly cancels out any timezone shift (e.g. UTC+8 shifting backwards by 8 hours to 16:00 previous day)
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return null;
      const snap = new Date(val.getTime() + 12 * 3600 * 1000);
      const y = snap.getUTCFullYear();
      const m = String(snap.getUTCMonth() + 1).padStart(2, '0');
      const d = String(snap.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    const str = String(val).trim();
    // 2. Formatted date string: YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD, YYYY年MM月DD日
    const match = str.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
    if (match) {
      return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
    }
    // 3. Formatted date string with month first: MM/DD/YYYY or M/D/YYYY
    const matchMDY = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
    if (matchMDY) {
      return `${matchMDY[3]}-${matchMDY[1].padStart(2, '0')}-${matchMDY[2].padStart(2, '0')}`;
    }
    // 4. Excel serial number (e.g. 46295 for 2026-09-30, 46296 for 2026-10-01)
    if (/^\d{4,5}(\.\d+)?$/.test(str)) {
      const num = parseFloat(str);
      const snap = new Date(Math.round((num - 25569) * 86400 * 1000) + 12 * 3600 * 1000);
      const y = snap.getUTCFullYear();
      const m = String(snap.getUTCMonth() + 1).padStart(2, '0');
      const d = String(snap.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return str;
  }

  
  // ==========================================
  // MODULE 2: 多时段效率对比分析 LOGIC
  // ==========================================
  
  // ==========================================
  // MODULE 2: 每周效率走势与波动分析 LOGIC
  // ==========================================
  let weeklyChart = null;
  const DEFAULT_DATE_WEEK_MAP = {"2026-09-01": 36, "2026-09-02": 36, "2026-09-03": 36, "2026-09-04": 36, "2026-09-05": 36, "2026-09-06": 36, "2026-09-07": 37, "2026-09-08": 37, "2026-09-09": 37, "2026-09-10": 37, "2026-09-11": 37, "2026-09-12": 37, "2026-09-13": 37, "2026-09-14": 38, "2026-09-15": 38, "2026-09-16": 38, "2026-09-17": 38, "2026-09-18": 38, "2026-09-19": 38, "2026-09-20": 38, "2026-09-21": 39, "2026-09-22": 39, "2026-09-23": 39, "2026-09-24": 39, "2026-09-25": 39, "2026-09-26": 39, "2026-09-27": 39, "2026-09-28": 40, "2026-09-29": 40, "2026-09-30": 40};
  const STORAGE_KEY_WEEK_MAP = 'suzhou_eff_date_week_map_v2';
  let dateWeekMap = {};

  function initDateWeekMap() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_WEEK_MAP);
      dateWeekMap = stored ? JSON.parse(stored) : DEFAULT_DATE_WEEK_MAP;
    } catch(e) {
      dateWeekMap = DEFAULT_DATE_WEEK_MAP;
    }
  }

  function initWeeklyChart() {
    const chartEl = document.getElementById('chartWeeklyTrend');
    if (chartEl && !weeklyChart) {
      weeklyChart = echarts.init(chartEl);
    }
  }

  function updateWeeklyEfficiencyAnalysis() {
    if (!weeklyChart) {
      initWeeklyChart();
    }
    const deptSelect = document.getElementById('weeklyDeptSelect');
    if (!deptSelect) return;
    const dept = deptSelect.value;
    const deptInfo = EFF_DEPTS.find(d => d.dept === dept);
    const target = deptInfo ? deptInfo.numericTarget : null;

    // Get all available dates that have data
    const allDates = getAllAvailableDates();
    if (allDates.length === 0) return;

    // Collect weeks present in data
    const weekMap = {};
    allDates.forEach(d => {
      let w = dateWeekMap[d];
      if (!w) {
        // Fallback compute ISO week
        const dt = new Date(d);
        if (isNaN(dt.getTime())) {
          w = 0;
        } else {
          dt.setHours(0, 0, 0, 0);
          const day = dt.getDay() || 7;
          dt.setDate(dt.getDate() + 4 - day);
          const yStart = new Date(dt.getFullYear(), 0, 1);
          w = Math.ceil((((dt - yStart) / 86400000) + 1) / 7);
        }
        dateWeekMap[d] = w;
      }
      if (!weekMap[w]) weekMap[w] = [];
      weekMap[w].push(d);
    });

    const sortedWeeks = Object.keys(weekMap).map(Number).sort((a, b) => a - b);
    if (sortedWeeks.length === 0) return;

    // Compute weekly data
    const weeklyDataList = [];
    sortedWeeks.forEach(w => {
      const datesInWeek = weekMap[w];
      let sumHours = 0;
      let sumDriver = 0;
      datesInWeek.forEach(d => {
        sumHours += Number(hoursData[d]?.[dept]) || 0;
        sumDriver += Number(driverData[d]?.[dept]) || 0;
      });
      const eff = sumHours > 0 ? (sumDriver / sumHours) : 0;
      weeklyDataList.push({
        week: w,
        daysCount: datesInWeek.length,
        driverSum: sumDriver,
        hoursSum: sumHours,
        eff: eff,
        dateRange: `${datesInWeek[0]} ~ ${datesInWeek[datesInWeek.length - 1]}`
      });
    });

    // Fluctuation statistics
    const effValues = weeklyDataList.map(item => item.eff);
    const totalEffSum = effValues.reduce((a, b) => a + b, 0);
    const avgEff = effValues.length > 0 ? (totalEffSum / effValues.length) : 0;
    const maxEff = Math.max(...effValues);
    const minEff = Math.min(...effValues);
    const maxItem = weeklyDataList.find(item => item.eff === maxEff) || {};
    const minItem = weeklyDataList.find(item => item.eff === minEff) || {};

    const fluctuationRange = maxEff - minEff;
    const fluctuationRate = avgEff > 0 ? ((fluctuationRange / avgEff) * 100) : 0;

    // Update KPI cards
    document.getElementById('weeklyAvgEff').textContent = avgEff.toFixed(2);
    document.getElementById('weeklyTotalWeeksTag').textContent = `共覆盖 ${weeklyDataList.length} 个自然周 (${weeklyDataList[0]?.week}周 ~ ${weeklyDataList[weeklyDataList.length - 1]?.week}周)`;

    document.getElementById('weeklyMaxEff').textContent = maxEff.toFixed(2);
    document.getElementById('weeklyMaxWeekDesc').textContent = `最高峰：第 ${maxItem.week || '--'} 周 (${maxItem.dateRange || ''})`;

    document.getElementById('weeklyMinEff').textContent = minEff.toFixed(2);
    document.getElementById('weeklyMinWeekDesc').textContent = `最低谷：第 ${minItem.week || '--'} 周 (${minItem.dateRange || ''})`;

    document.getElementById('weeklyFluctuationRange').textContent = `极差 ${fluctuationRange.toFixed(2)}`;
    document.getElementById('weeklyFluctuationRate').textContent = `波动幅度：±${(fluctuationRange / 2).toFixed(2)} (${fluctuationRate.toFixed(1)}%)`;

    // Render Weekly Data Table
    const tbody = document.getElementById('weeklyTableBody');
    if (tbody) {
      tbody.innerHTML = weeklyDataList.map((item, idx) => {
        let wowBadge = '<span class="text-slate-400 font-medium">基准周</span>';
        if (idx > 0) {
          const prevEff = weeklyDataList[idx - 1].eff;
          const wowDiff = item.eff - prevEff;
          const wowRate = prevEff > 0 ? ((wowDiff / prevEff) * 100) : 0;
          if (wowRate >= 0) {
            wowBadge = `<span class="badge bg-emerald-100 text-emerald-800">▲ +${wowRate.toFixed(1)}% (+${wowDiff.toFixed(1)})</span>`;
          } else {
            wowBadge = `<span class="badge bg-rose-100 text-rose-800">▼ ${wowRate.toFixed(1)}% (${wowDiff.toFixed(1)})</span>`;
          }
        }
        return `<tr class="hover:bg-slate-50 transition">
          <td class="px-4 py-3 font-bold text-slate-800">第 ${item.week} 周 <span class="text-[11px] text-slate-400 font-normal">(${item.dateRange})</span></td>
          <td class="px-4 py-3 text-center text-slate-600 font-medium">${item.daysCount} 天</td>
          <td class="px-4 py-3 text-right font-medium text-slate-700">${item.driverSum.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</td>
          <td class="px-4 py-3 text-right font-medium text-slate-700">${item.hoursSum.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</td>
          <td class="px-4 py-3 text-right font-extrabold text-emerald-700 text-base bg-emerald-50/40">${item.eff.toFixed(2)}</td>
          <td class="px-4 py-3 text-center">${wowBadge}</td>
        </tr>`;
      }).join('');
    }

    // Render Weekly ECharts Curve
    if (!weeklyChart) return;
    const categories = weeklyDataList.map(item => `第 ${item.week} 周`);
    const seriesData = weeklyDataList.map(item => ({
      value: Number(item.eff.toFixed(2)),
      week: item.week,
      driver: item.driverSum,
      hours: item.hoursSum,
      days: item.daysCount,
      dateRange: item.dateRange
    }));

    const markLineData = [
      {
        yAxis: Number(avgEff.toFixed(2)),
        name: '平均效率',
        lineStyle: { color: '#3b82f6', type: 'dashed', width: 2 },
        label: { show: true, position: 'insideEndTop', formatter: `均值: ${avgEff.toFixed(2)}`, color: '#2563eb', fontWeight: 'bold' }
      }
    ];

    if (target) {
      markLineData.push({
        yAxis: target,
        name: '目标效率',
        lineStyle: { color: '#f43f5e', type: 'dashed', width: 2 },
        label: { show: true, position: 'insideEndBottom', formatter: `目标: ${target}`, color: '#e11d48', fontWeight: 'bold' }
      });
    }

    weeklyChart.setOption({
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(15, 23, 42, 0.9)',
        borderColor: '#334155',
        textStyle: { color: '#f8fafc', fontSize: 12 },
        formatter: function(params) {
          if (!params || params.length === 0) return '';
          const p = params[0];
          const d = p.data;
          let tip = `<div class="font-bold pb-1 mb-1 border-b border-slate-700">${p.name} · 每周效率明细</div>`;
          tip += `<div class="py-0.5"><span class="text-emerald-400 font-bold text-sm">实际效率: ${d.value}</span></div>`;
          tip += `<div class="py-0.5 text-slate-300">关键驱动指标周累加: <b>${d.driver.toLocaleString()}</b></div>`;
          tip += `<div class="py-0.5 text-slate-300">工时周累加: <b>${d.hours.toLocaleString()} 小时</b> (${d.days} 天)</div>`;
          tip += `<div class="py-0.5 text-slate-400 text-[11px]">覆盖区间: ${d.dateRange}</div>`;
          return tip;
        }
      },
      grid: { left: 55, right: target ? 85 : 40, top: 35, bottom: 28 },
      xAxis: {
        type: 'category',
        data: categories,
        axisLine: { lineStyle: { color: '#cbd5e1' } },
        axisLabel: { color: '#475569', fontSize: 11, fontWeight: 'bold' }
      },
      yAxis: {
        type: 'value',
        scale: true,
        axisLabel: { color: '#64748b', fontSize: 11 },
        splitLine: { lineStyle: { color: '#f1f5f9' } }
      },
      series: [{
        name: '周实际效率',
        type: 'line',
        smooth: true,
        showSymbol: true,
        symbolSize: 8,
        itemStyle: { color: '#059669' },
        lineStyle: { width: 3.5, color: '#10b981' },
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: 'rgba(16, 185, 129, 0.28)' },
            { offset: 1, color: 'rgba(16, 185, 129, 0.01)' }
          ])
        },
        label: {
          show: true,
          position: 'top',
          color: '#065f46',
          fontWeight: 'bold',
          fontSize: 11,
          formatter: '{c}'
        },
        markPoint: {
          symbolSize: 45,
          data: [
            { type: 'max', name: '波峰最高', itemStyle: { color: '#059669' } },
            { type: 'min', name: '波谷最低', itemStyle: { color: '#e11d48' } }
          ],
          label: { fontSize: 10, fontWeight: 'bold' }
        },
        markLine: {
          symbol: ['none', 'none'],
          data: markLineData
        },
        data: seriesData
      }]
    }, true);
  }

  
  // ==========================================
  // MODULE 3: 一周七天工时占比分析 LOGIC
  // ==========================================
  let weekdayChart = null;
  const DEFAULT_DATE_WEEKDAY_MAP = {"2026-09-01": "Tue", "2026-09-02": "Wed", "2026-09-03": "Thu", "2026-09-04": "Fri", "2026-09-05": "Sat", "2026-09-06": "Sun", "2026-09-07": "Mon", "2026-09-08": "Tue", "2026-09-09": "Wed", "2026-09-10": "Thu", "2026-09-11": "Fri", "2026-09-12": "Sat", "2026-09-13": "Sun", "2026-09-14": "Mon", "2026-09-15": "Tue", "2026-09-16": "Wed", "2026-09-17": "Thu", "2026-09-18": "Fri", "2026-09-19": "Sat", "2026-09-20": "Sun", "2026-09-21": "Mon", "2026-09-22": "Tue", "2026-09-23": "Wed", "2026-09-24": "Thu", "2026-09-25": "Fri", "2026-09-26": "Sat", "2026-09-27": "Sun", "2026-09-28": "Mon", "2026-09-29": "Tue", "2026-09-30": "Wed"};
  const STORAGE_KEY_WEEKDAY_MAP = 'suzhou_eff_date_weekday_map_v2';
  let dateWeekdayMap = {};

  const WEEKDAYS_ORDER = [
    { key: 'Sun', cn: '周日', full: '周日 (Sun)', index: 0 },
    { key: 'Mon', cn: '周一', full: '周一 (Mon)', index: 1 },
    { key: 'Tue', cn: '周二', full: '周二 (Tue)', index: 2 },
    { key: 'Wed', cn: '周三', full: '周三 (Wed)', index: 3 },
    { key: 'Thu', cn: '周四', full: '周四 (Thu)', index: 4 },
    { key: 'Fri', cn: '周五', full: '周五 (Fri)', index: 5 },
    { key: 'Sat', cn: '周六', full: '周六 (Sat)', index: 6 }
  ];

  function matchWeekdayIndex(str, dateStr) {
    const getSafeDay = (ds) => {
      const d = new Date(ds);
      const day = d.getDay();
      return isNaN(day) ? 0 : day;
    };
    if (!str && dateStr) return getSafeDay(dateStr);
    const s = String(str).trim().toLowerCase();
    if (s.includes('sun') || s.includes('日') || s === '7' || s === '0') return 0;
    if (s.includes('mon') || s.includes('一') || s === '1') return 1;
    if (s.includes('tue') || s.includes('二') || s === '2') return 2;
    if (s.includes('wed') || s.includes('三') || s === '3') return 3;
    if (s.includes('thu') || s.includes('四') || s === '4') return 4;
    if (s.includes('fri') || s.includes('五') || s === '5') return 5;
    if (s.includes('sat') || s.includes('六') || s === '6') return 6;
    if (dateStr) return getSafeDay(dateStr);
    return 0;
  }

  function initDateWeekdayMap() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_WEEKDAY_MAP);
      dateWeekdayMap = stored ? JSON.parse(stored) : DEFAULT_DATE_WEEKDAY_MAP;
    } catch(e) {
      dateWeekdayMap = DEFAULT_DATE_WEEKDAY_MAP;
    }
  }

  function initWeekdayChart() {
    const chartEl = document.getElementById('chartWeekdayDistribution');
    if (chartEl && !weekdayChart) {
      weekdayChart = echarts.init(chartEl);
    }
  }

  function initWeekdayDatePickers(forceReset = false) {
    const dates = getAllAvailableDates();
    if (dates.length === 0) return;
    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];

    const s = document.getElementById('weekdayStartDate');
    const e = document.getElementById('weekdayEndDate');
    if (s) {
      s.min = minDate; s.max = maxDate;
      if (forceReset || !s.value) s.value = minDate;
    }
    if (e) {
      e.min = minDate; e.max = maxDate;
      if (forceReset || !e.value || e.value < maxDate) e.value = maxDate;
    }
  }

  function setWeekdayRangeQuick(type) {
    const dates = getAllAvailableDates();
    if (dates.length === 0) return;
    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];
    const maxD = new Date(maxDate);

    const s = document.getElementById('weekdayStartDate');
    const e = document.getElementById('weekdayEndDate');

    if (type === 'all') {
      s.value = minDate; e.value = maxDate;
    } else if (type === 'latest14') {
      const d14 = new Date(maxD);
      d14.setDate(d14.getDate() - 13);
      const d14Str = d14.toISOString().split('T')[0];
      s.value = d14Str < minDate ? minDate : d14Str;
      e.value = maxDate;
    } else if (type === 'latest30') {
      const d30 = new Date(maxD);
      d30.setDate(d30.getDate() - 29);
      const d30Str = d30.toISOString().split('T')[0];
      s.value = d30Str < minDate ? minDate : d30Str;
      e.value = maxDate;
    }
    updateWeekdayDistributionAnalysis();
  }

  function goToModule3() {
    switchPage('module2');
    setTimeout(() => {
      const el = document.getElementById('section-weekday-distribution');
      if (el) {
        el.scrollIntoView({ behavior: 'smooth' });
      }
    }, 80);
  }

  function updateWeekdayDistributionAnalysis() {
    if (!weekdayChart) {
      initWeekdayChart();
    }
    const deptSelect = document.getElementById('weekdayDeptSelect');
    if (!deptSelect) return;
    const dept = deptSelect.value;

    const s = document.getElementById('weekdayStartDate')?.value;
    const e = document.getElementById('weekdayEndDate')?.value;

    const allDates = getAllAvailableDates();
    const rangeDates = allDates.filter(d => (!s || d >= s) && (!e || d <= e));

    if (rangeDates.length === 0) return;

    // Group dates by weekday index 0..6 (Sun..Sat)
    const weekdayBuckets = Array.from({ length: 7 }, () => ({
      dates: [],
      totalHours: 0,
      totalSales: 0,
      avgHours: 0,
      avgSales: 0,
      ratio: 0
    }));

    rangeDates.forEach(d => {
      const rawWeekday = dateWeekdayMap[d];
      const wIdx = matchWeekdayIndex(rawWeekday, d);
      if (weekdayBuckets[wIdx]) {
        const h = Number(hoursData[d]?.[dept]) || 0;
        const sl = Number(driverData[d]?.[dept]) || 0;
        weekdayBuckets[wIdx].dates.push(d);
        weekdayBuckets[wIdx].totalHours += h;
        weekdayBuckets[wIdx].totalSales += sl;
      }
    });

    // Compute averages
    let sumOfAvgs = 0;
    weekdayBuckets.forEach((b, idx) => {
      const count = b.dates.length;
      b.avgHours = count > 0 ? (b.totalHours / count) : 0;
      b.avgSales = count > 0 ? (b.totalSales / count) : 0;
      sumOfAvgs += b.avgHours;
    });

    // Compute percentage
    weekdayBuckets.forEach(b => {
      b.ratio = sumOfAvgs > 0 ? ((b.avgHours / sumOfAvgs) * 100) : 0;
    });

    // Update status text
    const statusEl = document.getElementById('weekdayRangeStatus');
    if (statusEl) {
      const daySummary = WEEKDAYS_ORDER.map((w, i) => `${w.cn}${weekdayBuckets[i].dates.length}天`).join('、');
      statusEl.innerHTML = `已匹配 <b>${rangeDates.length}</b> 天数据 (${daySummary})`;
    }

    // Update KPI Cards
    document.getElementById('weekdaySumAvgHours').textContent = sumOfAvgs.toFixed(1) + ' 小时';

    let maxIdx = 0, minIdx = 0;
    for (let i = 1; i < 7; i++) {
      if (weekdayBuckets[i].ratio > weekdayBuckets[maxIdx].ratio) maxIdx = i;
      if (weekdayBuckets[i].ratio < weekdayBuckets[minIdx].ratio) minIdx = i;
    }

    const peakW = WEEKDAYS_ORDER[maxIdx];
    const valleyW = WEEKDAYS_ORDER[minIdx];

    document.getElementById('weekdayPeakDay').textContent = `${peakW.full} (${weekdayBuckets[maxIdx].ratio.toFixed(2)}%)`;
    document.getElementById('weekdayPeakDetail').textContent = `匹配 ${weekdayBuckets[maxIdx].dates.length} 天 | 日均: ${weekdayBuckets[maxIdx].avgHours.toFixed(1)} 小时`;

    document.getElementById('weekdayValleyDay').textContent = `${valleyW.full} (${weekdayBuckets[minIdx].ratio.toFixed(2)}%)`;
    document.getElementById('weekdayValleyDetail').textContent = `匹配 ${weekdayBuckets[minIdx].dates.length} 天 | 日均: ${weekdayBuckets[minIdx].avgHours.toFixed(1)} 小时`;

    const weekendRatio = weekdayBuckets[0].ratio + weekdayBuckets[6].ratio;
    const weekdayRatio = 100 - weekendRatio;

    document.getElementById('weekendRatioVal').textContent = weekendRatio.toFixed(1) + '%';
    document.getElementById('weekdayRatioDetail').textContent = `工作日(周一~周五)：${weekdayRatio.toFixed(1)}%`;

    // Render Table
    const tbody = document.getElementById('weekdayTableBody');
    if (tbody) {
      tbody.innerHTML = WEEKDAYS_ORDER.map((w, i) => {
        const b = weekdayBuckets[i];
        return `<tr class="hover:bg-amber-50/30 transition">
          <td class="px-4 py-3 font-bold text-slate-800">${w.full}</td>
          <td class="px-4 py-3 text-center text-slate-600 font-medium">${b.dates.length} 天</td>
          <td class="px-4 py-3 text-right font-medium text-slate-700">${b.totalHours.toLocaleString(undefined, {maximumFractionDigits: 1})}</td>
          <td class="px-4 py-3 text-right font-semibold text-emerald-700">${b.avgHours.toFixed(2)}</td>
          <td class="px-4 py-3 text-right font-extrabold text-amber-700 text-base bg-amber-50/40">${b.ratio.toFixed(2)}%</td>
        </tr>`;
      }).join('');
    }

    // Render Dual Chart
    if (!weekdayChart) return;
    const categories = WEEKDAYS_ORDER.map(w => w.cn);

    weekdayChart.setOption({
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(15, 23, 42, 0.9)',
        borderColor: '#334155',
        textStyle: { color: '#f8fafc', fontSize: 12 },
        formatter: function(params) {
          if (!params || params.length === 0) return '';
          let tip = `<div class="font-bold pb-1 mb-1 border-b border-slate-700">${WEEKDAYS_ORDER[params[0].dataIndex].full} · 趋势对比</div>`;
          params.forEach(p => {
             const dot = `<span style="display:inline-block;margin-right:6px;border-radius:10px;width:9px;height:9px;background-color:${p.color};"></span>`;
             if (p.seriesName === '工时占比') {
               tip += `<div class="py-0.5">${dot} <span class="text-amber-400 font-bold">${p.seriesName}: ${p.value}%</span></div>`;
             } else {
               tip += `<div class="py-0.5">${dot} <span class="text-sky-400 font-bold">${p.seriesName}: ${p.value.toLocaleString()}</span></div>`;
             }
          });
          return tip;
        }
      },
      legend: {
        data: ['工时占比', '日均销售'],
        textStyle: { color: '#64748b', fontSize: 11 },
        top: 0
      },
      grid: { left: 55, right: 55, top: 40, bottom: 28 },
      xAxis: {
        type: 'category',
        data: categories,
        axisLine: { lineStyle: { color: '#cbd5e1' } },
        axisLabel: { color: '#475569', fontSize: 12, fontWeight: 'bold' }
      },
      yAxis: [
        {
          type: 'value',
          name: '占比',
          scale: true,
          boundaryGap: ['10%', '10%'],
          axisLabel: { formatter: '{value}%', color: '#f59e0b' },
          splitLine: { lineStyle: { color: '#f1f5f9' } }
        },
        {
          type: 'value',
          name: '销售',
          axisLabel: { 
            formatter: (val) => val >= 1000 ? (val/1000).toFixed(0) + 'k' : val,
            color: '#0ea5e9' 
          },
          splitLine: { show: false }
        }
      ],
      series: [
        {
          name: '日均销售',
          type: 'bar',
          yAxisIndex: 1,
          barWidth: '35%',
          itemStyle: {
            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
              { offset: 0, color: '#38bdf8' },
              { offset: 1, color: '#0ea5e9' }
            ]),
            borderRadius: [4, 4, 0, 0]
          },
          data: weekdayBuckets.map(b => Number(b.avgSales.toFixed(0)))
        },
        {
          name: '工时占比',
          type: 'line',
          yAxisIndex: 0,
          smooth: true,
          symbolSize: 8,
          itemStyle: { color: '#f59e0b' },
          lineStyle: { width: 3, color: '#f59e0b' },
          label: {
            show: true,
            position: 'top',
            formatter: '{c}%',
            fontSize: 10,
            color: '#b45309'
          },
          data: weekdayBuckets.map(b => Number(b.ratio.toFixed(2)))
        }
      ]
    }, true);
  }

  let compareChart = null;

  function initCompareChart() {
    const chartEl = document.getElementById('chartEfficiencyTrend');
    if (chartEl && !compareChart) {
      compareChart = echarts.init(chartEl);
    }
  }

  function initComparisonDatePickers(forceReset = false) {
    const dates = getAllAvailableDates();
    if (dates.length === 0) return;
    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];

    ['compareStart1', 'compareEnd1', 'compareStart2', 'compareEnd2'].forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        el.min = minDate;
        el.max = maxDate;
      }
    });

    const s1 = document.getElementById('compareStart1');
    const e1 = document.getElementById('compareEnd1');
    const s2 = document.getElementById('compareStart2');
    const e2 = document.getElementById('compareEnd2');

    if (forceReset || (s1 && !s1.value)) {
      if (dates.length >= 14) {
        const mid = Math.floor(dates.length / 2);
        s1.value = dates[0];
        e1.value = dates[mid - 1];
        s2.value = dates[mid];
        e2.value = maxDate;
      } else {
        s1.value = dates[0];
        e1.value = dates[Math.min(6, dates.length - 1)];
        s2.value = dates[0];
        e2.value = maxDate;
      }
    } else if (e2 && e2.value < maxDate) {
      e2.value = maxDate;
    }
  }

  function updateEfficiencyComparison() {
    if (!compareChart) {
      initCompareChart();
    }
    const deptSelect = document.getElementById('compareDeptSelect');
    if (!deptSelect) return;
    const dept = deptSelect.value;
    const deptInfo = EFF_DEPTS.find(d => d.dept === dept);
    const target = deptInfo ? deptInfo.numericTarget : null;

    const s1 = document.getElementById('compareStart1')?.value;
    const e1 = document.getElementById('compareEnd1')?.value;
    const s2 = document.getElementById('compareStart2')?.value;
    const e2 = document.getElementById('compareEnd2')?.value;

    const allDates = getAllAvailableDates();
    const dates1 = allDates.filter(d => (!s1 || d >= s1) && (!e1 || d <= e1));
    const dates2 = allDates.filter(d => (!s2 || d >= s2) && (!e2 || d <= e2));

    const tag1 = document.getElementById('period1DaysTag');
    if (tag1) tag1.textContent = `${dates1.length} 天数据 (${s1||''}~${e1||''})`;
    const tag2 = document.getElementById('period2DaysTag');
    if (tag2) tag2.textContent = `${dates2.length} 天数据 (${s2||''}~${e2||''})`;

    let sumDr1 = 0, sumH1 = 0;
    const p1Daily = dates1.map(d => {
      const h = Number(hoursData[d]?.[dept]) || 0;
      const dr = Number(driverData[d]?.[dept]) || 0;
      sumH1 += h;
      sumDr1 += dr;
      const eff = h > 0 ? (dr / h) : 0;
      return { date: d, hours: h, driver: dr, eff: eff };
    });
    const compEff1 = sumH1 > 0 ? (sumDr1 / sumH1) : 0;

    let sumDr2 = 0, sumH2 = 0;
    const p2Daily = dates2.map(d => {
      const h = Number(hoursData[d]?.[dept]) || 0;
      const dr = Number(driverData[d]?.[dept]) || 0;
      sumH2 += h;
      sumDr2 += dr;
      const eff = h > 0 ? (dr / h) : 0;
      return { date: d, hours: h, driver: dr, eff: eff };
    });
    const compEff2 = sumH2 > 0 ? (sumDr2 / sumH2) : 0;

    const diff = compEff2 - compEff1;
    const rate = compEff1 > 0 ? ((diff / compEff1) * 100) : 0;

    // Update Numeric Cards
    const v1El = document.getElementById('compEffVal1');
    const d1El = document.getElementById('compDetail1');
    if (v1El) v1El.textContent = compEff1 > 0 ? compEff1.toFixed(2) : '—';
    if (d1El) d1El.textContent = `驱动: ${sumDr1.toLocaleString(undefined, {maximumFractionDigits: 1})} | 工时: ${sumH1.toFixed(1)}`;

    const v2El = document.getElementById('compEffVal2');
    const d2El = document.getElementById('compDetail2');
    if (v2El) v2El.textContent = compEff2 > 0 ? compEff2.toFixed(2) : '—';
    if (d2El) d2El.textContent = `驱动: ${sumDr2.toLocaleString(undefined, {maximumFractionDigits: 1})} | 工时: ${sumH2.toFixed(1)}`;

    const diffEl = document.getElementById('compDiffVal');
    const diffDesc = document.getElementById('compDiffDesc');
    if (diffEl) {
      diffEl.textContent = (diff >= 0 ? '+' : '') + diff.toFixed(2);
      diffEl.className = 'text-2xl font-black mt-1 ' + (diff >= 0 ? 'text-emerald-600' : 'text-rose-600');
    }
    if (diffDesc) diffDesc.textContent = diff >= 0 ? '效率提高' : '效率回落';

    const rateEl = document.getElementById('compRateVal');
    const rateBadge = document.getElementById('compRateBadge');
    if (rateEl) {
      rateEl.textContent = (rate >= 0 ? '+' : '') + rate.toFixed(1) + '%';
      rateEl.className = 'text-2xl font-black mt-1 ' + (rate >= 0 ? 'text-emerald-600' : 'text-rose-600');
    }
    if (rateBadge) rateBadge.textContent = rate >= 0 ? '较时段1增长' : '较时段1下降';

    // Chart rendering
    if (!compareChart) return;
    const maxDays = Math.max(dates1.length, dates2.length, 1);
    const categories = Array.from({ length: maxDays }, (_, i) => `第 ${i + 1} 天`);

    const s1Data = [];
    const s2Data = [];

    for (let i = 0; i < maxDays; i++) {
      if (i < dates1.length) {
        s1Data.push({
          value: Number(p1Daily[i].eff.toFixed(2)),
          dateStr: p1Daily[i].date,
          driver: p1Daily[i].driver,
          hours: p1Daily[i].hours
        });
      } else {
        s1Data.push(null);
      }

      if (i < dates2.length) {
        s2Data.push({
          value: Number(p2Daily[i].eff.toFixed(2)),
          dateStr: p2Daily[i].date,
          driver: p2Daily[i].driver,
          hours: p2Daily[i].hours
        });
      } else {
        s2Data.push(null);
      }
    }

    const markLine = target ? {
      symbol: ['none', 'none'],
      lineStyle: { color: '#f43f5e', type: 'dashed', width: 2 },
      label: { show: true, position: 'end', formatter: `目标: ${target}`, color: '#e11d48', fontWeight: 'bold' },
      data: [{ yAxis: target, name: '部门目标' }]
    } : undefined;

    compareChart.setOption({
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(15, 23, 42, 0.9)',
        borderColor: '#334155',
        textStyle: { color: '#f8fafc', fontSize: 12 },
        formatter: function(params) {
          if (!params || params.length === 0) return '';
          let tip = `<div class="font-bold pb-1 mb-1 border-b border-slate-700">${params[0].name} 效率比对</div>`;
          params.forEach(p => {
            if (p.data && p.data.value !== null && p.data.value !== undefined) {
              const colorDot = `<span style="display:inline-block;margin-right:6px;border-radius:10px;width:9px;height:9px;background-color:${p.color};"></span>`;
              tip += `<div class="py-0.5">${colorDot} <b>${p.seriesName}</b> [${p.data.dateStr}]: <span class="font-bold text-amber-300">${p.data.value}</span> (指标:${p.data.driver.toLocaleString()}, 工时:${p.data.hours})</div>`;
            }
          });
          return tip;
        }
      },
      legend: { show: false },
      grid: { left: 50, right: target ? 80 : 30, top: 30, bottom: 28 },
      xAxis: {
        type: 'category',
        data: categories,
        axisLine: { lineStyle: { color: '#cbd5e1' } },
        axisLabel: { color: '#64748b', fontSize: 11 }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: '#64748b', fontSize: 11 },
        splitLine: { lineStyle: { color: '#f1f5f9' } }
      },
      series: [
        {
          name: '时段 1',
          type: 'line',
          smooth: true,
          showSymbol: true,
          symbolSize: 6,
          itemStyle: { color: '#2563eb' },
          lineStyle: { width: 3, color: '#2563eb' },
          areaStyle: {
            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
              { offset: 0, color: 'rgba(37, 99, 235, 0.22)' },
              { offset: 1, color: 'rgba(37, 99, 235, 0.01)' }
            ])
          },
          data: s1Data
        },
        {
          name: '时段 2',
          type: 'line',
          smooth: true,
          showSymbol: true,
          symbolSize: 6,
          itemStyle: { color: '#f59e0b' },
          lineStyle: { width: 3, color: '#f59e0b' },
          areaStyle: {
            color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
              { offset: 0, color: 'rgba(245, 158, 11, 0.22)' },
              { offset: 1, color: 'rgba(245, 158, 11, 0.01)' }
            ])
          },
          markLine: markLine,
          data: s2Data
        }
      ]
    }, true);
  }

  async function initModule2Data() {
    console.log("Initializing persistent data storage...");
    try {
      const hStored = await EfficiencyStorage.get(STORAGE_KEY_HOURS);
      hoursData = hStored ? hStored : JSON.parse(JSON.stringify(DEFAULT_HOURS_DATA));
    } catch(e) {
      console.error("Storage load error (Hours):", e);
      hoursData = JSON.parse(JSON.stringify(DEFAULT_HOURS_DATA));
    }

    try {
      const dStored = await EfficiencyStorage.get(STORAGE_KEY_DRIVER);
      driverData = dStored ? dStored : JSON.parse(JSON.stringify(DEFAULT_DRIVER_DATA));
    } catch(e) {
      console.error("Storage load error (Driver):", e);
      driverData = JSON.parse(JSON.stringify(DEFAULT_DRIVER_DATA));
    }

    try {
      const wStored = await EfficiencyStorage.get(STORAGE_KEY_WEEKDAY_MAP);
      if (wStored) {
        Object.assign(dateWeekdayMap, wStored);
      }
    } catch(e) {}

    updateUploadStatusIndicators();
    initDatePickers();
    applyDateFilter();
    initDateWeekMap();
    initCompareChart();
    initComparisonDatePickers();
    updateEfficiencyComparison();
    initWeeklyChart();
    updateWeeklyEfficiencyAnalysis();
    initWeekdayChart();
    initWeekdayDatePickers();
    updateWeekdayDistributionAnalysis();
  }

  function updateUploadStatusIndicators() {
    const hDates = Object.keys(hoursData).filter(d => Object.keys(hoursData[d] || {}).length > 0);
    const dDates = Object.keys(driverData).filter(d => Object.keys(driverData[d] || {}).length > 0);

    const hStatusEl = document.getElementById('hoursStatus');
    const hDaysEl = document.getElementById('hoursDaysCount');
    if (hDates.length > 0) {
      hStatusEl.textContent = `已载入 ${hDates.length} 天工时记录`;
      hDaysEl.textContent = `${hDates[0]} ~ ${hDates[hDates.length - 1]}`;
    } else {
      hStatusEl.textContent = '暂无工时数据，请上传表格';
      hDaysEl.textContent = '';
    }

    const dStatusEl = document.getElementById('driverStatus');
    const dDaysEl = document.getElementById('driverDaysCount');
    if (dDates.length > 0) {
      dStatusEl.textContent = `已载入 ${dDates.length} 天关键驱动记录`;
      dDaysEl.textContent = `${dDates[0]} ~ ${dDates[dDates.length - 1]}`;
    } else {
      dStatusEl.textContent = '暂无关键驱动数据，请上传表格';
      dDaysEl.textContent = '';
    }
  }

  function getAllAvailableDates() {
    const set = new Set([...Object.keys(hoursData), ...Object.keys(driverData)]);
    return Array.from(set).filter(Boolean).sort();
  }

  function initDatePickers(forceReset = false) {
    const dates = getAllAvailableDates();
    if (dates.length === 0) return;

    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];

    const startInput = document.getElementById('effStartDate');
    const endInput = document.getElementById('effEndDate');

    startInput.min = minDate;
    startInput.max = maxDate;
    endInput.min = minDate;
    endInput.max = maxDate;

    if (forceReset || !startInput.value || !endInput.value || endInput.value < maxDate) {
      startInput.value = minDate;
      endInput.value = maxDate;
    }
  }

  function setDateRangeQuick(type) {
    const dates = getAllAvailableDates();
    if (dates.length === 0) return;

    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];
    const maxD = new Date(maxDate);

    const startInput = document.getElementById('effStartDate');
    const endInput = document.getElementById('effEndDate');

    if (type === 'all') {
      startInput.value = minDate;
      endInput.value = maxDate;
    } else if (type === 'latest7') {
      const d7 = new Date(maxD);
      d7.setDate(d7.getDate() - 6);
      const d7Str = d7.toISOString().split('T')[0];
      startInput.value = d7Str < minDate ? minDate : d7Str;
      endInput.value = maxDate;
    } else if (type === 'latest30') {
      const d30 = new Date(maxD);
      d30.setDate(d30.getDate() - 29);
      const d30Str = d30.toISOString().split('T')[0];
      startInput.value = d30Str < minDate ? minDate : d30Str;
      endInput.value = maxDate;
    }
    applyDateFilter();
  }

  // Parse Matrix format table (as in 人员效率工时.xlsx or 人员效率关键驱动.xlsx)
  function parseMatrixExcel(workbook) {
    const sheetName = workbook.SheetNames[0];
    const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, dateNF: 'yyyy-mm-dd', defval: null });
    const resultDict = {};

    rawRows.forEach(row => {
      // Find date column
      let dateVal = null;
      Object.keys(row).forEach(k => {
        const keyTrim = k.trim();
        if (keyTrim.includes('日期') || keyTrim.toLowerCase().includes('date')) {
          dateVal = normalizeDate(row[k]);
        }
      });

      if (!dateVal) return;
      if (!resultDict[dateVal]) resultDict[dateVal] = {};

      // Check week column and weekday column
      Object.keys(row).forEach(k => {
        const kt = k.trim();
        if (kt.includes('周数') || kt.includes('周') || kt.toLowerCase().includes('week')) {
          const num = parseInt(String(row[k]).replace(/[^0-9]/g, ''));
          if (!isNaN(num) && num > 0) {
            dateWeekMap[dateVal] = num;
          }
        }
        if (kt.includes('星期') || kt.toLowerCase().includes('weekday') || kt.toLowerCase().includes('day')) {
          const val = String(row[k]).trim();
          if (val) {
            dateWeekdayMap[dateVal] = val;
          }
        }
      });

      // Match department columns
      EFF_DEPTS.forEach(d => {
        // Robust matching for department names
        const matchCol = Object.keys(row).find(k => {
          const kt = k.trim();
          // Precise match
          if (kt === d.dept) return true;
          // Partial match for common variations
          const cleanD = d.dept.replace(/\(线下\)/g, '').replace(/OMNI/g, '').trim();
          return kt.includes(cleanD) || cleanD.includes(kt);
        });

        if (matchCol && row[matchCol] !== null && row[matchCol] !== undefined && row[matchCol] !== '') {
          const num = parseFloat(String(row[matchCol]).replace(/,/g, ''));
          resultDict[dateVal][d.dept] = isNaN(num) ? 0 : num;
        } else {
          resultDict[dateVal][d.dept] = 0;
        }
      });
    });

    return resultDict;
  }

  async function handleHoursUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const parsed = parseMatrixExcel(workbook);
        const count = Object.keys(parsed).length;

        if (count === 0) {
          alert('未能识别到有效的日期与部门工时数据，请检查表格是否包含【日期】与部门列！');
          return;
        }

        // Merge logic: incrementally update the existing object
        Object.assign(hoursData, parsed);

        // Save to persistent storage
        await EfficiencyStorage.set(STORAGE_KEY_HOURS, hoursData);
        await EfficiencyStorage.set(STORAGE_KEY_WEEKDAY_MAP, dateWeekdayMap);

        updateUploadStatusIndicators();
        initDatePickers(true);
        applyDateFilter();
        initComparisonDatePickers(true);
        updateEfficiencyComparison();
        updateWeeklyEfficiencyAnalysis();
        initWeekdayDatePickers(true);
        updateWeekdayDistributionAnalysis();

        alert(`成功上传并合并数据！系统当前累计存有 ${Object.keys(hoursData).length} 天的工时数据。`);
      } catch(err) {
        console.error("Upload error:", err);
        alert('解析工时表格失败：' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function handleDriverUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const parsed = parseMatrixExcel(workbook);
        const count = Object.keys(parsed).length;

        if (count === 0) {
          alert('未能识别到有效的日期与部门关键驱动指标，请检查表格是否包含【日期】与部门列！');
          return;
        }

        // Merge logic
        Object.assign(driverData, parsed);

        // Save to persistent storage
        await EfficiencyStorage.set(STORAGE_KEY_DRIVER, driverData);
        await EfficiencyStorage.set(STORAGE_KEY_WEEKDAY_MAP, dateWeekdayMap);

        updateUploadStatusIndicators();
        initDatePickers(true);
        applyDateFilter();
        initComparisonDatePickers(true);
        updateEfficiencyComparison();
        updateWeeklyEfficiencyAnalysis();
        initWeekdayDatePickers(true);
        updateWeekdayDistributionAnalysis();

        alert(`成功上传并合并数据！系统当前累计存有 ${Object.keys(driverData).length} 天的驱动指标数据。`);
      } catch(err) {
        console.error("Upload error:", err);
        alert('解析驱动指标表格失败：' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function resetModule2Defaults() {
    if (confirm('是否确定重置为初始示范数据？')) {
      hoursData = DEFAULT_HOURS_DATA;
      driverData = DEFAULT_DRIVER_DATA;
      localStorage.removeItem(STORAGE_KEY_HOURS);
      localStorage.removeItem(STORAGE_KEY_DRIVER);
      updateUploadStatusIndicators();
      initDatePickers();
      applyDateFilter();
    }
  }

  function downloadHoursTemplate() {
    const ws = XLSX.utils.json_to_sheet(DEFAULT_HOURS_ROWS);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "人员效率工时");
    XLSX.writeFile(wb, "人员效率工时.xlsx");
  }

  function downloadDriverTemplate() {
    const ws = XLSX.utils.json_to_sheet(DEFAULT_DRIVER_ROWS);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "每日人员效率数据");
    XLSX.writeFile(wb, "人员效率关键驱动.xlsx");
  }

  // Calculate efficiency based on calendar range
  function applyDateFilter() {
    const startDate = document.getElementById('effStartDate').value;
    const endDate = document.getElementById('effEndDate').value;

    const allDates = getAllAvailableDates();
    const activeDates = allDates.filter(d => {
      if (startDate && d < startDate) return false;
      if (endDate && d > endDate) return false;
      return true;
    });

    document.getElementById('effStatusText').innerHTML = 
      `统计区间：<span class="font-bold text-indigo-700">${startDate || '起始'}</span> 至 <span class="font-bold text-indigo-700">${endDate || '截止'}</span>（共 <b>${activeDates.length}</b> 天有效数据）`;

    // Aggregate by department
    const agg = {};
    EFF_DEPTS.forEach(d => {
      agg[d.dept] = { driverSum: 0, hoursSum: 0, dept: d.dept, target: d.target, numericTarget: d.numericTarget };
    });

    activeDates.forEach(date => {
      const hRow = hoursData[date] || {};
      const dRow = driverData[date] || {};

      EFF_DEPTS.forEach(d => {
        if (hRow[d.dept] !== undefined && hRow[d.dept] !== null) {
          agg[d.dept].hoursSum += Number(hRow[d.dept]) || 0;
        }
        if (dRow[d.dept] !== undefined && dRow[d.dept] !== null) {
          agg[d.dept].driverSum += Number(dRow[d.dept]) || 0;
        }
      });
    });

    // Render table
    const tbody = document.getElementById('effTableBody');
    let totalHoursAll = 0;
    let totalDriverAll = 0;
    let storeActualEff = 0;

    tbody.innerHTML = EFF_DEPTS.map(d => {
      const item = agg[d.dept];
      const driver = item.driverSum;
      const hours = item.hoursSum;
      totalHoursAll += hours;
      totalDriverAll += driver;

      let eff = hours > 0 ? (driver / hours) : 0;
      let effDisplay = hours > 0 ? eff.toFixed(2) : '—';
      if (d.dept === '商场') storeActualEff = eff;

      let compareBadge = '<span class="text-slate-400 text-xs">—</span>';
      if (d.numericTarget && hours > 0) {
        const achieveRate = (eff / d.numericTarget) * 100;
        if (achieveRate >= 100) {
          compareBadge = `<span class="badge bg-emerald-100 text-emerald-800">达成 ${achieveRate.toFixed(1)}%</span>`;
        } else if (achieveRate >= 95) {
          compareBadge = `<span class="badge bg-amber-100 text-amber-800">达成 ${achieveRate.toFixed(1)}%</span>`;
        } else {
          compareBadge = `<span class="badge bg-rose-100 text-rose-800">达成 ${achieveRate.toFixed(1)}%</span>`;
        }
      }

      return `<tr class="hover:bg-indigo-50/40 transition">
        <td class="px-5 py-3.5 font-bold text-slate-800">${d.dept}</td>
        <td class="px-5 py-3.5 text-center font-semibold text-slate-600 bg-slate-50">${d.target}</td>
        <td class="px-5 py-3.5 text-right font-medium text-slate-700">${driver.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</td>
        <td class="px-5 py-3.5 text-right font-medium text-slate-700">${hours.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</td>
        <td class="px-5 py-3.5 text-right font-extrabold text-indigo-700 text-base bg-indigo-50/50">${effDisplay}</td>
        <td class="px-5 py-3.5 text-center">${compareBadge}</td>
      </tr>`;
    }).join('');

    // Summary cards removed as requested
  }


  // ==========================================
  // MODULE 1: 员工工时总体情况 LOGIC
  // ==========================================
  const MONTHLY_ROWS = ["线下（商场）", "餐饮"];
  const MONTHS = ["2026年9月", "2026年10月", "2026年11月", "2026年12月", "2027年1月", "2027年2月", "2027年3月", "2027年4月", "2027年5月", "2027年6月", "2027年7月", "2027年8月"];
  const MONTHLY_KEY = 'suzhou_efficiency_dashboard_monthly_v1';
  let monthlyStore = {};

  function renderMonthOptions() {
    const sel = document.getElementById('monthSelect');
    sel.innerHTML = MONTHS.map(m => `<option value="${m}">${m}</option>`).join('');
  }

  function renderMonthlyTable() {
    document.getElementById('monthlyBody').innerHTML = MONTHLY_ROWS.map((dept, i) => `<tr class="hover:bg-slate-50">
      <td class="px-4 py-3 font-bold fixed-cell">${dept}</td>
      <td class="px-4 py-3"><input id="mSalesTarget-${i}" class="table-input" type="number" min="0" placeholder="可填" oninput="updateMonthly();autoSaveMonthly()"></td>
      <td class="px-4 py-3"><input id="mSalesActual-${i}" class="table-input" type="number" min="0" placeholder="可填" oninput="updateMonthly();autoSaveMonthly()"></td>
      <td id="mSalesRate-${i}" class="px-4 py-3 text-center font-bold text-blue-700 fixed-cell"></td>
      <td class="px-4 py-3"><input id="mLaborTarget-${i}" class="table-input" type="number" min="0" placeholder="可填" oninput="updateMonthly();autoSaveMonthly()"></td>
      <td class="px-4 py-3"><input id="mLaborActual-${i}" class="table-input" type="number" min="0" placeholder="可填" oninput="updateMonthly();autoSaveMonthly()"></td>
      <td id="mLaborRate-${i}" class="px-4 py-3 text-center font-bold text-indigo-700 fixed-cell"></td>
    </tr>`).join('');
  }

  function loadMonthlyStore() {
    try { monthlyStore = JSON.parse(localStorage.getItem(MONTHLY_KEY) || '{}') || {}; }
    catch(e) { monthlyStore = {}; }
  }

  function currentMonth() { return document.getElementById('monthSelect').value; }

  function fillMonthly() {
    const data = monthlyStore[currentMonth()] || [];
    MONTHLY_ROWS.forEach((d, i) => {
      const r = data[i] || {};
      ['SalesTarget', 'SalesActual', 'LaborTarget', 'LaborActual'].forEach(k => {
        const el = document.getElementById('m' + k + '-' + i);
        if (el) el.value = r[k] || '';
      });
    });
    updateMonthly();
  }

  function collectMonthly() {
    return MONTHLY_ROWS.map((d, i) => ({
      SalesTarget: document.getElementById('mSalesTarget-' + i).value || '',
      SalesActual: document.getElementById('mSalesActual-' + i).value || '',
      LaborTarget: document.getElementById('mLaborTarget-' + i).value || '',
      LaborActual: document.getElementById('mLaborActual-' + i).value || ''
    }));
  }

  function saveMonthlyData(show) {
    monthlyStore[currentMonth()] = collectMonthly();
    localStorage.setItem(MONTHLY_KEY, JSON.stringify(monthlyStore));
    const el = document.getElementById('monthlySaveStatus');
    if (el) {
      el.textContent = (show ? '已保存 ' : '已自动暂存 ') + currentMonth() + ' 的数据，可切换月份或继续修改。';
      el.className = 'text-xs text-green-600 font-semibold';
    }
  }

  let monthlyTimer = null;
  function autoSaveMonthly() {
    clearTimeout(monthlyTimer);
    monthlyTimer = setTimeout(() => saveMonthlyData(false), 250);
  }

  function switchMonth() {
    fillMonthly();
    const el = document.getElementById('monthlySaveStatus');
    if (el) {
      el.textContent = '当前月份：' + currentMonth() + '。';
      el.className = 'text-xs text-slate-500 font-medium';
    }
  }

  function clearMonthlyData() {
    delete monthlyStore[currentMonth()];
    localStorage.setItem(MONTHLY_KEY, JSON.stringify(monthlyStore));
    fillMonthly();
    const el = document.getElementById('monthlySaveStatus');
    if (el) {
      el.textContent = '已清空 ' + currentMonth() + ' 的数据。';
      el.className = 'text-xs text-slate-500 font-medium';
    }
  }

  function updateMonthly() {
    MONTHLY_ROWS.forEach((d, i) => {
      const st = num(document.getElementById('mSalesTarget-' + i).value);
      const sa = num(document.getElementById('mSalesActual-' + i).value);
      const lt = num(document.getElementById('mLaborTarget-' + i).value);
      const la = num(document.getElementById('mLaborActual-' + i).value);
      document.getElementById('mSalesRate-' + i).textContent = (st && sa !== null) ? pct(sa / st) : '';
      document.getElementById('mLaborRate-' + i).textContent = (lt && la !== null) ? pct(la / lt) : '';
    });
  }

  // ANNUAL MODULE
  const rows = [{"dept": "线下（商场）", "salesTarget": 131954430, "laborTarget": 277257, "salesTargetDisplay": "131,954,430", "laborTargetDisplay": "277,257"}, {"dept": "餐饮", "salesTarget": 21270000, "laborTarget": 72635, "salesTargetDisplay": "21,270,000", "laborTargetDisplay": "72,635"}];
  const charts = [echarts.init(document.getElementById('chart0')), echarts.init(document.getElementById('chart1'))];

  function pct(v) { return Number.isFinite(v) ? (v * 100).toFixed(1) + '%' : ''; }
  function num(v) { const n = parseFloat(v); return Number.isFinite(n) ? n : null; }

  function badge(ratio) {
    if (!Number.isFinite(ratio)) return ['待输入', 'bg-slate-100 text-slate-600'];
    const p = ratio * 100;
    if (p < 102) return ['🟢 绿灯 比值 ' + p.toFixed(1) + '%', 'bg-green-100 text-green-700'];
    if (p >= 102 && p <= 108) return ['🟡 黄灯 比值 ' + p.toFixed(1) + '%', 'bg-yellow-100 text-yellow-700'];
    return ['🔴 红灯 比值 ' + p.toFixed(1) + '%', 'bg-red-100 text-red-700'];
  }

  function renderTable() {
    document.getElementById('dataBody').innerHTML = rows.map((r, i) => `<tr class="hover:bg-slate-50">
      <td class="px-4 py-3 font-bold fixed-cell">${r.dept}</td>
      <td class="px-4 py-3 text-right fixed-cell">${r.salesTargetDisplay}</td>
      <td class="px-4 py-3"><input id="sales-${i}" class="table-input" type="number" min="0" placeholder="留空输入" oninput="updateAll();autoSave()"></td>
      <td id="salesRate-${i}" class="px-4 py-3 text-center font-bold text-blue-700 fixed-cell"></td>
      <td class="px-4 py-3 text-right fixed-cell">${r.laborTargetDisplay}</td>
      <td class="px-4 py-3"><input id="labor-${i}" class="table-input" type="number" min="0" placeholder="留空输入" oninput="updateAll();autoSave()"></td>
      <td id="laborRate-${i}" class="px-4 py-3 text-center font-bold text-indigo-700 fixed-cell"></td>
    </tr>`).join('');
  }

  function updateChart(i, salesRate, laborRate) {
    const safeSales = Number.isFinite(salesRate) ? salesRate * 100 : 0;
    const safeLabor = Number.isFinite(laborRate) ? laborRate * 100 : 0;
    charts[i].setOption({
      tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => v.toFixed(1) + '%' },
      grid: { left: 42, right: 18, top: 26, bottom: 30 },
      xAxis: { type: 'category', data: ['销售完成率', '工时使用率'], axisLabel: { fontWeight: 700 } },
      yAxis: { type: 'value', axisLabel: { formatter: '{value}%' }, splitLine: { lineStyle: { color: '#e2e8f0' } } },
      series: [{
        type: 'bar',
        barWidth: '34%',
        data: [
          { value: safeSales, itemStyle: { color: '#2563eb', borderRadius: [8, 8, 0, 0] } },
          { value: safeLabor, itemStyle: { color: '#7c3aed', borderRadius: [8, 8, 0, 0] } }
        ],
        label: { show: true, position: 'top', formatter: p => p.value.toFixed(1) + '%', fontWeight: 700 }
      }]
    });
    const ratio = (Number.isFinite(salesRate) && salesRate !== 0 && Number.isFinite(laborRate)) ? laborRate / salesRate : NaN;
    const b = badge(ratio);
    const el = document.getElementById('badge-' + i);
    el.className = 'shrink-0 rounded-full px-2.5 py-1 text-xs font-bold ' + b[1];
    el.textContent = b[0];
  }

  const STORAGE_KEY = 'suzhou_efficiency_dashboard_inputs_v1';
  function saveData(show) {
    const data = {
      updateTime: document.getElementById('updateTime')?.value || '',
      rows: rows.map((r, i) => ({
        sales: document.getElementById('sales-' + i)?.value || '',
        labor: document.getElementById('labor-' + i)?.value || ''
      }))
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    const el = document.getElementById('saveStatus');
    if (el) {
      el.textContent = show ? '已保存，可继续修改。' : '已自动暂存。';
      el.className = 'text-xs text-green-600 font-semibold';
    }
  }

  let saveTimer = null;
  function autoSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => saveData(false), 250);
  }

  function loadSavedData() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
      if (!data) return;
      if (document.getElementById('updateTime')) document.getElementById('updateTime').value = data.updateTime || '';
      (data.rows || []).forEach((r, i) => {
        const s = document.getElementById('sales-' + i);
        const l = document.getElementById('labor-' + i);
        if (s) s.value = r.sales || '';
        if (l) l.value = r.labor || '';
      });
      const el = document.getElementById('saveStatus');
      if (el) el.textContent = '已载入上次保存的数据，可继续修改。';
    } catch(e) {}
  }

  function clearSavedData() {
    localStorage.removeItem(STORAGE_KEY);
    rows.forEach((r, i) => {
      document.getElementById('sales-' + i).value = '';
      document.getElementById('labor-' + i).value = '';
    });
    document.getElementById('updateTime').value = '';
    updateAll();
    const el = document.getElementById('saveStatus');
    if (el) {
      el.textContent = '已清空保存数据。';
      el.className = 'text-xs text-slate-500 font-medium';
    }
  }

  function updateAll() {
    rows.forEach((r, i) => {
      const actualSales = num(document.getElementById('sales-' + i).value);
      const actualLabor = num(document.getElementById('labor-' + i).value);
      const salesRate = actualSales === null ? NaN : actualSales / r.salesTarget;
      const laborRate = actualLabor === null ? NaN : actualLabor / r.laborTarget;
      document.getElementById('salesRate-' + i).textContent = pct(salesRate);
      document.getElementById('laborRate-' + i).textContent = pct(laborRate);
      updateChart(i, salesRate, laborRate);
    });
  }

  // Window onload initialization
  window.addEventListener('DOMContentLoaded', async () => {
    // Init Module 2 Data
    await initModule2Data();

    // Init Module 1 Tables
    renderMonthOptions();
    renderMonthlyTable();
    loadMonthlyStore();
    fillMonthly();
    renderTable();
    loadSavedData();
    updateAll();

    // Restore last visited tab or default to module2
    const lastPage = localStorage.getItem(PAGE_STORAGE_KEY) || 'module2';
    switchPage(lastPage);
  });

  window.addEventListener('resize', () => {
    charts.forEach(c => c && c.resize());
    if (compareChart) compareChart.resize();
    if (weeklyChart) weeklyChart.resize();
    if (weekdayChart) weekdayChart.resize();
  });
  
  async function clearAllEfficiencyData() {
    if (!confirm('确定要清空所有已上传的工时与驱动指标数据吗？此操作将使看板恢复至初始默认状态且不可撤销。')) return;
    try {
      await EfficiencyStorage.del(STORAGE_KEY_HOURS);
      await EfficiencyStorage.del(STORAGE_KEY_DRIVER);
      await EfficiencyStorage.del(STORAGE_KEY_WEEKDAY_MAP);
      alert('已成功清空本地缓存数据，即将刷新页面恢复默认值。');
      location.reload();
    } catch(e) {
      alert('清空数据失败：' + e.message);
    }
  }
</script>
</body>
</html>
