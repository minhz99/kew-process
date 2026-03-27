// ════════════════════════════════════
    // System Config & State
    // ════════════════════════════════════
    const AppConfig = {
      thdvWarn: 5,
      thdaWarn: 15,
      vunbWarn: 2,
      pfMin: 0.9
    };

    function loadSettings() {
      const saved = localStorage.getItem('plt_config');
      if (saved) {
        try { Object.assign(AppConfig, JSON.parse(saved)); } catch(e){}
      }
      const elV = document.getElementById('set-thdv-warn');
      const elA = document.getElementById('set-thda-warn');
      const elU = document.getElementById('set-vunb-warn');
      const elP = document.getElementById('set-pf-min');
      if(elV) elV.value = AppConfig.thdvWarn;
      if(elA) elA.value = AppConfig.thdaWarn;
      if(elU) elU.value = AppConfig.vunbWarn;
      if(elP) elP.value = AppConfig.pfMin;
    }

    function saveSettings() {
      AppConfig.thdvWarn = parseFloat(document.getElementById('set-thdv-warn').value) || 5;
      AppConfig.thdaWarn = parseFloat(document.getElementById('set-thda-warn').value) || 15;
      AppConfig.vunbWarn = parseFloat(document.getElementById('set-vunb-warn').value) || 2;
      AppConfig.pfMin = parseFloat(document.getElementById('set-pf-min').value) || 0.9;
      
      localStorage.setItem('plt_config', JSON.stringify(AppConfig));
      
      if (typeof DATA !== 'undefined' && DATA) {
        renderAssessmentTab(DATA.summary, DATA.series);
        renderTHDGrid(DATA.series);
      }
    }

    document.addEventListener('DOMContentLoaded', loadSettings);

    // ════════════════════════════════════
    // Load data
    // ════════════════════════════════════
    let DATA = null;
    let CHARTS = {};

    const COLORS = {
      V1: '#3b82f6', V2: '#8b5cf6', V3: '#ec4899',
      A1: '#10b981', A2: '#14b8a6', A3: '#06b6d4',
      P1: '#f59e0b', P2: '#f97316', P3: '#ef4444', P: '#a78bfa',
    };

    const CHART_DEFAULTS = {
      responsive: true,
      maintainAspectRatio: true,
      animation: false, // Disable all animations for maximum speed
      normalized: true, // Tell Chart.js data is normalized
      spanGaps: true,   // Performance boost for large datasets
      interaction: {
        mode: 'index',
        intersect: false,
      },
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 11 }, boxWidth: 12 } },
        tooltip: {
          enabled: true,
          position: 'nearest',
          external: null,
          backgroundColor: '#1a2235',
          borderColor: '#1f2d45',
          borderWidth: 1,
          titleColor: '#e2e8f0',
          bodyColor: '#94a3b8',
          padding: 8,
          cornerRadius: 4,
          displayColors: true,
          // Optimization: only update tooltip when necessary
          animation: false
        }
      },
      scales: {
        x: {
          ticks: { color: '#475569', maxTicksLimit: 10, font: { size: 10 } },
          grid: { color: '#1e2a3a' },
        },
        y: {
          beginAtZero: false,
          ticks: { color: '#475569', font: { size: 10 } },
          grid: { color: '#1e2a3a' },
        }
      },
      elements: {
        line: {
          tension: 0, // Straight lines are much faster than splines
          borderWidth: 1.5,
          capBezierPoints: false // Disable bezier for speed
        },
        point: {
          radius: 0, // No points for speed
          hitRadius: 5, 
          hoverRadius: 4
        }
      }
    };

    function cloneDefaults() { return JSON.parse(JSON.stringify(CHART_DEFAULTS)); }

    // ─── Tab switching ─────────────────
    const RENDERED_TABS = new Set();

    function renderTabWithData(name) {
      if (!DATA) return;
      // Skip if already rendered, UNLESS it's harmonics which might need re-render on channel change
      // (Actually harmonics handles itself, but we should allow re-render if needed)
      if (RENDERED_TABS.has(name) && name !== 'harmonics') return;
      
      const s = DATA.series;
      const sum = DATA.summary;
      
      try {
        if (name === 'overview') {
          renderKPIs(sum, s);
          renderOverviewCharts(s, sum);
        } else if (name === 'voltage') {
          renderVoltageTab(s);
        } else if (name === 'current') {
          renderCurrentTab(s);
        } else if (name === 'power') {
          renderPowerTab(s);
        } else if (name === 'harmonics') {
          renderHarmonicsSelector(s);
          renderHarmonics();
          renderTHDGrid(s);
        } else if (name === 'events') {
          renderEventsTab(sum);
        } else if (name === 'assessment') {
          renderAssessmentTab(sum, s);
        } else if (name === 'correction') {
          // Correction tab charts depend on current session state, but we can init preview
          if (typeof updateCorrPreview === 'function') updateCorrPreview();
        }
        
        RENDERED_TABS.add(name);
        console.log(`[Perf] Lazy rendered tab: ${name}`);
      } catch (e) {
        console.error(`[Error] Render tab ${name} failed:`, e);
      }
    }

    function switchKewTab(name) {
      const contentEl = document.querySelector(`#workspace-kew #tab-${name}`);
      if (!contentEl) return;

      document.querySelectorAll('#workspace-kew .content').forEach(el => el.classList.remove('active'));
      document.querySelectorAll('#kew-nav-section .tab').forEach(el => el.classList.remove('active'));
      contentEl.classList.add('active');
      const activeTabBtn = document.querySelector(`#kew-nav-section [data-kew-tab="${name}"]`);
      if (activeTabBtn) activeTabBtn.classList.add('active');

      // Lazy render data if available
      if (DATA) {
        renderTabWithData(name);
      }
    }

    // ─── Subsample timestamps for display ─────────────────
    function subsample(arr, maxPoints = 300) {
      if (arr.length <= maxPoints) return arr;
      const step = Math.ceil(arr.length / maxPoints);
      return arr.filter((_, i) => i % step === 0);
    }

    function subsamplePair(ts, vals, maxPoints = 300) {
      if (ts.length <= maxPoints) return [ts, vals];
      const step = Math.ceil(ts.length / maxPoints);
      return [
        ts.filter((_, i) => i % step === 0),
        vals.filter((_, i) => i % step === 0),
      ];
    }

    function fmtTS(tsList) {
      return tsList.map(s => s.substring(11, 19));  // just HH:MM:SS
    }

    function makeLineDataset(label, data, color, fill = false) {
      return {
        label,
        data,
        borderColor: color,
        backgroundColor: fill ? color + '22' : 'transparent',
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0, // Straight lines are faster to render than curves
        fill,
        spanGaps: true
      };
    }

    // ═══════════ PHASE A: RENDER ALL ═══════════
    function renderAll(data) {
      DATA = data;
      RENDERED_TABS.clear(); // Clear cache on new data load

      renderHeaderMeta(data.summary);
      
      // Re-initialize active tab
      const activeTabEl = document.querySelector('#workspace-kew .content.active');
      const activeName = activeTabEl ? activeTabEl.id.replace('tab-', '') : 'overview';
      
      if (activeName.includes('workspace')) {
          // if we are in another workspace, switch back to overview in kew workspace
          switchKewTab('overview');
      } else {
          renderTabWithData(activeName);
      }
    }

    // ─── Header meta ─────────────────
    function renderHeaderMeta(sum) {
      const dev = sum.device;
      const deviceInfoEl = document.getElementById('deviceInfo');
      const headerMetaEl = document.getElementById('headerMeta');

      if (deviceInfoEl) {
        deviceInfoEl.textContent = `Thiết bị: S/N ${dev['SERIAL NUMBER'] || '—'} | ${dev['WIRING'] || ''} | ${dev['NOMINAL VOLTAGE'] || ''} | ${dev['FREQUENCY'] || ''}`;
      }

      if (headerMetaEl) {
        headerMetaEl.innerHTML = `
    <div class="meta-chip">Bắt đầu: <span>${sum.time_start?.substring(0, 16) || '—'}</span></div>
    <div class="meta-chip">Kết thúc: <span>${sum.time_end?.substring(0, 16) || '—'}</span></div>
    <div class="meta-chip">Mẫu: <span>${sum.num_samples}</span></div>
  `;
      }
    }

    // ─── KPI cards ─────────────────
    function renderKPIs(sum, s) {
      const grid = document.getElementById('kpiGrid');
      const kpis = [];
      const _n = (v, d = 1) => (v !== null && v !== undefined) ? v.toFixed(d) : '—';
      const _nk = (v) => (v !== null && v !== undefined) ? (v / 1000).toFixed(1) : '—';

      // Voltage
      for (const ch of ['V1[V]', 'V2[V]', 'V3[V]']) {
        if (!sum.voltage[ch]) continue;
        const v = sum.voltage[ch];
        kpis.push({ label: `Điện áp ${ch.replace('[V]', '')} (RMS fund.)`, value: _n(v.avg), unit: 'V', sub: `Min: ${_n(v.min)} | Max: ${_n(v.max)}`, color: COLORS[ch.substring(0, 2)] });
      }
      // Current
      for (const ch of ['A1[A]', 'A2[A]', 'A3[A]']) {
        if (!sum.current[ch]) continue;
        const v = sum.current[ch];
        kpis.push({ label: `Dòng điện ${ch.replace('[A]', '')} (fund.)`, value: _n(v.avg), unit: 'A', sub: `Min: ${_n(v.min)} | Max: ${_n(v.max)}`, color: COLORS[ch.substring(0, 2)] });
      }
      // Total power
      if (sum.power['P[W]']) {
        const p = sum.power['P[W]'];
        kpis.push({ label: 'Tổng công suất P', value: _nk(p.avg), unit: 'kW', sub: `Max: ${_nk(p.max)} kW`, color: '#a78bfa' });
      }
      // Events
      kpis.push({ label: 'Sự kiện PQ phát hiện', value: sum.event_count || 0, unit: 'events', sub: 'Dip, Swell, Transient…', color: (sum.event_count || 0) > 0 ? '#ef4444' : '#10b981' });

      grid.innerHTML = kpis.map(k => `
    <div class="kpi" style="--bar-color:${k.color}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value !== undefined && k.value !== null ? k.value : '—'}</div>
      <div class="kpi-unit">${k.unit}</div>
      <div class="kpi-sub">${k.sub || ''}</div>
    </div>
  `).join('');
    }

    // ─── Overview charts ─────────────────
    function renderOverviewCharts(s, sum) {
      // Voltage
      const [vts, v1s] = subsamplePair(s['V1[V]'].timestamps, s['V1[V]'].fundamental);
      const [, v2s] = subsamplePair(s['V2[V]'].timestamps, s['V2[V]'].fundamental);
      const [, v3s] = subsamplePair(s['V3[V]'].timestamps, s['V3[V]'].fundamental);
      createOrUpdate('overviewVoltChart', {
        type: 'line',
        data: {
          labels: fmtTS(vts),
          datasets: [
            makeLineDataset('V1', v1s, COLORS.V1),
            makeLineDataset('V2', v2s, COLORS.V2),
            makeLineDataset('V3', v3s, COLORS.V3),
          ]
        },
        options: { ...cloneDefaults() }
      });

      // Current
      const [ats, a1s] = subsamplePair(s['A1[A]'].timestamps, s['A1[A]'].fundamental);
      const [, a2s] = subsamplePair(s['A2[A]'].timestamps, s['A2[A]'].fundamental);
      const [, a3s] = subsamplePair(s['A3[A]'].timestamps, s['A3[A]'].fundamental);
      createOrUpdate('overviewCurrChart', {
        type: 'line',
        data: {
          labels: fmtTS(ats),
          datasets: [
            makeLineDataset('A1', a1s, COLORS.A1),
            makeLineDataset('A2', a2s, COLORS.A2),
            makeLineDataset('A3', a3s, COLORS.A3),
          ]
        },
        options: { ...cloneDefaults() }
      });

      // Power
      if (s['P1[W]'] && s['P2[W]'] && s['P3[W]']) {
        const [pts, p1s] = subsamplePair(s['P1[W]'].timestamps, s['P1[W]'].fundamental);
        const [, p2s] = subsamplePair(s['P2[W]'].timestamps, s['P2[W]'].fundamental);
        const [, p3s] = subsamplePair(s['P3[W]'].timestamps, s['P3[W]'].fundamental);
        createOrUpdate('overviewPowerChart', {
          type: 'line',
          data: {
            labels: fmtTS(pts),
            datasets: [
              makeLineDataset('P1', p1s, COLORS.P1),
              makeLineDataset('P2', p2s, COLORS.P2),
              makeLineDataset('P3', p3s, COLORS.P3),
            ]
          },
          options: { ...cloneDefaults() }
        });
      }

      // THD trend (voltage)
      const [tts, thd1] = subsamplePair(s['V1[V]'].timestamps, s['V1[V]'].thd);
      const [, thd2] = subsamplePair(s['V2[V]'].timestamps, s['V2[V]'].thd);
      const [, thd3] = subsamplePair(s['V3[V]'].timestamps, s['V3[V]'].thd);
      createOrUpdate('overviewTHDChart', {
        type: 'line',
        data: {
          labels: fmtTS(tts),
          datasets: [
            makeLineDataset('THD V1', thd1, COLORS.V1),
            makeLineDataset('THD V2', thd2, COLORS.V2),
            makeLineDataset('THD V3', thd3, COLORS.V3),
          ]
        },
        options: { ...cloneDefaults() }
      });
    }

    // ─── Voltage tab ─────────────────
    let voltPhase = 'all';
    function setVoltPhase(ph, btnEl) {
      if (!DATA) return;
      voltPhase = ph;
      document.querySelectorAll('#voltPhaseButtons .btn').forEach(b => b.classList.remove('active'));
      if (btnEl) btnEl.classList.add('active');
      renderVoltageTab(DATA.series);
    }

    function renderVoltageTab(s) {
      const phases = voltPhase === 'all' ? ['V1[V]', 'V2[V]', 'V3[V]'] : [voltPhase];
      const datasets = phases.map(ch => {
        const [ts, vals] = subsamplePair(s[ch].timestamps, s[ch].fundamental);
        return { ts, vals, ch };
      });
      const ts = datasets[0].ts;
      createOrUpdate('voltTimeChart', {
        type: 'line',
        data: {
          labels: fmtTS(ts),
          datasets: datasets.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)], datasets.length === 1))
        },
        options: { ...cloneDefaults() }
      });

      // KPIs
      const grid = document.getElementById('voltKpiGrid');
      const vch = ['V1[V]', 'V2[V]', 'V3[V]'];
      grid.innerHTML = vch.map(ch => {
        const d = s[ch];
        return `<div class="kpi" style="--bar-color:${COLORS[ch.substring(0, 2)]}">
      <div class="kpi-label">${ch} — Trung bình</div>
      <div class="kpi-value">${d.avg_fundamental?.toFixed(1)}</div>
      <div class="kpi-unit">V</div>
      <div class="kpi-sub">Min: ${d.min_fundamental?.toFixed(1)} | Max: ${d.max_fundamental?.toFixed(1)} | THD: ${d.avg_thd?.toFixed(2)}%</div>
    </div>`;
      }).join('');

      // THD line
      const [tts, thd1] = subsamplePair(s['V1[V]'].timestamps, s['V1[V]'].thd);
      const [, thd2] = subsamplePair(s['V2[V]'].timestamps, s['V2[V]'].thd);
      const [, thd3] = subsamplePair(s['V3[V]'].timestamps, s['V3[V]'].thd);
      createOrUpdate('voltTHDChart', {
        type: 'line',
        data: {
          labels: fmtTS(tts),
          datasets: [
            makeLineDataset('THD V1', thd1, COLORS.V1),
            makeLineDataset('THD V2', thd2, COLORS.V2),
            makeLineDataset('THD V3', thd3, COLORS.V3),
          ]
        },
        options: { ...cloneDefaults() }
      });

      // THD bar
      createOrUpdate('voltTHDBarChart', {
        type: 'bar',
        data: {
          labels: ['V1', 'V2', 'V3'],
          datasets: [{
            label: 'THD% trung bình',
            data: ['V1[V]', 'V2[V]', 'V3[V]'].map(ch => s[ch]?.avg_thd?.toFixed(3)),
            backgroundColor: [COLORS.V1 + '99', COLORS.V2 + '99', COLORS.V3 + '99'],
            borderColor: [COLORS.V1, COLORS.V2, COLORS.V3],
            borderWidth: 1,
            borderRadius: 6,
          }]
        },
        options: { ...cloneDefaults(), plugins: { ...cloneDefaults().plugins, legend: { display: false } } }
      });
    }

    // ─── Current tab ─────────────────
    let currPhase = 'all';
    function setCurrPhase(ph, btnEl) {
      if (!DATA) return;
      currPhase = ph;
      document.querySelectorAll('#currPhaseButtons .btn').forEach(b => b.classList.remove('active'));
      if (btnEl) btnEl.classList.add('active');
      renderCurrentTab(DATA.series);
    }

    function renderCurrentTab(s) {
      const phases = currPhase === 'all' ? ['A1[A]', 'A2[A]', 'A3[A]'] : [currPhase];
      const datasets = phases.map(ch => {
        const [ts, vals] = subsamplePair(s[ch].timestamps, s[ch].fundamental);
        return { ts, vals, ch };
      });
      const ts = datasets[0].ts;
      createOrUpdate('currTimeChart', {
        type: 'line',
        data: {
          labels: fmtTS(ts),
          datasets: datasets.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)], datasets.length === 1))
        },
        options: { ...cloneDefaults() }
      });

      const grid = document.getElementById('currKpiGrid');
      const ach = ['A1[A]', 'A2[A]', 'A3[A]'];
      grid.innerHTML = ach.map(ch => {
        const d = s[ch];
        return `<div class="kpi" style="--bar-color:${COLORS[ch.substring(0, 2)]}">
      <div class="kpi-label">${ch} — Trung bình</div>
      <div class="kpi-value">${d.avg_fundamental?.toFixed(1)}</div>
      <div class="kpi-unit">A</div>
      <div class="kpi-sub">Min: ${d.min_fundamental?.toFixed(1)} | Max: ${d.max_fundamental?.toFixed(1)} | THD: ${d.avg_thd?.toFixed(2)}%</div>
    </div>`;
      }).join('');

      const [tts, thd1] = subsamplePair(s['A1[A]'].timestamps, s['A1[A]'].thd);
      const [, thd2] = subsamplePair(s['A2[A]'].timestamps, s['A2[A]'].thd);
      const [, thd3] = subsamplePair(s['A3[A]'].timestamps, s['A3[A]'].thd);
      createOrUpdate('currTHDChart', {
        type: 'line',
        data: {
          labels: fmtTS(tts),
          datasets: [
            makeLineDataset('THD A1', thd1, COLORS.A1),
            makeLineDataset('THD A2', thd2, COLORS.A2),
            makeLineDataset('THD A3', thd3, COLORS.A3),
          ]
        },
        options: { ...cloneDefaults() }
      });
      createOrUpdate('currTHDBarChart', {
        type: 'bar',
        data: {
          labels: ['A1', 'A2', 'A3'],
          datasets: [{
            label: 'THD% trung bình',
            data: ['A1[A]', 'A2[A]', 'A3[A]'].map(ch => s[ch]?.avg_thd?.toFixed(3)),
            backgroundColor: [COLORS.A1 + '99', COLORS.A2 + '99', COLORS.A3 + '99'],
            borderColor: [COLORS.A1, COLORS.A2, COLORS.A3],
            borderWidth: 1,
            borderRadius: 6,
          }]
        },
        options: { ...cloneDefaults(), plugins: { ...cloneDefaults().plugins, legend: { display: false } } }
      });
    }

    // ─── Power tab ─────────────────
    function renderPowerTab(s) {
      const pChs = ['P1[W]', 'P2[W]', 'P3[W]', 'P[W]'].filter(ch => s[ch]);
      const datasets = pChs.map(ch => {
        const [ts, vals] = subsamplePair(s[ch].timestamps, s[ch].fundamental);
        return { ts, vals, ch };
      });
      const ts = datasets[0]?.ts || [];
      createOrUpdate('powerTimeChart', {
        type: 'line',
        data: {
          labels: fmtTS(ts),
          datasets: datasets.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)]))
        },
        options: { ...cloneDefaults() }
      });

      const grid = document.getElementById('powerKpiGrid');
      grid.innerHTML = pChs.map(ch => {
        const d = s[ch];
        const isKW = d.avg_fundamental > 1000;
        const fmt = v => isKW ? (v / 1000).toFixed(1) : v?.toFixed(1);
        const unit = isKW ? 'kW' : 'W';
        return `<div class="kpi" style="--bar-color:${COLORS[ch.substring(0, 2)]}">
      <div class="kpi-label">${ch}</div>
      <div class="kpi-value">${fmt(d.avg_fundamental)}</div>
      <div class="kpi-unit">${unit}</div>
      <div class="kpi-sub">Min: ${fmt(d.min_fundamental)} | Max: ${fmt(d.max_fundamental)}</div>
    </div>`;
      }).join('');

      // Power factor (from VA deg channels)
      const pf_chs = ['VA1[deg]', 'VA2[deg]', 'VA3[deg]'].filter(ch => s[ch]);
      if (pf_chs.length > 0) {
        const pfTs = subsample(s[pf_chs[0]]?.timestamps || [], 300);
        createOrUpdate('pfChart', {
          type: 'line',
          data: {
            labels: fmtTS(pfTs),
            datasets: pf_chs.map((ch, i) => {
              const [, vals] = subsamplePair(s[ch].timestamps, s[ch].fundamental);
              return makeLineDataset(ch, vals, [COLORS.V1, COLORS.V2, COLORS.V3][i]);
            })
          },
          options: { ...cloneDefaults() }
        });
        // cos(phi) estimate
        const pfAvgs = pf_chs.map(ch => {
          const deg = s[ch].avg_fundamental;
          return Math.cos(deg * Math.PI / 180).toFixed(3);
        });
        createOrUpdate('pfBarChart', {
          type: 'bar',
          data: {
            labels: ['Pha 1', 'Pha 2', 'Pha 3'].slice(0, pf_chs.length),
            datasets: [{
              label: 'cos(φ)',
              data: pfAvgs,
              backgroundColor: [COLORS.V1 + 'aa', COLORS.V2 + 'aa', COLORS.V3 + 'aa'],
              borderColor: [COLORS.V1, COLORS.V2, COLORS.V3],
              borderWidth: 1, borderRadius: 6,
            }]
          },
          options: {
            ...cloneDefaults(),
            scales: {
              ...cloneDefaults().scales,
              y: { ...cloneDefaults().scales.y, min: -1, max: 1 }
            }
          }
        });
      }
    }

    // ─── Harmonics tab ─────────────────
    function renderHarmonicsSelector(s) {
      const sel = document.getElementById('harmonicsChannel');
      sel.innerHTML = Object.keys(s).map(ch => `<option value="${ch}">${ch}</option>`).join('');
    }

    function renderHarmonics() {
      if (!DATA) return;
      const s = DATA.series;
      const ch = document.getElementById('harmonicsChannel').value;
      document.getElementById('harmonicsChLabel').textContent = ch;
      const d = s[ch];
      if (!d) return;

      const labels = Array.from({ length: 50 }, (_, i) => `H${i + 1}`);
      const spectrum = d.spectrum || [];

      createOrUpdate('spectrumChart', {
        type: 'bar',
        data: {
          labels,
          datasets: [{
            label: `${ch} — Trung bình`,
            data: spectrum,
            backgroundColor: spectrum.map((v, i) => i === 0 ? '#3b82f6cc' : '#8b5cf666'),
            borderColor: spectrum.map((v, i) => i === 0 ? '#3b82f6' : '#8b5cf6'),
            borderWidth: 1,
            borderRadius: 3,
          }]
        },
        options: { ...cloneDefaults(), plugins: { ...cloneDefaults().plugins, legend: { display: false } } }
      });

      const [tts, thd] = subsamplePair(d.timestamps, d.thd);
      createOrUpdate('harmonicsTHDLine', {
        type: 'line',
        data: {
          labels: fmtTS(tts),
          datasets: [makeLineDataset(`THD ${ch}`, thd, '#3b82f6', true)]
        },
        options: { ...cloneDefaults() }
      });
    }

    function renderTHDGrid(s) {
      const grid = document.getElementById('thdGrid');
      const allChs = Object.keys(s).sort();
      grid.innerHTML = allChs.map(ch => {
        const thd = s[ch].avg_thd || 0;
        const isVolt = ch.includes('V');
        const warnThreshold = isVolt ? AppConfig.thdvWarn : AppConfig.thdaWarn;
        const isWarn = thd > warnThreshold;
        const pct = Math.min(thd * (isVolt ? 10 : 2.5), 100);  // scale visually
        return `<div class="thd-bar">
      <div class="thd-bar-label">${ch}</div>
      <div class="thd-bar-track"><div class="thd-bar-fill ${isWarn ? 'warn' : ''}" style="width:${pct}%"></div></div>
      <div class="thd-bar-value">THD: ${thd.toFixed(2)}% ${isWarn ? '⚠️' : '✅'}</div>
    </div>`;
      }).join('');
    }

    // ─── Events tab ─────────────────
    function renderEventsTab(sum) {
      const events = sum.events || [];
      const typeMap = {
        'Voltage Dip': 'badge-dip',
        'Voltage Swell': 'badge-swell',
        'Transient': 'badge-transient',
        'Interrupt': 'badge-interrupt',
        'Inrush Current': 'badge-inrush',
      };

      const countByType = {};
      events.forEach(e => { countByType[e.type] = (countByType[e.type] || 0) + 1; });

      const kpiGrid = document.getElementById('eventKpiGrid');
      const types = Object.keys(typeMap);
      kpiGrid.innerHTML = types.map(t => `
    <div class="kpi" style="--bar-color:${t === 'Voltage Dip' ? '#ef4444' : t === 'Voltage Swell' ? '#f59e0b' : '#8b5cf6'}">
      <div class="kpi-label">${t}</div>
      <div class="kpi-value">${countByType[t] || 0}</div>
      <div class="kpi-unit">sự kiện</div>
    </div>
  `).join('');

      const body = document.getElementById('eventsBody');
      if (events.length === 0) {
        body.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--text-muted);padding:2rem">✅ Không có sự kiện PQ nào được phát hiện trong khoảng thời gian đo.</td></tr>`;
        return;
      }
      body.innerHTML = events.map((e, i) => {
        const cls = typeMap[e.type] || '';
        const val = e.value != null ? e.value.toFixed(1) : '—';
        const unit = e.type.includes('Dip') || e.type.includes('Swell') || e.type.includes('Transient') ? 'V' : 'A';
        let comment = '';
        if (e.type === 'Voltage Dip') comment = `Sụt áp xuống ${val}V`;
        else if (e.type === 'Voltage Swell') comment = `Quá áp lên ${val}V`;
        else if (e.type === 'Transient') comment = 'Xung điện nhất thời';
        else if (e.type === 'Interrupt') comment = 'Mất điện ngắn hạn';
        else if (e.type === 'Inrush Current') comment = `Dòng khởi động ${val}A`;
        return `<tr>
      <td style="color:var(--text-muted)">${i + 1}</td>
      <td>${e.datetime?.substring(0, 19) || '—'}</td>
      <td style="color:var(--text-muted)">${e.elapsed || '—'}</td>
      <td><span class="badge-event ${cls}">${e.type}</span></td>
      <td style="font-weight:600">${val} ${val !== '—' ? unit : ''}</td>
      <td style="color:var(--text-muted)">${comment}</td>
    </tr>`;
      }).join('');
    }

    // ─── Assessment tab ─────────────────
    function updateCommentaryMBA() {
      if (!DATA || !DATA.commentary || !DATA.commentary.text) return;
      // Get max apparent power S in kVA
      let sMax = 0;
      if (DATA.summary?.inps?.apparent_power) {
        const spKey = Object.keys(DATA.summary.inps.apparent_power).find(k => k.includes('[kVA]') || k.includes('[VA]'));
        if (spKey) sMax = DATA.summary.inps.apparent_power[spKey].max || 0;
        if (spKey && spKey.includes('[VA]') && !spKey.includes('k')) sMax /= 1000;
      }
      if (!sMax) {
        // Fallback if INPS not available: Use max Power P[W] from INHS
        sMax = (DATA.summary?.power?.['P[W]']?.max || 0) / 1000;
      }

      const target = parseFloat(document.getElementById('mbaPower').value);
      const textEl = document.getElementById('commentaryText');
      let text = DATA.commentary.text;

      if (target > 0 && sMax > 0) {
        const pct = (sMax / target * 100).toFixed(2).replace('.', ',');
        text = `Nhận xét: Công suất tiêu thụ của máy biến áp đạt ${pct}% công suất thiết kế. ` + text.replace("Nhận xét: ", "");
      }
      textEl.innerText = text;
    }

    function copySummaryTable() {
      const table = document.getElementById('summaryTable');
      let text = '';
      for (let i = 0; i < table.rows.length; i++) {
        const row = table.rows[i];
        const rowData = [];
        for (let j = 0; j < row.cells.length; j++) {
          let cellText = row.cells[j].innerText.trim();
          if (i > 0) cellText = cellText.replace(/\n/g, ' ');
          rowData.push(cellText);
        }
        text += rowData.join('\t') + '\n';
      }
      const tempTextarea = document.createElement('textarea');
      tempTextarea.value = text;
      document.body.appendChild(tempTextarea);
      tempTextarea.select();
      document.execCommand('copy');
      document.body.removeChild(tempTextarea);
      alert('Đã copy bảng tổng hợp vào clipboard!');
    }

    function renderAssessmentTab(sum, s) {
      const commentary = DATA.commentary || {};
      if (commentary.text) {
        document.getElementById('commentaryText').innerText = commentary.text;
        updateCommentaryMBA();
      }
      if (commentary.table) {
        const tbody = document.getElementById('summaryTableBody');
        tbody.innerHTML = commentary.table.map(row =>
          `<tr>${row.map((cell, i) => `<td ${i < 3 ? 'style="color:var(--text-muted);font-weight:500;"' : ''}>${cell || ''}</td>`).join('')}</tr>`
        ).join('');
      }

      const grid = document.getElementById('assessmentGrid');
      const dev = sum.device;
      const nomVolt = 600; // from device
      const items = [];

      // Voltage balance
      const v1 = sum.voltage['V1[V]']?.avg;
      const v2 = sum.voltage['V2[V]']?.avg;
      const v3 = sum.voltage['V3[V]']?.avg;
      if (v1 && v2 && v3) {
        const vAvg = (v1 + v2 + v3) / 3;
        const vImbalance = (Math.max(Math.abs(v1 - vAvg), Math.abs(v2 - vAvg), Math.abs(v3 - vAvg)) / vAvg * 100);
        items.push({
          icon: vImbalance < AppConfig.vunbWarn ? '✅' : vImbalance < AppConfig.vunbWarn + 3 ? '⚠️' : '❌',
          cls: vImbalance < AppConfig.vunbWarn ? 'assess-pass' : vImbalance < AppConfig.vunbWarn + 3 ? 'assess-warn' : 'assess-fail',
          title: 'Cân bằng điện áp 3 pha',
          desc: `Mức mất cân bằng: ${vImbalance.toFixed(2)}% (ngưỡng cảnh báo: >${AppConfig.vunbWarn}%). Điện áp pha trung bình: ${vAvg.toFixed(1)} V.`,
        });
      }

      // THD voltage
      const vThds = ['V1[V]', 'V2[V]', 'V3[V]'].map(ch => s[ch]?.avg_thd).filter(v => v != null);
      if (vThds.length > 0) {
        const maxVthd = Math.max(...vThds);
        items.push({
          icon: maxVthd < AppConfig.thdvWarn ? '✅' : maxVthd < AppConfig.thdvWarn + 3 ? '⚠️' : '❌',
          cls: maxVthd < AppConfig.thdvWarn ? 'assess-pass' : maxVthd < AppConfig.thdvWarn + 3 ? 'assess-warn' : 'assess-fail',
          title: 'THD Điện áp (IEEE 519)',
          desc: `THD max: ${maxVthd.toFixed(2)}%. Giới hạn cảnh báo: >${AppConfig.thdvWarn}%. ${maxVthd < AppConfig.thdvWarn ? 'Đạt tiêu chuẩn.' : 'Vượt ngưỡng khuyến nghị!'}`,
        });
      }

      // THD current
      const aThds = ['A1[A]', 'A2[A]', 'A3[A]'].map(ch => s[ch]?.avg_thd).filter(v => v != null);
      if (aThds.length > 0) {
        const maxAthd = Math.max(...aThds);
        items.push({
          icon: maxAthd < AppConfig.thdaWarn ? '✅' : maxAthd < AppConfig.thdaWarn + 5 ? '⚠️' : '❌',
          cls: maxAthd < AppConfig.thdaWarn ? 'assess-pass' : maxAthd < AppConfig.thdaWarn + 5 ? 'assess-warn' : 'assess-fail',
          title: 'THD Dòng điện',
          desc: `THD dòng max: ${maxAthd.toFixed(2)}%. Giới hạn cảnh báo: >${AppConfig.thdaWarn}%. ${maxAthd < AppConfig.thdaWarn ? 'Đạt tiêu chuẩn.' : 'Cần kiểm tra tải phi tuyến!'}`,
        });
      }

      // Events
      items.push({
        icon: (sum.event_count || 0) === 0 ? '✅' : '⚠️',
        cls: (sum.event_count || 0) === 0 ? 'assess-pass' : 'assess-warn',
        title: 'Sự kiện chất lượng điện',
        desc: `Phát hiện ${sum.event_count || 0} sự kiện (Dip, Swell, Transient, Inrush). ${(sum.event_count || 0) === 0 ? 'Chất lượng điện ổn định.' : 'Cần theo dõi thêm!'}`,
      });

      // Power factor
      const pf_chs = ['VA1[deg]', 'VA2[deg]', 'VA3[deg]'].filter(ch => s[ch] && s[ch].avg_fundamental != null);
      if (pf_chs.length > 0) {
        const pfs = pf_chs.map(ch => Math.abs(Math.cos(s[ch].avg_fundamental * Math.PI / 180)));
        const minPF = Math.min(...pfs);
        items.push({
          icon: minPF > AppConfig.pfMin ? '✅' : minPF > AppConfig.pfMin - 0.1 ? '⚠️' : '❌',
          cls: minPF > AppConfig.pfMin ? 'assess-pass' : minPF > AppConfig.pfMin - 0.1 ? 'assess-warn' : 'assess-fail',
          title: 'Hệ số công suất cos(φ)',
          desc: `cos(φ) thấp nhất: ${minPF.toFixed(3)}. Yêu cầu tối thiểu: >${AppConfig.pfMin}. ${minPF > AppConfig.pfMin ? 'Đạt yêu cầu.' : 'Cân nhắc bù cos(φ)!'}`,
        });
      }

      // Device info
      items.push({
        icon: 'ℹ️',
        cls: 'assess-info',
        title: 'Thông tin đo lường',
        desc: `Thiết bị: KEW ${dev['FILE ID'] || ''} | S/N: ${dev['SERIAL NUMBER'] || ''} | Thời gian: ${sum.time_start?.substring(0, 16)} — ${sum.time_end?.substring(0, 16)} | Tổng ${sum.num_samples} mẫu.`,
      });

      grid.innerHTML = items.map(it => `
    <div class="assessment-item">
      <div class="assess-icon ${it.cls}">${it.icon}</div>
      <div class="assess-body">
        <h4>${it.title}</h4>
        <p>${it.desc}</p>
      </div>
    </div>
  `).join('');

      // Device metadata table
      const body = document.getElementById('deviceBody');
      body.innerHTML = Object.entries(dev).map(([k, v]) => `<tr><td><strong>${k}</strong></td><td>${v}</td></tr>`).join('');
    }

    // ─── Chart helper ─────────────────
    function createOrUpdate(canvasId, config) {
      const canvas = document.getElementById(canvasId);
      if (!canvas) return;
      if (CHARTS[canvasId]) {
        CHARTS[canvasId].destroy();
      }
      CHARTS[canvasId] = new Chart(canvas, config);
    }

    // ════════════════════════════════════
    // Upload Logic
    // ════════════════════════════════════
    const dropArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const loadingScreen = document.getElementById('loading-screen');
    const loadingText = document.getElementById('loading-text');
    const errorMsg = document.getElementById('upload-error');
    const successMsg = document.getElementById('upload-success');

    // Store last uploaded FormData for the gen action
    let lastFormData = null;
    let uploadStateToken = 0;

    function setUploadState(state, msg = '') {
      if (!dropArea || !loadingScreen || !errorMsg) return;

      dropArea.style.display = 'block';
      dropArea.classList.remove('is-busy');
      loadingScreen.style.display = 'none';
      errorMsg.style.display = 'none';
      if (successMsg) successMsg.style.display = 'none';

      if (state === 'loading') {
        dropArea.classList.add('is-busy');
        loadingScreen.style.display = 'block';
        if (loadingText) loadingText.textContent = msg || 'Đang phân tích dữ liệu...';
        return;
      }

      if (state === 'error') {
        errorMsg.textContent = msg;
        errorMsg.style.display = 'block';
        return;
      }

      if (state === 'success' && successMsg) {
        successMsg.textContent = msg;
        successMsg.style.display = 'block';
      }
    }

    function preventDefaults(e) {
      e.preventDefault();
      e.stopPropagation();
    }

    if (dropArea && fileInput && loadingScreen && errorMsg) {
      setUploadState('idle');

      ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
      });

      ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.add('dragover'), false);
      });
      ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.remove('dragover'), false);
      });

      dropArea.addEventListener('drop', handleDrop, false);
      fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        e.target.value = '';
      }, false);
    }

    function handleDrop(e) {
      handleFiles(e.dataTransfer.files);
    }

    function handleFiles(files) {
      if (!files || files.length === 0) return;

      const formData = new FormData();
      let hasKew = false;
      for (let i = 0; i < files.length; i++) {
        if (files[i].name.toUpperCase().endsWith('.KEW') || files[i].name.toUpperCase().endsWith('.ZIP')) {
          formData.append('files', files[i]);
          hasKew = true;
        }
      }

      if (!hasKew) {
        showError("Vui lòng chọn ít nhất 1 file định dạng .KEW hoặc .ZIP");
        return;
      }

      lastFormData = formData;
      uploadFiles(formData);
    }

    function showError(msg) {
      setUploadState('error', msg);
    }

    function uploadFiles(formData) {
      const stateToken = ++uploadStateToken;
      setUploadState('loading', 'Đang phân tích dữ liệu...');
      setTimeout(() => {
        if (stateToken === uploadStateToken && loadingText) {
          loadingText.textContent = 'Đang trích xuất sóng hài...';
        }
      }, 1800);
      setTimeout(() => {
        if (stateToken === uploadStateToken && loadingText) {
          loadingText.textContent = 'Tính toán THD và sự kiện...';
        }
      }, 3600);

      fetch('/api/kew/upload', {
        method: 'POST',
        body: formData
      })
        .then(res => {
          if (!res.ok) return res.json().then(e => { throw new Error(e.error || 'Server error'); });
          return res.json();
        })
        .then(data => {
          uploadStateToken++;
          setUploadState('success', 'Đã phân tích xong. Biểu đồ đã được cập nhật.');
          renderAll(data);
          // Run phase detection in background
          if (lastFormData) detectAndShowGenPanel(lastFormData);
        })
        .catch(err => {
          uploadStateToken++;
          showError(err.message);
        });
    }

    // ── Phase Detection & Gen Panel ─────────────────
    function detectAndShowGenPanel(formData) {
      // Build a fresh FormData from the uploaded files
      fetch('/api/kew/detect', { method: 'POST', body: formData })
        .then(r => r.json())
        .then(data => {
          if (!data.results || data.results.length === 0) return;
          const info = data.results[0];
          if (!info.missing || info.missing.length === 0) return;

          // Build phase status badges
          const phaseRow = document.getElementById('gen-phase-status');
          phaseRow.innerHTML = ['1', '2', '3'].map(p => {
            const isMissing = info.missing.includes(p);
            return `<span class="phase-badge ${isMissing ? 'phase-missing' : 'phase-ok'}">Pha ${p} ${isMissing ? '✗ Thiếu' : '✓ OK'}</span>`;
          }).join('');

          document.getElementById('gen-desc').textContent =
            `Phát hiện pha ${info.missing.join(', ')} bị khuyết (pha tham chiếu: ${info.ref_phase}). Gen dữ liệu mô phỏng thực tế bằng thuật toán Ornstein-Uhlenbeck.`;

          document.getElementById('gen-panel').style.display = 'flex';
        })
        .catch(() => { }); // silent fail
    }

    function runGenData() {
      if (!lastFormData) return;
      const btn = document.getElementById('btn-gen');
      const spinner = document.getElementById('gen-spinner');
      const btnText = document.getElementById('gen-btn-text');

      btn.disabled = true;
      spinner.style.display = 'inline-block';
      btnText.textContent = 'Đang xử lý...';

      fetch('/api/kew/fix', { method: 'POST', body: lastFormData })
        .then(res => {
          if (!res.ok) return res.json().then(e => { throw new Error(e.error || 'Lỗi hệ thống'); });
          return res.blob();
        })
        .then(blob => {
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'KEW_Fixed_Data.zip';
          document.body.appendChild(a);
          a.click();
          a.remove();

          btn.disabled = false;
          spinner.style.display = 'none';
          btnText.textContent = '✅ Tải lại';
        })
        .catch(err => {
          alert('Lỗi: ' + err.message);
          btn.disabled = false;
          spinner.style.display = 'none';
          btnText.textContent = '🛠 Gen dữ liệu & Tải ZIP';
        });
    }

    // ── Correction Tab Logic ─────────────────────────────────────────
    function getCorrConfig() {
      const config = {};
      document.querySelectorAll('.corr-input').forEach(inp => {
        const g = inp.dataset.g;
        const t = inp.dataset.t;
        if (!config[g]) config[g] = { multiply: 1.0, offset: 0.0 };
        config[g][t] = parseFloat(inp.value) || (t === 'multiply' ? 1.0 : 0.0);
      });
      return config;
    }

    function _applyCorr(v, mul, off) { return v * mul + off; }

    const CORR_GROUP_MAP = {
      'V1[V]': 'V', 'V2[V]': 'V', 'V3[V]': 'V',
      'A1[A]': 'A', 'A2[A]': 'A', 'A3[A]': 'A',
      'P1[W]': 'P', 'P2[W]': 'P', 'P3[W]': 'P', 'P[W]': 'P',
      'Q1[var]': 'Q', 'Q2[var]': 'Q', 'Q3[var]': 'Q', 'Q[var]': 'Q',
      'S1[VA]': 'S', 'S2[VA]': 'S', 'S3[VA]': 'S', 'S[VA]': 'S',
    };

    function updateCorrPreview() {
      if (!DATA) return;
      const cfg = getCorrConfig();
      const s = DATA.series;
      const _n = (v, d = 1) => v != null ? v.toFixed(d) : '—';

      // Update preview table
      const groupAvgs = { A: [], V: [], P: [], Q: [], S: [] };
      Object.entries(s).forEach(([ch, d]) => {
        const g = CORR_GROUP_MAP[ch];
        if (g && groupAvgs[g] !== undefined && d.avg_fundamental != null) {
          groupAvgs[g].push(d.avg_fundamental);
        }
      });
      Object.entries(groupAvgs).forEach(([g, vals]) => {
        if (!vals.length) return;
        const origAvg = vals.reduce((a, b) => a + b, 0) / vals.length;
        const c = cfg[g] || { multiply: 1.0, offset: 0.0 };
        const corrAvg = _applyCorr(origAvg, c.multiply, c.offset);
        const origEl = document.getElementById(`orig-${g}`);
        const prevEl = document.getElementById(`prev-${g}`);
        if (origEl) origEl.textContent = _n(origAvg, 3);
        if (prevEl) prevEl.textContent = _n(corrAvg, 3);
      });

      // Update preview charts
      const corrSeries = (chs) => chs.map(ch => {
        if (!s[ch]) return null;
        const g = CORR_GROUP_MAP[ch];
        const c = cfg[g] || { multiply: 1.0, offset: 0.0 };
        const [ts, vals] = subsamplePair(s[ch].timestamps, s[ch].fundamental);
        return { ts, vals: vals.map(v => v != null ? _applyCorr(v, c.multiply, c.offset) : null), ch };
      }).filter(Boolean);

      const vSeries = corrSeries(['V1[V]', 'V2[V]', 'V3[V]']);
      if (vSeries.length) {
        createOrUpdate('corrVoltChart', {
          type: 'line',
          data: { labels: fmtTS(vSeries[0].ts), datasets: vSeries.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)])) },
          options: cloneDefaults()
        });
      }

      const aSeries = corrSeries(['A1[A]', 'A2[A]', 'A3[A]']);
      if (aSeries.length) {
        createOrUpdate('corrCurrChart', {
          type: 'line',
          data: { labels: fmtTS(aSeries[0].ts), datasets: aSeries.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)])) },
          options: cloneDefaults()
        });
      }

      const pSeries = corrSeries(['P1[W]', 'P2[W]', 'P3[W]']);
      if (pSeries.length) {
        createOrUpdate('corrPowerChart', {
          type: 'line',
          data: { labels: fmtTS(pSeries[0].ts), datasets: pSeries.map(d => makeLineDataset(d.ch, d.vals, COLORS[d.ch.substring(0, 2)])) },
          options: cloneDefaults()
        });
      }
    }

    function applyCTPreset() {
      const real = parseFloat(document.getElementById('ct-real').value);
      const set = parseFloat(document.getElementById('ct-set').value);
      if (!real || !set || set === 0) return;
      const factor = real / set;
      document.querySelectorAll('.corr-input[data-g="A"][data-t="multiply"]').forEach(el => { el.value = factor.toFixed(4); });
      updateCorrPreview();
    }

    function applyVTPreset() {
      const real = parseFloat(document.getElementById('vt-real').value);
      const set = parseFloat(document.getElementById('vt-set').value);
      if (!real || !set || set === 0) return;
      const factor = real / set;
      document.querySelectorAll('.corr-input[data-g="V"][data-t="multiply"]').forEach(el => { el.value = factor.toFixed(4); });
      updateCorrPreview();
    }

    // ── Quick Calculator Logic ──────────────────────────────────────
    function updateCalcTarget() {
      const target = document.getElementById('calc-target').value;
      // Disable the target field, enable others
      ['u', 'i', 'p'].forEach(key => {
        const el = document.getElementById(`calc-${key}`);
        el.disabled = (key === target);
        if (key === target) el.style.background = 'var(--surface2)';
        else el.style.background = '';
      });
      quickCalc();
    }

    function quickCalc() {
      const target = document.getElementById('calc-target').value;
      const u = parseFloat(document.getElementById('calc-u').value);
      const i = parseFloat(document.getElementById('calc-i').value);
      const p = parseFloat(document.getElementById('calc-p').value);
      const pf = parseFloat(document.getElementById('calc-pf').value) || 0.85;
      const is3p = document.getElementById('calc-type').value === '3p';
      const root3 = is3p ? Math.sqrt(3) : 1;
      const hint = document.getElementById('calc-hint');

      if (target === 'p') {
        if (!isNaN(u) && !isNaN(i)) {
          const pVal = (root3 * u * i * pf) / 1000;
          document.getElementById('calc-p').value = pVal.toFixed(2);
          hint.textContent = `Kết quả: P = ${pVal.toFixed(2)} kW`;
        }
      } else if (target === 'i') {
        if (!isNaN(p) && !isNaN(u) && u > 0) {
          const iVal = (p * 1000) / (root3 * u * pf);
          document.getElementById('calc-i').value = iVal.toFixed(2);
          hint.textContent = `Kết quả: I = ${iVal.toFixed(2)} A`;
        }
      } else if (target === 'u') {
        if (!isNaN(p) && !isNaN(i) && i > 0) {
          const uVal = (p * 1000) / (root3 * i * pf);
          document.getElementById('calc-u').value = uVal.toFixed(0);
          hint.textContent = `Kết quả: U = ${uVal.toFixed(0)} V`;
        }
      }
    }

    function resetCalc() {
      ['calc-u', 'calc-i', 'calc-p'].forEach(id => document.getElementById(id).value = '');
      document.getElementById('calc-hint').textContent = '';
      updateCalcTarget();
    }

    function autoDetectCorrection() {
      if (!DATA) {
        alert('Vui lòng upload file KEW trước.');
        return;
      }
      const s = DATA.series;
      const dev = DATA.summary?.device || {};
      const findings = [];
      const suggestions = {};

      // ── 1. Voltage detection ─────────────────────────────────────
      const nomStr = dev['NOMINAL VOLTAGE'] || '';
      const nomVolt = parseFloat(nomStr);   // e.g. 400
      const vAvgs = ['V1[V]', 'V2[V]', 'V3[V]'].map(c => s[c]?.avg_fundamental).filter(v => v != null && v > 0);
      if (nomVolt > 0 && vAvgs.length > 0) {
        const measuredVolt = vAvgs.reduce((a, b) => a + b, 0) / vAvgs.length;
        const vtFactor = nomVolt / measuredVolt;

        if (Math.abs(vtFactor - 1) > 0.04) {  // >4% deviation → likely error
          const pct = ((vtFactor - 1) * 100).toFixed(1);
          findings.push({
            icon: '⚡', color: 'var(--accent)',
            title: 'Phát hiện lệch điện áp',
            desc: `Điện áp đo được trung bình: <strong>${measuredVolt.toFixed(1)} V</strong>, điện áp định mức trong cấu hình: <strong>${nomVolt} V</strong>.<br>
               Đề xuất: VT Factor = ${nomVolt} ÷ ${measuredVolt.toFixed(1)} = <strong>${vtFactor.toFixed(4)}</strong> (${pct > 0 ? '+' : ''}${pct}%).`
          });
          suggestions.V = { multiply: vtFactor, offset: 0 };
        } else {
          findings.push({ icon: '✅', color: 'var(--accent3)', title: 'Điện áp bình thường', desc: `Trung bình ${measuredVolt.toFixed(1)} V ≈ định mức ${nomVolt} V. Không cần hiệu chỉnh.` });
        }
      }

      // ── 2. Current detection ──────────────────────────────────────
      const currRangeStr = (dev['CURRENT RANGE'] || '').split(",'")[0].trim(); // e.g. "3000 A"
      const currRange = parseFloat(currRangeStr);  // 3000
      const aAvgs = ['A1[A]', 'A2[A]', 'A3[A]'].map(c => s[c]?.avg_fundamental).filter(v => v != null && v > 0);
      const pAvgs = ['P1[W]', 'P2[W]', 'P3[W]'].map(c => s[c]?.avg_fundamental).filter(v => v != null && v > 0);
      if (aAvgs.length > 0 && vAvgs.length > 0 && pAvgs.length > 0) {
        const measuredA = aAvgs.reduce((a, b) => a + b, 0) / aAvgs.length;
        const measuredV = vAvgs.reduce((a, b) => a + b, 0) / vAvgs.length;
        const measuredP = pAvgs.reduce((a, b) => a + b, 0) / pAvgs.length;
        // Apparent power per phase = V × A
        const calcS = measuredV * measuredA;
        // Calculated PF from measured P and derived S
        const derivedPF = calcS > 0 ? measuredP / calcS : null;

        if (derivedPF !== null && (derivedPF > 1.05 || derivedPF < 0.1)) {
          // PF out of range → likely current scaling error
          // If PF >> 1: currents are too small (CT multiplier too low)
          // Estimate: expected A = P / (V × typicalPF ≈ 0.85)
          const expectedA = measuredP / (measuredV * 0.85);
          const ctFactor = expectedA / measuredA;
          if (ctFactor > 1.5 || ctFactor < 0.5) {
            // Round to nearest common ratio
            const commonRatios = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000];
            const nearest = commonRatios.reduce((prev, r) => Math.abs(r - ctFactor) < Math.abs(prev - ctFactor) ? r : prev);
            findings.push({
              icon: '🔌', color: 'var(--accent3)',
              title: 'Phát hiện bất thường dòng điện',
              desc: `Hệ số công suất tính được: <strong>${derivedPF.toFixed(3)}</strong> (bất thường).<br>
                 Dòng đo: ${measuredA.toFixed(2)} A — Dòng ước tính: ${expectedA.toFixed(2)} A.<br>
                 Đề xuất CT Factor ≈ <strong>${nearest}×</strong> (tương đương CT thực / CT cài = ${nearest}).`
            });
            suggestions.A = { multiply: nearest, offset: 0 };
            // Power follows: P = V × I, so P factor = CT factor
            suggestions.P = { multiply: nearest, offset: 0 };
            suggestions.Q = { multiply: nearest, offset: 0 };
            suggestions.S = { multiply: nearest, offset: 0 };
          }
        } else if (derivedPF !== null) {
          findings.push({ icon: '✅', color: 'var(--accent3)', title: 'Dòng điện bình thường', desc: `Hệ số công suất tính được: ${derivedPF.toFixed(3)} — trong phạm vi bình thường.` });
        }
      }

      // ── 3. Cross-check CT range vs measured current ───────────────
      if (currRange > 0 && aAvgs.length > 0) {
        const maxMeasuredA = Math.max(...aAvgs);
        const rangeUsage = maxMeasuredA / currRange;
        if (rangeUsage < 0.001) {
          findings.push({
            icon: '⚠️', color: 'var(--accent4)',
            title: 'Dòng điện cực kỳ thấp so với dải đo',
            desc: `Dải đo cấu hình: ${currRange} A — Dòng trung bình: ${aAvgs[0].toFixed(4)} A (${(rangeUsage * 100).toFixed(4)}% của dải đo).<br>
               Có thể CT ratio bị cài sai. Kiểm tra tỉ số CT thực tế.`
          });
        }
      }

      if (!findings.length) {
        findings.push({ icon: 'ℹ️', color: 'var(--accent)', title: 'Không phát hiện bất thường rõ ràng', desc: 'Các thông số điện áp và dòng điện nằm trong phạm vi bình thường. Nếu cần hiệu chỉnh, nhập thủ công bên dưới.' });
      }

      // ── Render result card ───────────────────────────────────────
      const resultEl = document.getElementById('auto-detect-result');
      resultEl.style.display = 'block';
      resultEl.innerHTML = `
    <div class="chart-card" style="border-color:var(--accent3)20">
      <h3>🧠 Kết quả nhận dạng tự động</h3>
      <div style="display:flex;flex-direction:column;gap:10px;margin-top:0.5rem">
        ${findings.map(f => `
          <div style="display:flex;gap:12px;align-items:flex-start;padding:10px 12px;background:var(--surface2);border-radius:10px">
            <span style="font-size:1.2rem">${f.icon}</span>
            <div style="flex:1">
              <div style="font-size:0.82rem;font-weight:600;color:${f.color};margin-bottom:4px">${f.title}</div>
              <div style="font-size:0.76rem;color:var(--text-muted);line-height:1.6">${f.desc}</div>
            </div>
          </div>
        `).join('')}
      </div>
      ${Object.keys(suggestions).length > 0 ? `
        <div style="margin-top:1rem;display:flex;gap:8px;align-items:center">
          <button class="btn" style="border-color:var(--accent3);color:var(--accent3);font-weight:600" onclick="applyAutoSuggestions(${JSON.stringify(suggestions).replace(/"/g, '&quot;')})">
            ✅ Áp dụng đề xuất
          </button>
          <span style="font-size:0.75rem;color:var(--text-muted)">Sẽ điền vào bảng thông số bên dưới</span>
        </div>
      ` : ''}
    </div>`;
    }

    function applyAutoSuggestions(suggestions) {
      Object.entries(suggestions).forEach(([g, c]) => {
        document.querySelectorAll(`.corr-input[data-g="${g}"][data-t="multiply"]`).forEach(el => { el.value = c.multiply.toFixed(4); });
        document.querySelectorAll(`.corr-input[data-g="${g}"][data-t="offset"]`).forEach(el => { el.value = (c.offset || 0).toFixed(2); });
      });
      updateCorrPreview();
    }

    function resetCorrections() {
      document.querySelectorAll('.corr-input[data-t="multiply"]').forEach(el => { el.value = '1'; });
      document.querySelectorAll('.corr-input[data-t="offset"]').forEach(el => { el.value = '0'; });
      ['ct-real', 'ct-set', 'vt-real', 'vt-set'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
      updateCorrPreview();
    }

    function applyAndDownloadCorrection() {
      if (!lastFormData) {
        document.getElementById('corr-error').textContent = 'Vui lòng upload file KEW trước.';
        document.getElementById('corr-error').style.display = 'block';
        return;
      }
      const cfg = getCorrConfig();
      // Only send groups that have non-default values
      const effective = {};
      Object.entries(cfg).forEach(([g, c]) => {
        if (c.multiply !== 1.0 || c.offset !== 0.0) effective[g] = c;
      });
      if (!Object.keys(effective).length) {
        document.getElementById('corr-error').textContent = 'Chưa nhập thông số hiệu chỉnh nào.';
        document.getElementById('corr-error').style.display = 'block';
        return;
      }
      document.getElementById('corr-error').style.display = 'none';
      const btn = document.getElementById('btn-apply-corr');
      const spinner = document.getElementById('corr-spinner');
      const btnText = document.getElementById('corr-btn-text');
      btn.disabled = true;
      spinner.style.display = 'inline-block';
      btnText.textContent = 'Đang xử lý...';

      const fd = new FormData(lastFormData); // clone won't work, re-append
      const sendFd = lastFormData;
      // Append correction config as form field
      const fd2 = new FormData();
      for (const [k, v] of sendFd.entries()) fd2.append(k, v);
      fd2.append('corrections', JSON.stringify(effective));

      fetch('/api/kew/correct', { method: 'POST', body: fd2 })
        .then(res => {
          if (!res.ok) return res.json().then(e => { throw new Error(e.error || 'Lỗi hệ thống'); });
          return res.blob();
        })
        .then(blob => {
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a'); a.href = url; a.download = 'KEW_Corrected.zip';
          document.body.appendChild(a); a.click(); a.remove();
          btn.disabled = false; spinner.style.display = 'none'; btnText.textContent = '✅ Tải lại';
        })
        .catch(err => {
          document.getElementById('corr-error').textContent = 'Lỗi: ' + err.message;
          document.getElementById('corr-error').style.display = 'block';
          btn.disabled = false; spinner.style.display = 'none'; btnText.textContent = '📥 Áp dụng & Tải ZIP';
        });
    }
