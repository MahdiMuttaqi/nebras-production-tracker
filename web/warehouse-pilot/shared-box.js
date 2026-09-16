/* Nebras Warehouse — shared-box view/withdrawal enhancement.
   Isolated from the production tracker and from unrelated warehouse flows. */
(() => {
  'use strict';

  let sharedOutSource = null;

  function sharedMemberLocations(x) {
    const bits = String(x?.display || '').split('/').map(v => v.trim());
    if (bits.length < 3) return [];
    const shelf = bits[0];
    const row = bits[1];
    const boxes = bits.slice(2).join('/').split('-').map(v => v.trim()).filter(Boolean);
    return boxes.map(box => boxLocation(shelf, row, box)).filter(Boolean);
  }

  function locationInventoryRows(location) {
    const l = norm(location);
    return db.stock.filter(x => x.q > 0 && (
      norm(x.l) === l ||
      ((x.shared || x.unassigned) && sharedMemberLocations(x).includes(l))
    ));
  }

  function renderScannedBox() {
    const l = norm(document.querySelector('#findLocation').value);
    const rows = locationInventoryRows(l);
    const codeCount = new Set(rows.map(x => x.p)).size;

    document.querySelector('#searchResults').innerHTML = `<div class="result"><div class="result-head"><b>${html(l || 'مکان انتخاب نشده')}</b><span class="badge">${fa(codeCount)} کد مرتبط</span></div>${rows.length ? rows.map(x => {
      const shared = !!(x.shared || x.unassigned);
      const name = product(x.p)?.name || '';
      const detail = shared
        ? `<small style="display:block;color:#8a6615;margin-top:4px">موجودی مشترک: ${html(x.display || x.l)}</small>`
        : '<small style="display:block;color:#647984;margin-top:4px">موجودی همین جعبه</small>';
      const buttons = shared
        ? `<div style="display:flex;flex-direction:column;gap:6px"><button class="primary" onclick="quickOutShared('${x.p}','${x.l}','${l}')">برداشت از مشترک</button><button class="secondary" onclick="openSplit('${x.p}','${x.l}')">تفکیک</button></div>`
        : `<button class="secondary" onclick="quickOut('${x.p}','${l}')">برداشت</button>`;
      return `<div class="result-row"><span><b>${html(x.p)} — ${html(name)}</b>${detail}</span><span class="qty"><small>${shared ? 'کل مشترک' : 'موجودی'}</small>${fa(x.q)}</span>${buttons}</div>`;
    }).join('') : '<div class="empty">برای این جعبه موجودی مستقیم یا مشترک ثبت نشده است.</div>'}</div>`;
  }

  window.quickOut = (p, l) => {
    sharedOutSource = null;
    document.querySelector('#outProduct').value = p;
    document.querySelector('#outLoc').value = l;
    switchTab('out');
    showAvailable();
  };

  window.quickOutShared = (p, source, physical) => {
    sharedOutSource = {p: norm(p), l: norm(source), physical: norm(physical)};
    document.querySelector('#outProduct').value = p;
    document.querySelector('#outLoc').value = physical;
    switchTab('out');
    showAvailable();
    document.querySelector('#outQty').focus();
  };

  window.showAvailable = function () {
    const p = norm(document.querySelector('#outProduct').value);
    const l = norm(document.querySelector('#outLoc').value);
    const el = document.querySelector('#outAvailable');
    if (!p) {
      el.innerHTML = '';
      return;
    }

    const sharedSelected = sharedOutSource && sharedOutSource.p === p && sharedOutSource.physical === l
      ? item(p, sharedOutSource.l)
      : null;

    if (sharedSelected?.q > 0) {
      el.innerHTML = `<div class="shared-box"><strong>برداشت از موجودی مشترک</strong><div>جعبه اسکن‌شده: <span class="location-code">${html(l)}</span></div><div style="margin-top:6px">گروه مشترک: ${html(sharedSelected.display || sharedSelected.l)}</div><div class="split-summary"><span class="badge">موجودی مشترک کل: ${fa(sharedSelected.q)}</span></div><small>این برداشت از موجودی مشترک کم می‌شود و نیازی به تفکیک قبلی ندارد.</small></div>`;
      return;
    }

    const selected = l ? item(p, l) : null;
    if (selected?.q > 0) {
      el.innerHTML = `<div class="notice" style="margin-top:12px;margin-bottom:0">موجودی مکان انتخاب‌شده: <b>${fa(selected.q)}</b></div>`;
      return;
    }

    const rows = db.stock.filter(x => x.p === p && x.q > 0 && locationValid(x.l));
    const warning = l ? '<div class="notice" style="margin-top:12px">این محصول در مکان واردشده موجود نیست؛ یکی از مکان‌های زیر را انتخاب کنید.</div>' : '';
    el.innerHTML = rows.length
      ? warning + `<div class="out-locations"><div class="out-locations-title">این محصول در ${fa(rows.length)} مکان موجود است؛ مکان برداشت را انتخاب یا QR آن را اسکن کنید.</div>${rows.map(x => `<button type="button" class="out-location" onclick="selectOutLocation('${html(x.l)}')"><span class="location-code">${html(x.display || x.l)}</span><span class="qty"><small>موجودی</small>${fa(x.q)}</span></button>`).join('')}</div>`
      : '<div class="notice" style="margin-top:12px;margin-bottom:0">برای این کد، لوکیشن فعال و قابل برداشت پیدا نشد.</div>';
  };

  window.selectOutLocation = l => {
    sharedOutSource = null;
    document.querySelector('#outLoc').value = l;
    showAvailable();
    document.querySelector('#outQty').focus();
  };

  document.querySelector('#showBoxBtn').onclick = renderScannedBox;

  document.querySelector('#outProduct').oninput = () => {
    if (sharedOutSource && sharedOutSource.p !== norm(document.querySelector('#outProduct').value)) sharedOutSource = null;
    showAvailable();
  };

  document.querySelector('#outLoc').oninput = () => {
    if (sharedOutSource && sharedOutSource.physical !== norm(document.querySelector('#outLoc').value)) sharedOutSource = null;
    showAvailable();
  };

  document.querySelector('#outSubmit').onclick = () => run(() => {
    const l = checkLoc('#outLoc');
    const p = norm(document.querySelector('#outProduct').value);
    const q = Number(document.querySelector('#outQty').value);
    if (!p || q <= 0) throw Error('کد محصول و تعداد خروج را کامل کنید.');

    let shared = null;
    if (sharedOutSource && sharedOutSource.p === p && sharedOutSource.physical === l) {
      shared = item(p, sharedOutSource.l);
      if (!shared || shared.q <= 0) throw Error('موجودی مشترک تغییر کرده است؛ ابتدا از اکسل تازه‌سازی کنید.');
    }

    const source = shared ? sharedOutSource.l : l;
    const changed = changeStock(p, source, -q);
    const wasShared = !!shared;

    // For shared stock, FROM is the shared pool and TO records the scanned physical box as context.
    log('خروج', p, q, source, wasShared ? l : '', document.querySelector('#outNote').value, changed.before, changed.after);
    sharedOutSource = null;
    save();
    showAvailable();
    toast(wasShared ? `خروج ${fa(q)} عدد از موجودی مشترک ثبت شد.` : `خروج ${fa(q)} عدد ثبت شد.`);
  });
})();
