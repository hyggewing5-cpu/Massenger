/**
 * ระบบบันทึกการใช้เชื้อเพลิง — Google Apps Script (Web App)
 *
 * วางไฟล์นี้ในสคริปต์ที่ผูกกับ Google Sheet (Extensions > Apps Script)
 * จากนั้นรันฟังก์ชัน setup() หนึ่งครั้ง แล้ว Deploy เป็น Web App
 * ขั้นตอนทั้งหมดอยู่ในไฟล์ README ที่แนบมา
 *
 * หลักการ
 *  - ตรวจรหัส 6 หลักที่เซิร์ฟเวอร์เท่านั้น เก็บเป็นแฮช SHA-256 + salt + pepper
 *  - เข้าสำเร็จได้โทเค็นอายุ 6 ชั่วโมง และตรวจสิทธิ์ซ้ำทุกคำขอจากชีต Users
 *  - ใส่รหัสผิดครบ 5 ครั้งต่ออุปกรณ์ ล็อก 10 นาที
 *  - เฉพาะหัวหน้างาน (sup) จัดการรถและผู้ใช้ได้ ลูกน้อง (staff) บันทึกและดูรายงานได้
 */

var SCHEMA = {
  Vehicles: ['vehicle_id', 'name', 'plate', 'type', 'fuel_type', 'std_kmpl', 'tank_l', 'odo_start', 'is_main', 'active'],
  FuelLogs: ['log_id', 'date', 'vehicle_id', 'driver', 'odo_start', 'odo_end', 'liters', 'price_per_l', 'trips', 'note', 'created_by', 'created_at'],
  Users: ['user_id', 'name', 'role', 'pin_hash', 'salt', 'active', 'pin_changed_at', 'created_at'],
  Trips: ['trip_id', 'vehicle_id', 'driver', 'user_id', 'start_time', 'end_time', 'start_lat', 'start_lng', 'end_lat', 'end_lng',
    'gps_km', 'odo_start', 'odo_end', 'est_liters', 'points', 'gaps', 'route', 'dest', 'note', 'status', 'created_at']
};
var SUP_ONLY = ['saveVehicle', 'deleteVehicle', 'addUser', 'setUserActive', 'resetUserPin', 'deleteUser'];
var TOKEN_TTL = 21600;   // วินาที (6 ชั่วโมง คือค่าสูงสุดของ CacheService)
var MAX_FAILS = 5;       // ผิดกี่ครั้งต่ออุปกรณ์จึงล็อก
var GLOBAL_FAILS = 40;   // ผิดรวมทุกอุปกรณ์กี่ครั้งจึงล็อกทั้งระบบ
var LOCK_SEC = 600;      // ระยะเวลาล็อก (วินาที)
var TZ = 'Asia/Bangkok';

/* ================= ติดตั้งครั้งแรก ================= */

function setup() {
  var ss = _ss();
  Object.keys(SCHEMA).forEach(function (name) {
    var head = SCHEMA[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var logs = ss.getSheetByName('FuelLogs');
  logs.getRange(2, 2, Math.max(1, logs.getMaxRows() - 1), 1).setNumberFormat('@');
  var trips = ss.getSheetByName('Trips');
  trips.getRange(2, 5, Math.max(1, trips.getMaxRows() - 1), 2).setNumberFormat('@'); // start_time, end_time

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN_PEPPER')) {
    props.setProperty('PIN_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }

  if (_rows('Vehicles').length === 0) {
    _append('Vehicles', {
      vehicle_id: _newId('v', 'Vehicles', 'vehicle_id'), name: 'Honda PCX 125', plate: '', type: 'มอเตอร์ไซค์',
      fuel_type: 'แก๊สโซฮอล์ 91', std_kmpl: 45, tank_l: 8, odo_start: 0, is_main: true, active: true
    });
  }
  if (_rows('Users').length === 0) {
    var pin = _newPin();
    _createUser('หัวหน้างาน', 'sup', pin);
    _showPin('สร้างบัญชีหัวหน้างานแล้ว\nรหัสเข้าระบบ: ' + pin + '\nเปลี่ยนรหัสได้ในแท็บ "ผู้ใช้งาน" หลังเข้าสู่ระบบ');
  } else {
    _showPin('ตรวจโครงสร้างชีตเรียบร้อยแล้ว (มีผู้ใช้อยู่แล้ว จึงไม่สร้างบัญชีใหม่)');
  }
}

/** ใช้เมื่อหัวหน้างานลืมรหัส: รันจากตัวแก้ไขสคริปต์ แล้วดูรหัสใหม่ */
function resetSupervisorPin() {
  var sups = _rows('Users').filter(function (u) { return u.role === 'sup'; });
  var pin = _newPin();
  if (sups.length) {
    var salt = Utilities.getUuid();
    _update('Users', sups[0]._row, { pin_hash: _hash(pin, salt), salt: salt, active: true, pin_changed_at: Date.now() });
    _showPin('ตั้งรหัสใหม่ให้ ' + sups[0].name + ' แล้ว\nรหัสเข้าระบบ: ' + pin);
  } else {
    _createUser('หัวหน้างาน', 'sup', pin);
    _showPin('สร้างบัญชีหัวหน้างานแล้ว\nรหัสเข้าระบบ: ' + pin);
  }
}

function _showPin(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* ไม่มีหน้าจอ ใช้ดูใน Execution log */ }
}

/* ================= Web App ================= */

function doGet() {
  return _json({ ok: true, service: 'fuel-log' });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    out = _route(req);
    out.ok = true;
  } catch (err) {
    out = { ok: false, code: err.code || 'server', error: err.code ? err.message : 'เกิดข้อผิดพลาดบนเซิร์ฟเวอร์ ลองใหม่อีกครั้ง' };
    if (!err.code) console.error(err.stack || err);
  }
  return _json(out);
}

function _json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function _route(req) {
  var a = String(req.action || '');
  if (a === 'login') return _login(req);
  if (a === 'logout') {
    if (req.token) CacheService.getScriptCache().remove('tk_' + req.token);
    return {};
  }
  var user = _auth(req.token, SUP_ONLY.indexOf(a) >= 0);
  switch (a) {
    case 'data': return _data(user);
    case 'addLog': return _addLog(user, req);
    case 'startTrip': return _startTrip(user, req);
    case 'finishTrip': return _finishTrip(user, req);
    case 'cancelTrip': return _cancelTrip(user, req);
    case 'tripRoute': return _tripRoute(req);
    case 'saveVehicle': return _saveVehicle(req);
    case 'deleteVehicle': return _deleteVehicle(req);
    case 'addUser': return _addUser(req);
    case 'setUserActive': return _setUserActive(user, req);
    case 'resetUserPin': return _resetUserPin(user, req);
    case 'deleteUser': return _deleteUser(user, req);
  }
  throw _err('ไม่รู้จักคำสั่งนี้');
}

/* ================= ล็อกอินและสิทธิ์ ================= */

function _login(p) {
  var pin = String(p.pin || '');
  var dev = String(p.device || 'unknown').replace(/[^A-Za-z0-9_-]/g, '').substring(0, 48) || 'unknown';
  if (!/^\d{6}$/.test(pin)) throw _err('รหัสต้องเป็นตัวเลข 6 หลัก');
  return _withLock(function () {
    var c = CacheService.getScriptCache();
    if (c.get('lock_' + dev) || c.get('glock')) {
      throw _err('ใส่รหัสผิดหลายครั้ง กรุณารอประมาณ 10 นาทีแล้วลองใหม่', 'locked');
    }
    var found = null;
    _rows('Users').forEach(function (u) {
      if (_bool(u.active) && u.salt && _hash(pin, String(u.salt)) === String(u.pin_hash)) found = u;
    });
    if (!found) {
      var n = Number(c.get('fc_' + dev) || 0) + 1;
      var g = Number(c.get('gfc') || 0) + 1;
      c.put('fc_' + dev, String(n), LOCK_SEC);
      c.put('gfc', String(g), LOCK_SEC);
      if (g >= GLOBAL_FAILS) c.put('glock', '1', LOCK_SEC);
      if (n >= MAX_FAILS) {
        c.put('lock_' + dev, '1', LOCK_SEC);
        c.remove('fc_' + dev);
        throw _err('ใส่รหัสผิดครบ ' + MAX_FAILS + ' ครั้ง ล็อก 10 นาที', 'locked');
      }
      throw _err('รหัสไม่ถูกต้อง เหลืออีก ' + (MAX_FAILS - n) + ' ครั้ง', 'badpin');
    }
    c.remove('fc_' + dev);
    return { token: _newToken(found.user_id) };
  });
}

function _newToken(uid) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('tk_' + token, JSON.stringify({ uid: uid, iat: Date.now() }), TOKEN_TTL);
  return token;
}

function _auth(token, needSup) {
  if (!token) throw _err('กรุณาเข้าสู่ระบบ', 'auth');
  var raw = CacheService.getScriptCache().get('tk_' + token);
  if (!raw) throw _err('หมดเวลาการเข้าใช้งาน กรุณาใส่รหัสใหม่', 'auth');
  var s = JSON.parse(raw);
  var u = _rows('Users').filter(function (x) { return x.user_id === s.uid; })[0];
  if (!u || !_bool(u.active) || Number(u.pin_changed_at || 0) > s.iat) {
    throw _err('บัญชีนี้ถูกปิดหรือเปลี่ยนรหัสแล้ว กรุณาเข้าสู่ระบบใหม่', 'auth');
  }
  if (needSup && u.role !== 'sup') throw _err('เฉพาะหัวหน้างานเท่านั้น', 'forbidden');
  return u;
}

/* ================= ข้อมูล ================= */

function _data(user) {
  var out = {
    me: { id: user.user_id, name: user.name, role: user.role },
    vehicles: _vehicles(),
    logs: _logs(),
    trips: _trips(),
    drivers: _drivers()
  };
  if (user.role === 'sup') out.users = _users();
  return out;
}

function _vehicles() {
  return _rows('Vehicles').filter(function (v) { return _bool(v.active); }).map(function (v) {
    return {
      id: v.vehicle_id, name: String(v.name), plate: String(v.plate || ''), type: String(v.type),
      fuel: String(v.fuel_type), std: Number(v.std_kmpl), tank: Number(v.tank_l),
      odo0: Number(v.odo_start) || 0, main: _bool(v.is_main)
    };
  });
}

function _logs() {
  var cutoff = _dstr(new Date(Date.now() - 400 * 86400000));
  return _rows('FuelLogs').map(function (l) {
    return {
      id: Number(l.log_id), date: _dstr(l.date), vehicleId: l.vehicle_id, driver: String(l.driver),
      odoStart: Number(l.odo_start), odoEnd: Number(l.odo_end), liters: Number(l.liters),
      price: Number(l.price_per_l), trips: Number(l.trips) || 0, note: String(l.note || '')
    };
  }).filter(function (l) { return l.date >= cutoff; });
}

function _users() {
  return _rows('Users').map(function (u) {
    return { id: u.user_id, name: String(u.name), role: u.role, active: _bool(u.active) };
  });
}

function _drivers() {
  return _rows('Users').filter(function (u) { return _bool(u.active); }).map(function (u) { return String(u.name); });
}

function _addLog(user, p) {
  var l = p.log || {};
  var date = String(l.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw _err('วันที่ไม่ถูกต้อง');
  if (date > _dstr(new Date(Date.now() + 86400000))) throw _err('วันที่เติมเป็นวันในอนาคตไม่ได้');
  var odoS = Number(l.odoStart), odoE = Number(l.odoEnd), liters = Number(l.liters), price = Number(l.price);
  var trips = Math.max(0, Math.round(Number(l.trips) || 0));
  if (!(isFinite(odoS) && odoS >= 0 && isFinite(odoE) && odoE > odoS)) throw _err('เลขไมล์สิ้นสุดต้องมากกว่าเลขไมล์เริ่มต้น');
  if (!(liters > 0 && liters <= 1000)) throw _err('ปริมาณเชื้อเพลิงไม่ถูกต้อง');
  if (!(price >= 0 && price <= 500)) throw _err('ราคาต่อลิตรไม่ถูกต้อง');
  return _withLock(function () {
    var v = _rows('Vehicles').filter(function (x) { return x.vehicle_id === l.vehicleId && _bool(x.active); })[0];
    if (!v) throw _err('ไม่พบยานพาหนะนี้');
    // ลูกน้องบันทึกในชื่อตัวเองเท่านั้น หัวหน้างานเลือกผู้ขับได้
    var driver = user.name;
    if (user.role === 'sup' && _drivers().indexOf(String(l.driver)) >= 0) driver = String(l.driver);
    var id = 1;
    _rows('FuelLogs').forEach(function (r) { id = Math.max(id, Number(r.log_id) + 1); });
    _append('FuelLogs', {
      log_id: id, date: date, vehicle_id: v.vehicle_id, driver: driver, odo_start: odoS, odo_end: odoE,
      liters: Math.round(liters * 100) / 100, price_per_l: price, trips: trips, note: _txt(l.note, 300),
      created_by: user.user_id, created_at: new Date().toISOString()
    });
    return { id: id, driver: driver };
  });
}

/* ================= งานปฏิบัติงาน (เริ่มงาน / จบงาน + GPS) ================= */

function _ts(v) { return v instanceof Date ? v.getTime() : (Date.parse(String(v)) || 0); }

function _num(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function _ll(lat, lng) {
  lat = _num(lat); lng = _num(lng);
  if (lat === null || lng === null) return { lat: '', lng: '' };
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) throw _err('พิกัดไม่ถูกต้อง');
  return { lat: Math.round(lat * 1e6) / 1e6, lng: Math.round(lng * 1e6) / 1e6 };
}

function _tripOut(r) {
  return {
    id: Number(r.trip_id), vehicleId: r.vehicle_id, driver: String(r.driver), userId: r.user_id,
    start: _ts(r.start_time), end: r.end_time === '' ? null : _ts(r.end_time),
    startLat: _num(r.start_lat), startLng: _num(r.start_lng), endLat: _num(r.end_lat), endLng: _num(r.end_lng),
    km: _num(r.gps_km), odoStart: Number(r.odo_start), odoEnd: _num(r.odo_end), estL: _num(r.est_liters),
    points: Number(r.points) || 0, gaps: Number(r.gaps) || 0, dest: String(r.dest || ''), note: String(r.note || ''),
    status: String(r.status)
  };
}

/** งานที่ยังไม่จบ + งานที่จบแล้วใน 60 วันล่าสุด (ไม่ส่งเส้นทางมาด้วย ขอแยกด้วย tripRoute) */
function _trips() {
  var cutoff = Date.now() - 60 * 86400000;
  return _rows('Trips').filter(function (r) {
    return r.status === 'open' || (r.status === 'done' && _ts(r.start_time) >= cutoff);
  }).map(_tripOut);
}

function _findTrip(id) {
  return _rows('Trips').filter(function (r) { return Number(r.trip_id) === Number(id); })[0];
}

function _startTrip(user, p) {
  var t = p.trip || {};
  var odo = Number(t.odoStart);
  if (!(isFinite(odo) && odo >= 0)) throw _err('กรอกเลขไมล์เริ่มต้น');
  var pos = _ll(t.lat, t.lng);
  return _withLock(function () {
    var v = _rows('Vehicles').filter(function (x) { return x.vehicle_id === t.vehicleId && _bool(x.active); })[0];
    if (!v) throw _err('ไม่พบยานพาหนะนี้');
    var rows = _rows('Trips'), open = rows.filter(function (r) { return r.status === 'open'; });
    if (open.some(function (r) { return r.user_id === user.user_id; })) throw _err('คุณมีงานที่ยังไม่จบอยู่แล้ว');
    var busy = open.filter(function (r) { return r.vehicle_id === v.vehicle_id; })[0];
    if (busy) throw _err('รถคันนี้มีงานที่ยังไม่จบ (ผู้ขับ ' + busy.driver + ')');
    var id = 1;
    rows.forEach(function (r) { id = Math.max(id, Number(r.trip_id) + 1); });
    _append('Trips', {
      trip_id: id, vehicle_id: v.vehicle_id, driver: user.name, user_id: user.user_id,
      start_time: new Date().toISOString(), start_lat: pos.lat, start_lng: pos.lng, odo_start: odo,
      dest: _txt(t.dest, 120), status: 'open', created_at: new Date().toISOString()
    });
    return { trip: _tripOut(_findTrip(id)) };
  });
}

function _finishTrip(user, p) {
  var t = p.trip || {}, id = Number(t.id);
  var odoEnd = Number(t.odoEnd), km = _num(t.gpsKm), pos = _ll(t.lat, t.lng);
  var route = String(t.route || '');
  if (route.length > 20000 || !/^[\x3f-\x7e]*$/.test(route)) route = '';
  var points = Math.max(0, Math.round(Number(t.points) || 0)), gaps = Math.max(0, Math.round(Number(t.gaps) || 0));
  if (km !== null && !(km >= 0 && km <= 3000)) throw _err('ระยะทางจาก GPS ไม่ถูกต้อง');
  return _withLock(function () {
    var row = _findTrip(id);
    if (!row || row.status !== 'open') throw _err('ไม่พบงานที่กำลังทำอยู่นี้ (อาจถูกจบหรือยกเลิกไปแล้ว)');
    if (row.user_id !== user.user_id && user.role !== 'sup') throw _err('งานนี้เป็นของผู้ขับคนอื่น');
    var odoStart = Number(row.odo_start);
    if (!(isFinite(odoEnd) && odoEnd >= odoStart)) throw _err('เลขไมล์สิ้นสุดต้องไม่น้อยกว่าเลขไมล์เริ่มต้น');
    var veh = _rows('Vehicles').filter(function (v) { return v.vehicle_id === row.vehicle_id; })[0];
    var used = (km !== null && km > 0) ? km : (odoEnd - odoStart);
    var std = veh ? Number(veh.std_kmpl) : 0;
    var est = std > 0 ? Math.round(used / std * 100) / 100 : 0;
    // เวลาจบ: ใช้เวลาจากเครื่องถ้าสมเหตุสมผล (กรณีส่งช้าเพราะไม่มีสัญญาณ) ไม่เช่นนั้นใช้เวลาเซิร์ฟเวอร์
    var now = Date.now(), end = Number(t.endTime);
    if (!(end >= _ts(row.start_time) && end <= now + 60000)) end = now;
    _update('Trips', row._row, {
      end_time: new Date(end).toISOString(), end_lat: pos.lat, end_lng: pos.lng,
      gps_km: km === null ? '' : Math.round(km * 100) / 100, odo_end: odoEnd, est_liters: est,
      points: points, gaps: gaps, route: route, note: _txt(t.note, 200), status: 'done'
    });
    return { trip: _tripOut(_findTrip(id)) };
  });
}

function _cancelTrip(user, p) {
  return _withLock(function () {
    var row = _findTrip(p.id);
    if (!row || row.status !== 'open') throw _err('ไม่พบงานที่กำลังทำอยู่นี้');
    if (row.user_id !== user.user_id && user.role !== 'sup') throw _err('งานนี้เป็นของผู้ขับคนอื่น');
    _update('Trips', row._row, { status: 'cancelled', end_time: new Date().toISOString() });
    return {};
  });
}

function _tripRoute(p) {
  var row = _findTrip(p.id);
  if (!row) throw _err('ไม่พบงานนี้');
  return { route: String(row.route || '') };
}

/* ================= ยานพาหนะ (หัวหน้างาน) ================= */

function _saveVehicle(p) {
  var v = p.vehicle || {};
  var name = _txt(v.name, 60), std = Number(v.std), tank = Number(v.tank);
  if (!name) throw _err('กรอกชื่อหรือรุ่นรถ');
  if (!(std > 0 && std <= 500)) throw _err('อัตราสิ้นเปลืองมาตรฐานต้องมากกว่า 0');
  if (!(tank > 0 && tank <= 2000)) throw _err('ความจุถังไม่ถูกต้อง');
  var rec = {
    name: name, plate: _txt(v.plate, 20), type: _txt(v.type, 20) || 'อื่นๆ', fuel_type: _txt(v.fuel, 40),
    std_kmpl: std, tank_l: tank, odo_start: Math.max(0, Number(v.odo0) || 0), is_main: !!v.main
  };
  return _withLock(function () {
    var id = v.id;
    if (id) {
      var target = _rows('Vehicles').filter(function (r) { return r.vehicle_id === id && _bool(r.active); })[0];
      if (!target) throw _err('ไม่พบยานพาหนะนี้');
      _update('Vehicles', target._row, rec);
    } else {
      id = _newId('v', 'Vehicles', 'vehicle_id');
      rec.vehicle_id = id;
      rec.active = true;
      _append('Vehicles', rec);
    }
    var rows = _rows('Vehicles').filter(function (r) { return _bool(r.active); });
    if (rec.is_main) {
      rows.forEach(function (r) {
        if (r.vehicle_id !== id && _bool(r.is_main)) _update('Vehicles', r._row, { is_main: false });
      });
    } else if (!rows.some(function (r) { return _bool(r.is_main); })) {
      _update('Vehicles', rows[0]._row, { is_main: true });
    }
    return { vehicles: _vehicles() };
  });
}

/** ไม่ลบแถวทิ้ง แค่ตั้ง active = FALSE เพื่อเก็บประวัติไว้ในชีต */
function _deleteVehicle(p) {
  return _withLock(function () {
    var rows = _rows('Vehicles').filter(function (r) { return _bool(r.active); });
    if (rows.length <= 1) throw _err('ต้องมียานพาหนะอย่างน้อย 1 คัน');
    var target = rows.filter(function (r) { return r.vehicle_id === p.id; })[0];
    if (!target) throw _err('ไม่พบยานพาหนะนี้');
    _update('Vehicles', target._row, { active: false, is_main: false });
    var left = _rows('Vehicles').filter(function (r) { return _bool(r.active); });
    if (!left.some(function (r) { return _bool(r.is_main); })) _update('Vehicles', left[0]._row, { is_main: true });
    return { vehicles: _vehicles() };
  });
}

/* ================= ผู้ใช้ (หัวหน้างาน) ================= */

function _usersPayload() {
  return { users: _users(), drivers: _drivers() };
}

function _supCount() {
  return _rows('Users').filter(function (u) { return u.role === 'sup' && _bool(u.active); }).length;
}

function _findUser(id) {
  var u = _rows('Users').filter(function (x) { return x.user_id === id; })[0];
  if (!u) throw _err('ไม่พบผู้ใช้นี้');
  return u;
}

function _addUser(p) {
  var u = p.user || {};
  var name = _txt(u.name, 40), role = u.role === 'sup' ? 'sup' : 'staff', pin = String(u.pin || '');
  if (!name) throw _err('กรอกชื่อผู้ใช้');
  if (!/^\d{6}$/.test(pin)) throw _err('รหัสต้องเป็นตัวเลข 6 หลัก');
  if (_weak(pin)) throw _err('รหัสเดาง่ายเกินไป เลือกรหัสอื่น');
  return _withLock(function () {
    if (_rows('Users').some(function (x) { return String(x.name).toLowerCase() === name.toLowerCase(); })) {
      throw _err('มีชื่อนี้อยู่แล้ว');
    }
    if (_pinTaken(pin, null)) throw _err('รหัสนี้ถูกใช้แล้ว เลือกรหัสอื่น');
    _createUser(name, role, pin);
    return _usersPayload();
  });
}

function _setUserActive(me, p) {
  return _withLock(function () {
    var u = _findUser(p.id), active = !!p.active;
    if (!active) {
      if (u.user_id === me.user_id) throw _err('ปิดการใช้งานบัญชีของตัวเองไม่ได้');
      if (u.role === 'sup' && _bool(u.active) && _supCount() <= 1) throw _err('ต้องมีหัวหน้างานที่ใช้งานอยู่อย่างน้อย 1 คน');
    }
    _update('Users', u._row, { active: active });
    return _usersPayload();
  });
}

function _resetUserPin(me, p) {
  return _withLock(function () {
    var u = _findUser(p.id);
    var pin = _newPin(), salt = Utilities.getUuid(), now = Date.now();
    _update('Users', u._row, { pin_hash: _hash(pin, salt), salt: salt, pin_changed_at: now });
    var out = _usersPayload();
    out.pin = pin;
    // เปลี่ยนรหัสตัวเอง: โทเค็นเดิมใช้ไม่ได้แล้ว ออกโทเค็นใหม่ให้เลย
    if (u.user_id === me.user_id) out.token = _newToken(u.user_id);
    return out;
  });
}

function _deleteUser(me, p) {
  return _withLock(function () {
    var u = _findUser(p.id);
    if (u.user_id === me.user_id) throw _err('ลบบัญชีของตัวเองไม่ได้');
    if (u.role === 'sup' && _bool(u.active) && _supCount() <= 1) throw _err('ต้องมีหัวหน้างานที่ใช้งานอยู่อย่างน้อย 1 คน');
    _delete('Users', u._row);
    return _usersPayload();
  });
}

/* ================= รหัสและแฮช ================= */

function _hash(pin, salt) {
  var pepper = PropertiesService.getScriptProperties().getProperty('PIN_PEPPER') || '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pepper + ':' + salt + ':' + pin, Utilities.Charset.UTF_8);
  var hex = bytes.map(function (b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length < 2 ? '0' + h : h; }).join('');
  return 'h' + hex; // นำหน้าด้วยตัวอักษร กันชีตแปลงเป็นตัวเลข
}

function _weak(p) {
  return /^(\d)\1{5}$/.test(p) || ['123456', '654321', '012345', '234567', '345678', '456789', '567890', '098765'].indexOf(p) >= 0;
}

function _pinTaken(pin, exceptId) {
  return _rows('Users').some(function (u) {
    return u.user_id !== exceptId && u.salt && _hash(pin, String(u.salt)) === String(u.pin_hash);
  });
}

function _newPin() {
  var p;
  do { p = String(Math.floor(100000 + Math.random() * 900000)); } while (_weak(p) || _pinTaken(p, null));
  return p;
}

function _createUser(name, role, pin) {
  var salt = Utilities.getUuid();
  _append('Users', {
    user_id: _newId('u', 'Users', 'user_id'), name: name, role: role, pin_hash: _hash(pin, salt), salt: salt,
    active: true, pin_changed_at: Date.now(), created_at: new Date().toISOString()
  });
}

/* ================= ตัวช่วยเกี่ยวกับชีต ================= */

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _sheet(name) {
  var sh = _ss().getSheetByName(name);
  if (!sh) throw _err('ยังไม่ได้ตั้งค่าระบบ ให้รันฟังก์ชัน setup() ก่อน', 'setup');
  return sh;
}

function _rows(name) {
  var head = SCHEMA[name], sh = _sheet(name), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, head.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
}

function _append(name, obj) {
  _sheet(name).appendRow(SCHEMA[name].map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

function _update(name, rowNum, obj) {
  var sh = _sheet(name);
  SCHEMA[name].forEach(function (h, j) {
    if (Object.prototype.hasOwnProperty.call(obj, h)) sh.getRange(rowNum, j + 1).setValue(obj[h]);
  });
}

function _delete(name, rowNum) { _sheet(name).deleteRow(rowNum); }

function _withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function _bool(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

function _dstr(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).substring(0, 10);
}

/** ตัดความยาว และกันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ชีตตีความเป็นสูตร */
function _txt(s, max) {
  s = String(s === undefined || s === null ? '' : s).trim().substring(0, max);
  return /^[=+\-@]/.test(s) ? ' ' + s : s;
}

/** รหัสประจำตัวที่ไม่ซ้ำกับแถวเดิม (กันสองรายการที่สร้างในเสี้ยววินาทีเดียวกัน) */
function _newId(prefix, sheet, col) {
  var used = {};
  _rows(sheet).forEach(function (r) { used[String(r[col])] = true; });
  var id;
  do { id = prefix + Date.now().toString(36) + Math.floor(Math.random() * 46656).toString(36); } while (used[id]);
  return id;
}

function _err(msg, code) {
  var e = new Error(msg);
  e.code = code || 'invalid';
  return e;
}
