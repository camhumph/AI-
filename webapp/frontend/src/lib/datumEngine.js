/* ═══════════════════════════════════════════════════════════════════════════
   DATUM ENGINE  ·  STL → recognised features → machining plan → cost
   ---------------------------------------------------------------------------
   Framework-free. No React, no three.js, no DOM. Runs in a browser, in a
   web worker, or in Node. Drop this straight into the SolidWorks web app.

   Typical use:

     import * as Datum from "./datum-engine.js";

     const pos = Datum.parseSTL(arrayBuffer);          // Float32Array, 9 per tri
     const an  = Datum.analyze(pos);                   // segment + fit surfaces
     const fx  = Datum.extract(an);                    // holes, pockets, chamfers
     const st  = Datum.makeStock(fx.ext, "block", 3);  // or "plate" / "custom"
     const pl  = Datum.plan(an, fx, mat, tol, rate, { st, tap: true });
     const cs  = Datum.cost(pl, an, fx, mat, tol, rate);

     pl.cycle       total cycle minutes for one part
     pl.setupMin    setup minutes, off cycle
     pl.ops[]       every operation: tool, rpm, feed, pass counts, minutes
     pl.schedule[]  per-hole-group op sequence with minutes
     pl.facing      face-by-face pass counts off the stock
     cs.hours       { cycle, setup, total }
     cs.total       one-part cost

   If a SolidWorks macro can supply exact data (material, mass properties,
   Hole Wizard diameters), pass it through applyManifest() to override the
   values inferred from the mesh. See HANDOFF.md.
   ═══════════════════════════════════════════════════════════════════════════ */

/* ---------------- shop data ---------------- */
const MATERIALS = [
  { id: "6061",   name: "6061-T6",      cls: "Aluminium",  rho: 2.70, kg: 6.40,  vc: 400, fz: 0.14, k: 1.00, tc: 4 },
  { id: "7075",   name: "7075-T651",    cls: "Aluminium",  rho: 2.81, kg: 11.80, vc: 340, fz: 0.13, k: 1.18, tc: 6 },
  { id: "brass",  name: "Brass C360",   cls: "Copper alloy", rho: 8.50, kg: 13.20, vc: 450, fz: 0.12, k: 0.80, tc: 5 },
  { id: "1018",   name: "1018",         cls: "Mild steel", rho: 7.87, kg: 3.10,  vc: 180, fz: 0.12, k: 2.10, tc: 16 },
  { id: "4140",   name: "4140 PH",      cls: "Alloy steel", rho: 7.85, kg: 4.80,  vc: 140, fz: 0.10, k: 2.60, tc: 24 },
  { id: "304",    name: "304",          cls: "Stainless",  rho: 8.00, kg: 7.60,  vc: 120, fz: 0.10, k: 3.20, tc: 42 },
  { id: "174",    name: "17-4 PH H900", cls: "Stainless",  rho: 7.80, kg: 15.40, vc: 95,  fz: 0.09, k: 3.80, tc: 58 },
  { id: "ti64",   name: "Ti-6Al-4V",    cls: "Titanium",   rho: 4.43, kg: 44.00, vc: 60,  fz: 0.10, k: 5.20, tc: 110 },
  { id: "in718",  name: "Inconel 718",  cls: "Nickel",     rho: 8.19, kg: 72.00, vc: 30,  fz: 0.07, k: 8.60, tc: 320 },
  { id: "pom",    name: "Delrin POM",   cls: "Acetal",     rho: 1.41, kg: 9.20,  vc: 600, fz: 0.18, k: 0.45, tc: 2 },
  { id: "peek",   name: "PEEK",         cls: "Polymer",    rho: 1.30, kg: 210.0, vc: 400, fz: 0.15, k: 0.72, tc: 4 },
];

const TOLS = [
  { v: 0.25,   label: "±0.25",   name: "As milled",      fin: 1.00, passes: 1, insp: 6,  scrap: 0.005, gauge: "Calipers, sample check" },
  { v: 0.10,   label: "±0.10",   name: "General",        fin: 1.22, passes: 1, insp: 14, scrap: 0.010, gauge: "Height gauge on every piece" },
  { v: 0.05,   label: "±0.05",   name: "Close",          fin: 1.55, passes: 2, insp: 32, scrap: 0.020, gauge: "CMM first article, bore gauge in process" },
  { v: 0.025,  label: "±0.025",  name: "Precision",      fin: 2.10, passes: 2, insp: 55, scrap: 0.040, gauge: "CMM, temperature-controlled room" },
  { v: 0.0125, label: "±0.0125", name: "High precision", fin: 3.00, passes: 3, insp: 92, scrap: 0.080, gauge: "CMM, air gauge, 20 °C soak before cutting" },
];

const RPM_MAX = 12000, DUTY = 0.35, ALLOW = 3.0;

/* ---------------- STL parsing ---------------- */
function parseSTL(buf) {
  const dv = new DataView(buf);
  let binary = buf.byteLength >= 84;
  if (binary) {
    const n = dv.getUint32(80, true);
    binary = 84 + n * 50 === buf.byteLength;
  }
  if (!binary) {
    const txt = new TextDecoder().decode(buf);
    if (!/facet\s+normal/i.test(txt)) throw new Error("Not an STL file.");
    const out = [];
    const re = /vertex\s+(-?[\d.eE+-]+)\s+(-?[\d.eE+-]+)\s+(-?[\d.eE+-]+)/g;
    let m;
    while ((m = re.exec(txt))) out.push(+m[1], +m[2], +m[3]);
    if (out.length < 9) throw new Error("No triangles found.");
    return new Float32Array(out);
  }
  const n = dv.getUint32(80, true);
  const pos = new Float32Array(n * 9);
  let o = 84, p = 0;
  for (let i = 0; i < n; i++) {
    o += 12;
    for (let v = 0; v < 9; v++) { pos[p++] = dv.getFloat32(o, true); o += 4; }
    o += 2;
  }
  return pos;
}

/* ---------------- geometry analysis ---------------- */
const COS_COARSE = Math.cos((20 * Math.PI) / 180);
const COS_FINE = Math.cos((5 * Math.PI) / 180);

function analyze(pos) {
  const nT = pos.length / 9;
  /* weld vertices */
  const Q = 1e4, map = new Map();
  const vid = new Int32Array(nT * 3);
  const vx = [], vy = [], vz = [];
  for (let i = 0; i < nT * 3; i++) {
    const x = pos[i * 3], y = pos[i * 3 + 1], z = pos[i * 3 + 2];
    const key = `${Math.round(x * Q)},${Math.round(y * Q)},${Math.round(z * Q)}`;
    let id = map.get(key);
    if (id === undefined) { id = vx.length; map.set(key, id); vx.push(x); vy.push(y); vz.push(z); }
    vid[i] = id;
  }
  const nV = vx.length;

  /* face normals, areas, centroids, volume, bbox */
  const nrm = new Float32Array(nT * 3), area = new Float32Array(nT);
  const cen = new Float32Array(nT * 3);
  let vol = 0, totalArea = 0;
  const bb = [Infinity, Infinity, Infinity, -Infinity, -Infinity, -Infinity];
  for (let f = 0; f < nT; f++) {
    const i = f * 9;
    const ax = pos[i], ay = pos[i + 1], az = pos[i + 2];
    const bx = pos[i + 3], by = pos[i + 4], bz = pos[i + 5];
    const cx = pos[i + 6], cy = pos[i + 7], cz = pos[i + 8];
    const ux = bx - ax, uy = by - ay, uz = bz - az;
    const wx = cx - ax, wy = cy - ay, wz = cz - az;
    const px = uy * wz - uz * wy, py = uz * wx - ux * wz, pz = ux * wy - uy * wx;
    const L = Math.hypot(px, py, pz) || 1e-12;
    area[f] = L / 2; totalArea += L / 2;
    nrm[f * 3] = px / L; nrm[f * 3 + 1] = py / L; nrm[f * 3 + 2] = pz / L;
    cen[f * 3] = (ax + bx + cx) / 3; cen[f * 3 + 1] = (ay + by + cy) / 3; cen[f * 3 + 2] = (az + bz + cz) / 3;
    vol += (ax * (by * cz - bz * cy) + ay * (bz * cx - bx * cz) + az * (bx * cy - by * cx)) / 6;
    for (let k = 0; k < 3; k++) {
      const X = pos[i + k * 3], Y = pos[i + k * 3 + 1], Z = pos[i + k * 3 + 2];
      if (X < bb[0]) bb[0] = X; if (Y < bb[1]) bb[1] = Y; if (Z < bb[2]) bb[2] = Z;
      if (X > bb[3]) bb[3] = X; if (Y > bb[4]) bb[4] = Y; if (Z > bb[5]) bb[5] = Z;
    }
  }
  vol = Math.abs(vol);

  /* face adjacency over welded edges */
  const eMap = new Map(), adjH = new Int32Array(nT * 3).fill(-1);
  for (let f = 0; f < nT; f++) {
    for (let e = 0; e < 3; e++) {
      const a = vid[f * 3 + e], b = vid[f * 3 + ((e + 1) % 3)];
      const key = a < b ? a * nV + b : b * nV + a;
      const prev = eMap.get(key);
      if (prev === undefined) eMap.set(key, f * 3 + e);
      else { adjH[prev] = f * 3 + e; adjH[f * 3 + e] = prev; }
    }
  }
  const nbr = (h) => (adjH[h] < 0 ? -1 : (adjH[h] / 3) | 0);

  /* region growing, optionally restricted to a subset */
  const mark = new Int32Array(nT).fill(-1);
  let pass = 0;
  const stack = new Int32Array(nT);
  function grow(faces, cosT) {
    pass++;
    const inSet = new Uint8Array(nT);
    for (const f of faces) inSet[f] = 1;
    const seen = new Int32Array(0);
    const local = new Map();
    const outGroups = [];
    const done = new Uint8Array(nT);
    for (const s of faces) {
      if (done[s]) continue;
      let sp = 0; stack[sp++] = s; done[s] = 1;
      const cur = [s];
      while (sp > 0) {
        const f = stack[--sp];
        for (let e = 0; e < 3; e++) {
          const g = nbr(f * 3 + e);
          if (g < 0 || !inSet[g] || done[g]) continue;
          const d = nrm[f * 3] * nrm[g * 3] + nrm[f * 3 + 1] * nrm[g * 3 + 1] + nrm[f * 3 + 2] * nrm[g * 3 + 2];
          if (d > cosT) { done[g] = 1; stack[sp++] = g; cur.push(g); }
        }
      }
      outGroups.push(cur);
    }
    return outGroups;
  }

  const snap = (a) => {
    const A = [Math.abs(a[0]), Math.abs(a[1]), Math.abs(a[2])];
    const i = A[0] > A[1] ? (A[0] > A[2] ? 0 : 2) : A[1] > A[2] ? 1 : 2;
    if (A[i] > 0.9994) { const o = [0, 0, 0]; o[i] = a[i] > 0 ? 1 : -1; return o; }
    return a;
  };

  /* classify one face group as plane, cylinder, or free-form */
  function classify(fs, regOf) {
    let ar = 0, mx = 0, my = 0, mz = 0;
    for (const f of fs) { ar += area[f]; mx += nrm[f * 3] * area[f]; my += nrm[f * 3 + 1] * area[f]; mz += nrm[f * 3 + 2] * area[f]; }
    const mL = Math.hypot(mx, my, mz) || 1e-12;
    let mn = [mx / mL, my / mL, mz / mL];
    let minDot = 1;
    for (const f of fs) {
      const d = nrm[f * 3] * mn[0] + nrm[f * 3 + 1] * mn[1] + nrm[f * 3 + 2] * mn[2];
      if (d < minDot) minDot = d;
    }
    const spread = (Math.acos(Math.max(-1, Math.min(1, minDot))) * 180) / Math.PI;

    /* boundary length that has material rising above it */
    let bTot = 0, bCon = 0;
    for (const f of fs) for (let e = 0; e < 3; e++) {
      const g = nbr(f * 3 + e);
      if (g >= 0 && regOf(g) === regOf(f)) continue;
      const va = vid[f * 3 + e], vb = vid[f * 3 + ((e + 1) % 3)];
      const len = Math.hypot(vx[va] - vx[vb], vy[va] - vy[vb], vz[va] - vz[vb]);
      bTot += len;
      if (g < 0) continue;
      const dz = (cen[g * 3] - cen[f * 3]) * mn[0] + (cen[g * 3 + 1] - cen[f * 3 + 1]) * mn[1]
               + (cen[g * 3 + 2] - cen[f * 3 + 2]) * mn[2];
      if (dz > 1e-4) bCon += len;
    }
    const encl = bTot ? bCon / bTot : 0;

    if (spread < 2.5) {
      mn = snap(mn);
      let off = 0;
      for (const f of fs) off += (cen[f * 3] * mn[0] + cen[f * 3 + 1] * mn[1] + cen[f * 3 + 2] * mn[2]) * area[f];
      off /= ar;
      const ang = (Math.acos(Math.min(1, Math.max(Math.abs(mn[0]), Math.abs(mn[1]), Math.abs(mn[2])))) * 180) / Math.PI;
      const t = Math.abs(mn[0]) < 0.9 ? [1, 0, 0] : [0, 1, 0];
      const u = unit(cross(mn, t)), w = cross(mn, u);
      let u0 = Infinity, u1 = -Infinity, w0 = Infinity, w1 = -Infinity;
      for (const f of fs) for (let k = 0; k < 3; k++) {
        const i = f * 9 + k * 3, P = [pos[i], pos[i + 1], pos[i + 2]];
        const a = dot(P, u), b = dot(P, w);
        if (a < u0) u0 = a; if (a > u1) u1 = a; if (b < w0) w0 = b; if (b > w1) w1 = b;
      }
      const e2 = [u1 - u0, w1 - w0].sort((a, b) => a - b);
      return { kind: "plane", faces: fs, area: ar, n: mn, off, ang, narrow: e2[0], long: e2[1], encl, spread };
    }

    /* cylinder: axis is perpendicular to every normal */
    const n0 = [nrm[fs[0] * 3], nrm[fs[0] * 3 + 1], nrm[fs[0] * 3 + 2]];
    let best = 1, bi = -1;
    for (const f of fs) {
      const d = Math.abs(n0[0] * nrm[f * 3] + n0[1] * nrm[f * 3 + 1] + n0[2] * nrm[f * 3 + 2]);
      if (d < best) { best = d; bi = f; }
    }
    if (bi < 0 || best > 0.995) return { kind: "free", faces: fs, area: ar, spread, encl };
    const ax = snap(unit(cross(n0, [nrm[bi * 3], nrm[bi * 3 + 1], nrm[bi * 3 + 2]])));
    for (const f of fs) {
      if (Math.abs(nrm[f * 3] * ax[0] + nrm[f * 3 + 1] * ax[1] + nrm[f * 3 + 2] * ax[2]) > 0.14)
        return { kind: "free", faces: fs, area: ar, spread, encl };
    }
    const t = Math.abs(ax[0]) < 0.9 ? [1, 0, 0] : [0, 1, 0];
    const u = unit(cross(ax, t)), w = cross(ax, u);
    let Sx = 0, Sy = 0, Sxx = 0, Syy = 0, Sxy = 0, Sxz = 0, Syz = 0, Sz = 0, n = 0;
    let a0 = Infinity, a1 = -Infinity;
    for (const f of fs) for (let k = 0; k < 3; k++) {
      const i = f * 9 + k * 3, P = [pos[i], pos[i + 1], pos[i + 2]];
      const x = dot(P, u), y = dot(P, w), h = dot(P, ax);
      if (h < a0) a0 = h; if (h > a1) a1 = h;
      const z = x * x + y * y;
      Sx += x; Sy += y; Sxx += x * x; Syy += y * y; Sxy += x * y; Sxz += x * z; Syz += y * z; Sz += z; n++;
    }
    const sol = solve3([[2 * Sxx, 2 * Sxy, Sx], [2 * Sxy, 2 * Syy, Sy], [2 * Sx, 2 * Sy, n]], [Sxz, Syz, Sz]);
    if (!sol) return { kind: "free", faces: fs, area: ar, spread, encl };
    const [cu, cw, cc] = sol;
    const rad = Math.sqrt(Math.max(cc + cu * cu + cw * cw, 1e-9));
    let res = 0, cnt = 0, vote = 0;
    const angs = [];
    for (const f of fs) {
      const gx = cen[f * 3], gy = cen[f * 3 + 1], gz = cen[f * 3 + 2];
      const x = gx * u[0] + gy * u[1] + gz * u[2] - cu;
      const y = gx * w[0] + gy * w[1] + gz * w[2] - cw;
      const Lr = Math.hypot(x, y) || 1e-9;
      res += Math.abs(Lr - rad); cnt++;
      const nu = nrm[f * 3] * u[0] + nrm[f * 3 + 1] * u[1] + nrm[f * 3 + 2] * u[2];
      const nw = nrm[f * 3] * w[0] + nrm[f * 3 + 1] * w[1] + nrm[f * 3 + 2] * w[2];
      vote += (nu * x + nw * y) / Lr;
      angs.push(Math.atan2(y, x));
    }
    res /= cnt;
    if (res > Math.max(0.09 * rad, 0.3)) return { kind: "free", faces: fs, area: ar, spread, encl };
    angs.sort((a, b) => a - b);
    let gap = angs[0] + Math.PI * 2 - angs[angs.length - 1];
    for (let i = 1; i < angs.length; i++) gap = Math.max(gap, angs[i] - angs[i - 1]);
    return {
      kind: "cyl", faces: fs, area: ar, axis: ax, rad,
      extent: ((Math.PI * 2 - gap) * 180) / Math.PI,
      concave: vote < 0, len: a1 - a0, a0, a1, cu, cw, u, w, spread,
      center: [cu * u[0] + cw * w[0], cu * u[1] + cw * w[1], cu * u[2] + cw * w[2]],
    };
  }

  /* pass 1 at 20°, then re-segment any leaked region at 5° */
  const all = new Int32Array(nT);
  for (let i = 0; i < nT; i++) all[i] = i;
  const coarse = grow(all, COS_COARSE);
  const finalGroups = [];
  const tmp = new Int32Array(nT);
  for (const g of coarse) { for (const f of g) tmp[f] = finalGroups.length; finalGroups.push(g); }
  const probe = (g) => classify(g, (f) => tmp[f]);

  const refined = [];
  for (const g of coarse) {
    const c = probe(g);
    if (c.kind === "free" && c.spread > 45 && g.length > 6) {
      for (const sub of grow(g, COS_FINE)) refined.push(sub);
    } else refined.push(g);
  }

  const reg = new Int32Array(nT).fill(-1);
  refined.forEach((g, i) => { for (const f of g) reg[f] = i; });
  const regions = refined.map((g, i) => {
    const c = classify(g, (f) => reg[f]);
    c.r = i;
    return c;
  });

  return { nT, nV, vol, area: totalArea, bb, reg, regions, pos, nrm, cen };
}

const cross = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]];
const dot = (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
const unit = (a) => { const L = Math.hypot(a[0], a[1], a[2]) || 1e-12; return [a[0] / L, a[1] / L, a[2] / L]; };
function solve3(M, b) {
  const d =
    M[0][0] * (M[1][1] * M[2][2] - M[1][2] * M[2][1]) -
    M[0][1] * (M[1][0] * M[2][2] - M[1][2] * M[2][0]) +
    M[0][2] * (M[1][0] * M[2][1] - M[1][1] * M[2][0]);
  if (Math.abs(d) < 1e-12) return null;
  const col = (j, v) => M.map((row, i) => row.map((x, k) => (k === j ? v[i] : x)));
  const det = (A) =>
    A[0][0] * (A[1][1] * A[2][2] - A[1][2] * A[2][1]) -
    A[0][1] * (A[1][0] * A[2][2] - A[1][2] * A[2][0]) +
    A[0][2] * (A[1][0] * A[2][1] - A[1][1] * A[2][0]);
  return [det(col(0, b)) / d, det(col(1, b)) / d, det(col(2, b)) / d];
}

/* ---------------- ray cast: is a point inside the solid ---------------- */
const RAY = unit([0.4472, 0.5477, 0.7071]);
function insideSolid(pos, nT, p) {
  const [dx, dy, dz] = RAY;
  let hits = 0;
  for (let f = 0; f < nT; f++) {
    const i = f * 9;
    const ax = pos[i], ay = pos[i + 1], az = pos[i + 2];
    const e1x = pos[i + 3] - ax, e1y = pos[i + 4] - ay, e1z = pos[i + 5] - az;
    const e2x = pos[i + 6] - ax, e2y = pos[i + 7] - ay, e2z = pos[i + 8] - az;
    const px = dy * e2z - dz * e2y, py = dz * e2x - dx * e2z, pz = dx * e2y - dy * e2x;
    const det = e1x * px + e1y * py + e1z * pz;
    if (det > -1e-12 && det < 1e-12) continue;
    const inv = 1 / det;
    const tx = p[0] - ax, ty = p[1] - ay, tz = p[2] - az;
    const u = (tx * px + ty * py + tz * pz) * inv;
    if (u < 0 || u > 1) continue;
    const qx = ty * e1z - tz * e1y, qy = tz * e1x - tx * e1z, qz = tx * e1y - ty * e1x;
    const v = (dx * qx + dy * qy + dz * qz) * inv;
    if (v < 0 || u + v > 1) continue;
    if ((e2x * qx + e2y * qy + e2z * qz) * inv > 1e-7) hits++;
  }
  return (hits & 1) === 1;
}

/* tap drill sizes — metric coarse and unified coarse */
const TAPS = [
  [1.60, "M2", 0.40], [2.05, "M2.5", 0.45], [2.50, "M3", 0.50], [3.30, "M4", 0.70],
  [4.20, "M5", 0.80], [5.00, "M6", 1.00], [6.80, "M8", 1.25], [8.50, "M10", 1.50],
  [10.20, "M12", 1.75], [12.00, "M14", 2.00], [14.00, "M16", 2.00], [17.50, "M20", 2.50],
  [21.00, "M24", 3.00], [2.70, "#6-32", 0.794], [3.50, "#8-32", 0.794],
  [3.90, "#10-24", 1.058], [5.10, "1/4-20", 1.270], [6.60, "5/16-18", 1.411],
  [8.00, "3/8-16", 1.588], [10.80, "1/2-13", 1.954],
];
const tapFor = (d) => {
  for (const [td, name, pitch] of TAPS) if (Math.abs(d - td) < 0.12) return { name, pitch };
  return null;
};

/* peck strategy straight off the depth-to-diameter ratio */
function peckPlan(ld) {
  if (ld <= 3) return { mode: "straight", fac: 1.00, feedScale: 1.00 };
  if (ld <= 5) return { mode: "chip break", fac: 1.25, feedScale: 0.95 };
  if (ld <= 8) return { mode: "full retract", fac: 1.70, feedScale: 0.85 };
  return { mode: "deep hole", fac: 2.40, feedScale: 0.70 };
}

/* ---------------- feature extraction ---------------- */
function extract(an) {
  const [x0, y0, z0, x1, y1, z1] = an.bb;
  const ext = [x1 - x0, y1 - y0, z1 - z0];
  const cyls = an.regions.filter((g) => g.kind === "cyl");
  const planes = an.regions.filter((g) => g.kind === "plane");
  const frees = an.regions.filter((g) => g.kind === "free");

  const akey = (a) => {
    let v = a.slice();
    const i = [0, 1, 2].reduce((m, k) => (Math.abs(v[k]) > Math.abs(v[m]) ? k : m), 0);
    if (v[i] < 0) v = v.map((q) => -q);
    return v.map((q) => Math.round(q * 50) / 50).join(",");
  };

  /* group coaxial, equal-radius cylinders */
  const gm = new Map();
  for (const c of cyls) {
    const key = `${akey(c.axis)}|${c.rad.toFixed(2)}|${Math.round(c.cu * 4)},${Math.round(c.cw * 4)}`;
    if (!gm.has(key)) gm.set(key, []);
    gm.get(key).push(c);
  }

  const holes = [], bosses = [], fillets = [];
  for (const g of gm.values()) {
    const extent = Math.min(360, g.reduce((t, c) => t + c.extent, 0));
    const len = Math.max(...g.map((c) => c.len));
    const c = g[0];
    const item = { d: c.rad * 2, depth: len, axis: c.axis, akey: akey(c.axis), regs: g.map((q) => q.r),
                   center: c.center, area: g.reduce((s, q) => s + q.area, 0), concave: c.concave,
                   a0: Math.min(...g.map((q) => q.a0)), a1: Math.max(...g.map((q) => q.a1)) };
    if (extent > 300 && c.concave) holes.push(item);
    else if (extent > 300) bosses.push(item);
    else if (c.rad < 30) fillets.push({ ...item, extent });
  }

  /* drop arc slivers that sit on an already-identified cylinder */
  const realCyl = [...holes, ...bosses];
  const fillets2 = fillets.filter((f) => !realCyl.some((h) =>
    h.akey === f.akey && Math.abs(h.d - f.d) / h.d < 0.08 &&
    Math.hypot(h.center[0] - f.center[0], h.center[1] - f.center[1], h.center[2] - f.center[2]) < 1.2));
  fillets.length = 0; fillets.push(...fillets2);

  /* through or blind: probe just beyond each end of the hole axis */
  for (const h of holes) {
    const at = (t) => [h.center[0] + h.axis[0] * t, h.center[1] + h.axis[1] * t, h.center[2] + h.axis[2] * t];
    const capLo = insideSolid(an.pos, an.nT, at(h.a0 - 0.6));
    const capHi = insideSolid(an.pos, an.nT, at(h.a1 + 0.6));
    h.blind = capLo !== capHi;
    h.through = !capLo && !capHi;
    h.entry = capLo ? "hi" : "lo";
    h.ld = h.depth / h.d;
    h.peck = peckPlan(h.ld);
    h.tap = h.ld <= 6 ? tapFor(h.d) : null;
  }

  /* pocket floors: a planar face fully enclosed by a larger parallel face above it */
  const ups = planes.filter((p) => p.ang < 3 && p.area > 20);
  const bbOf = (p) => {
    const t = Math.abs(p.n[0]) < 0.9 ? [1, 0, 0] : [0, 1, 0];
    const u = unit(cross(p.n, t)), w = cross(p.n, u);
    let a0 = Infinity, a1 = -Infinity, b0 = Infinity, b1 = -Infinity;
    for (const f of p.faces) for (let k = 0; k < 3; k++) {
      const i = f * 9 + k * 3, P = [an.pos[i], an.pos[i + 1], an.pos[i + 2]];
      const a = dot(P, u), b = dot(P, w);
      if (a < a0) a0 = a; if (a > a1) a1 = a; if (b < b0) b0 = b; if (b > b1) b1 = b;
    }
    return { u, w, a0, a1, b0, b1 };
  };
  ups.forEach((p) => (p._bb = bbOf(p)));
  const axisExt = (n) => Math.abs(n[0]) > 0.9 ? ext[0] : Math.abs(n[1]) > 0.9 ? ext[1] : ext[2];

  const pockets = [], faces = [];
  for (const p of ups) {
    const B = p._bb, lim = axisExt(p.n) * 0.55;
    const parents = ups.filter((o) => {
      if (o === p || dot(o.n, p.n) < 0.999) return false;
      const d = o.off - p.off;
      if (d < 0.05 || d > lim) return false;
      if (o.area < p.area * 1.05) return false;
      const C = o._bb;
      const oa = Math.min(C.a1, B.a1) - Math.max(C.a0, B.a0);
      const ob = Math.min(C.b1, B.b1) - Math.max(C.b0, B.b0);
      if (oa <= 0 || ob <= 0) return false;
      const own = Math.max((B.a1 - B.a0) * (B.b1 - B.b0), 1e-6);
      return (oa * ob) / own > 0.85;
    });
    if (!parents.length || p.encl < 0.45) { faces.push({ reg: p.r, area: p.area, n: p.n, regs: [p.r] }); continue; }
    const parent = parents.reduce((m, o) => (o.off < m.off ? o : m));
    const depth = parent.off - p.off;
    const span = Math.min(B.a1 - B.a0, B.b1 - B.b0);
    if (depth > span * 4.5 + 3) { faces.push({ reg: p.r, area: p.area, n: p.n, regs: [p.r] }); continue; }
    const ca = (B.a0 + B.a1) / 2, cb = (B.b0 + B.b1) / 2;
    const coax = (h) => Math.abs(dot(h.axis, p.n)) > 0.97 &&
      Math.hypot(dot(h.center, B.u) - ca, dot(h.center, B.w) - cb) < Math.max(2, h.d * 0.6);
    /* the flat bottom of a blind hole belongs to the drilling op, not to milling */
    if (holes.some((h) => h.blind && coax(h) && Math.abs(h.d - span) < h.d * 0.3)) continue;
    /* counterbore: a smaller hole passes through a shallow seat */
    const seated = holes.some((h) => coax(h) && h.d < span * 0.85);
    const cbore = seated && depth < 16 && p.area < 1400;
    p._bbRef = { ca, cb };
    if (p.encl < 0.50 && !cbore) { faces.push({ reg: p.r, area: p.area, n: p.n, regs: [p.r] }); continue; }
    pockets.push({ reg: p.r, depth, area: p.area, n: p.n, cbore, regs: [p.r], bb: { ca, cb } });
  }

  /* a hole under a counterbore is a fastener clearance hole, not a fit */
  const cb = pockets.filter((p) => p.cbore);
  for (const h of holes) {
    h.counterbored = cb.some((p) => {
      if (Math.abs(dot(h.axis, p.n)) < 0.97) return false;
      const t = Math.abs(p.n[0]) < 0.9 ? [1, 0, 0] : [0, 1, 0];
      const u = unit(cross(p.n, t)), w = cross(p.n, u);
      const pb = p.bb;
      return pb && Math.hypot(dot(h.center, u) - pb.ca, dot(h.center, w) - pb.cb) < Math.max(2.5, h.d);
    });
    h.purpose = h.d > 20 ? "bore" : h.tap ? "tapped" : h.counterbored ? "clearance" : "fit";
  }

  /* smallest concave (internal) radius — sets the smallest usable tool */
  const internalR = fillets.filter((f) => f.concave).map((f) => f.d / 2);
  const minInternalR = internalR.length ? Math.min(...internalR) : null;

  /* chamfers: narrow planar bands off-axis */
  const angled = frees.filter((f) => (f.spread ?? 99) < 42);
  const curved = frees.filter((f) => (f.spread ?? 99) >= 42);
  const chamfers = planes.filter((p) => p.ang > 20 && p.ang < 70 && p.narrow < Math.max(9, p.long * 0.12) && p.long > 4);
  const chamferLen = chamfers.reduce((s, p) => s + p.long, 0);

  /* setups from machined directions */
  const dirs = new Map();
  const bump = (k, w) => dirs.set(k, (dirs.get(k) || 0) + w);
  holes.forEach((h) => bump(h.akey, 3));
  bosses.forEach((b) => bump(b.akey, 2));
  pockets.forEach((p) => bump(akey(p.n), 3));
  faces.filter((f) => f.area > ext[0] * ext[1] * 0.04).forEach((f) => bump(akey(f.n), 1));
  const setups = Math.max(1, Math.min(6, [...dirs.values()].filter((v) => v >= 2).length));

  const stock = [ext[0] + 2 * ALLOW, ext[1] + 2 * ALLOW, ext[2] + 2 * ALLOW];
  const stockVol = stock[0] * stock[1] * stock[2];

  return { ext, stock, stockVol, holes, bosses, fillets, pockets, faces, chamfers, chamferLen,
          setups, frees, angled, curved, dirs, minInternalR, diag: Math.hypot(...ext) };
}

/* ---------------- stock and facing ---------------- */
const PLATE_MM = [3, 4, 5, 6, 8, 10, 12, 15, 16, 20, 25, 30, 32, 40, 50, 60, 70, 80, 100, 125, 150];
const PLATE_IN = [3.175, 4.7625, 6.35, 9.525, 12.7, 15.875, 19.05, 25.4, 31.75, 38.1, 50.8, 63.5, 76.2, 101.6];
const STOCK_MODES = [
  { id: "block", name: "Saw-cut block", note: "six faces to square up" },
  { id: "plate", name: "Plate", note: "thickness snaps to a size you can buy" },
  { id: "custom", name: "Have it already", note: "enter the stock in your rack" },
];

function makeStock(ext, mode, allow, custom, imperial) {
  if (mode === "custom" && custom) {
    const s = custom.map((v, i) => Math.max(v, ext[i]));
    return { stock: s, faces: [0, 1, 2], mode, short: custom.some((v, i) => v < ext[i] - 0.01) };
  }
  if (mode === "plate") {
    const table = imperial ? PLATE_IN : PLATE_MM;
    const need = ext[2] + 3;
    const t = table.find((v) => v >= need) ?? Math.ceil(need);
    return { stock: [ext[0] + 2 * allow, ext[1] + 2 * allow, t], faces: [2], mode, plateT: t, short: false };
  }
  return { stock: ext.map((v) => v + 2 * allow), faces: [0, 1, 2], mode, short: false };
}

/* facing: real pass counts off the stock you actually have */
function facingOps(ext, st, mat, rpm, kfac) {
  // 2" face mill = 50.8 mm. The shop finishes every face with the 2" cutter,
  // so that is the width the pass count has to be built on. This was 63 mm
  // (2.48"), which is 24% wider and therefore ~24% fewer width passes than the
  // machine actually makes. See FINISH_FACE_MILL_IN in lib/shopTools.ts.
  const Dface = 50.8, ae = Dface * 0.75;
  const nF = rpm(Dface), vf = nF * mat.fz * 5;
  const ap = mat.k <= 1.2 ? 2.0 : mat.k <= 2.5 ? 1.5 : 1.0;
  const AX = ["X", "Y", "Z"];
  const out = [];
  let total = 0, removed = 0;
  for (const a of st.faces) {
    const excess = st.stock[a] - ext[a];
    if (excess < 0.2) continue;
    const nFaces = st.mode === "plate" ? 2 : 2;
    const perFace = excess / nFaces;
    const others = [0, 1, 2].filter((i) => i !== a);
    const L = st.stock[others[0]], W = st.stock[others[1]];
    const wPasses = Math.ceil(W / ae);
    const dPasses = Math.ceil(perFace / ap);
    const min = nFaces * (dPasses * wPasses * ((L + Dface) / vf) + 0.45) * kfac;
    total += min;
    removed += excess * L * W;
    out.push({ axis: AX[a], nFaces, perFace, L, W, wPasses, dPasses, min });
  }
  return { list: out, total, removed, Dface, rpmFace: nF, vfFace: vf, ap };
}

/* ---------------- machining plan ---------------- */
function plan(an, fx, mat, tol, rate, opts = {}) {
  const tapOn = opts.tap !== false;
  const st = opts.st || { stock: fx.stock, faces: [0, 1, 2], mode: "block" };
  const stockVol = st.stock[0] * st.stock[1] * st.stock[2];
  const ops = [];
  const kfac = 1 / (1 + (mat.k - 1) * 0.32);
  const rpm = (D, f = 1) => Math.min(RPM_MAX, (mat.vc * f * 1000) / (Math.PI * D));
  /* slender tools deflect; feed has to come down with the length-to-diameter ratio */
  const reach = (ld) => 1 / (1 + Math.pow(Math.max(0, ld - 3) / 2.5, 1.5));

  /* ---- tool sizing, limited by the smallest internal corner ---- */
  const deepest = fx.pockets.length ? Math.max(...fx.pockets.map((p) => p.depth)) : 0;
  let Dr = Math.max(8, Math.min(20, Math.cbrt(fx.stockVol) * 0.12));
  let cornerLimited = false;
  if (fx.minInternalR !== null && fx.minInternalR * 2 < Dr) { Dr = Math.max(3, fx.minInternalR * 1.8); cornerLimited = true; }
  const Df = Math.max(4, Math.min(12, Dr * 0.65));
  const nR = rpm(Dr), vfR = nR * mat.fz * 4;
  const nF = rpm(Df), vfF = nF * mat.fz * 0.6 * 4;

  const holeVol = fx.holes.reduce((s, h) => s + Math.PI * (h.d / 2) ** 2 * h.depth, 0);
  const pocketVol = fx.pockets.reduce((s, p) => s + p.area * p.depth, 0);
  const removeVol = Math.max(0, stockVol - an.vol - holeVol);
  const openVol = Math.max(0, removeVol - pocketVol);
  const stockKgP = (stockVol / 1e6) * mat.rho;

  /* ---- 1. facing, driven by the stock you actually have ---- */
  const fc = facingOps(fx.ext, st, mat, rpm, kfac);
  for (const f of fc.list) {
    ops.push({
      id: `face${f.axis}`, label: `Face ${f.nFaces} ${f.axis} face${f.nFaces > 1 ? "s" : ""}`,
      tool: `Ø${fc.Dface} 5-insert face mill`,
      detail: `${f.perFace.toFixed(2)} mm per face · ${f.dPasses}×${f.wPasses} passes · ${fc.rpmFace.toFixed(0)} rpm · ${fc.vfFace.toFixed(0)} mm/min`,
      qty: `${f.L.toFixed(0)} × ${f.W.toFixed(0)} mm`,
      min: f.min,
    });
  }

  /* ---- 2. open roughing ---- */
  const apO = Dr * 0.5 * kfac, aeO = Dr * 0.45;
  const mrrO = apO * aeO * vfR;
  ops.push({
    id: "roughOpen", label: "Rough the open faces", tool: `Ø${Dr.toFixed(0)} 4-flute carbide${cornerLimited ? " (corner-limited)" : ""}`,
    detail: `${((mrrO * 0.42) / 1000).toFixed(0)} cm³/min sustained · ${nR.toFixed(0)} rpm · ap ${apO.toFixed(1)} ae ${aeO.toFixed(1)}`,
    qty: `${(openVol / 1000).toFixed(0)} cm³`,
    min: openVol / (mrrO * 0.42),
  });

  /* ---- 3. pocket roughing, derated for reach ---- */
  if (pocketVol > 500) {
    const ld = (deepest + 3) / Df, rf = reach(ld);
    const mrrP = Df * 0.45 * kfac * (Df * 0.4) * vfF * rf;
    ops.push({
      id: "roughPocket", label: `Rough ${fx.pockets.filter((p) => !p.cbore).length} pocket${fx.pockets.filter((p) => !p.cbore).length === 1 ? "" : "s"}`,
      tool: `Ø${Df.toFixed(0)} 4-flute, ${ld.toFixed(1)}×D reach`,
      detail: `${((mrrP * 0.24) / 1000).toFixed(1)} cm³/min · ramp entry · feed at ${(rf * 100).toFixed(0)}% for deflection`,
      qty: `${(pocketVol / 1000).toFixed(1)} cm³, deepest ${deepest.toFixed(1)} mm`,
      min: pocketVol / Math.max(mrrP * 0.24, 1),
      regs: fx.pockets.filter((p) => !p.cbore).flatMap((p) => p.regs),
    });
  }

  /* ---- 4. rest machining where the roughing tool could not reach the corners ---- */
  if (fx.minInternalR !== null && fx.minInternalR * 2 < Dr * 0.95) {
    const Dsm = Math.max(2, fx.minInternalR * 1.7);
    const restVol = fx.pockets.reduce((s, p) => s + 4 * (Dr / 2) ** 2 * (1 - Math.PI / 4) * p.depth, 0);
    const mrrS = Dsm * 0.3 * kfac * (Dsm * 0.35) * rpm(Dsm) * mat.fz * 0.5 * 4;
    ops.push({
      id: "rest", label: "Rest mill the corners", tool: `Ø${Dsm.toFixed(1)} carbide, R${fx.minInternalR.toFixed(1)} corners`,
      detail: `the Ø${Dr.toFixed(0)} rougher cannot reach an R${fx.minInternalR.toFixed(1)} corner`,
      qty: `${(restVol / 1000).toFixed(2)} cm³ left in corners`,
      min: restVol / Math.max(mrrS * 0.2, 0.5),
    });
  }

  /* ---- 5. semi-finish and finish ---- */
  const contourArea = fx.fillets.reduce((s, f) => s + f.area, 0) + fx.curved.reduce((s, f) => s + f.area, 0);
  const machArea = Math.max(an.area * 0.18, an.area * 0.60 - contourArea);
  const ldF = (deepest + 3) / Df, rfF = reach(ldF);
  if (tol.passes > 1) {
    ops.push({
      id: "semi", label: "Semi-finish walls and floors", tool: `Ø${Df.toFixed(0)} 4-flute, 0.3 mm left on`,
      detail: `${(Df * 0.3).toFixed(2)} mm stepover · ${nF.toFixed(0)} rpm`,
      qty: `${(machArea / 100).toFixed(0)} cm²`,
      min: (machArea / (Df * 0.3 * vfF * rfF)) * 0.55 * kfac,
    });
  }
  ops.push({
    id: "finish", label: "Finish walls and floors",
    tool: `Ø${Df.toFixed(0)} 4-flute, ${tol.passes > 2 ? "3 spring passes" : "1 spring pass"}`,
    detail: `${(Df * 0.15).toFixed(2)} mm stepover · ${vfF.toFixed(0)} mm/min · ${tol.label} band · reach ${(rfF * 100).toFixed(0)}%`,
    qty: `${(machArea / 100).toFixed(0)} cm²`,
    min: (machArea / (Df * 0.15 * vfF * rfF)) * tol.fin * kfac,
  });

  /* ═══════════ 6. the hole schedule ═══════════ */
  const HOLE_OVERHEAD = 0.075;            // rapid to position, Z approach, retract, per hole
  const key = (h) => `${h.d.toFixed(2)}|${h.akey}|${h.purpose}|${h.blind ? "B" : "T"}|${h.depth.toFixed(0)}`;
  const grp = new Map();
  for (const h of fx.holes) {
    const k = key(h);
    if (!grp.has(k)) grp.set(k, []);
    grp.get(k).push(h);
  }

  const schedule = [];
  let spotN = 0, tapEntries = 0;
  const drillOps = [], reamOps = [], tapOps = [], boreOps = [];

  for (const g of [...grp.values()].sort((a, b) => b[0].d - a[0].d)) {
    const h = g[0], n = g.length;
    const row = { d: h.d, n, depth: h.depth, ld: h.ld, blind: h.blind, purpose: h.purpose,
                  tap: h.tap, akey: h.akey, peck: h.peck.mode, seq: [], regs: g.flatMap((q) => q.regs) };

    if (h.d > 20) {
      /* too big to drill: helical interpolation, then a bore finish pass */
      const stepD = Df * 0.35 * kfac, laps = Math.ceil(h.depth / stepD);
      const path = Math.PI * (h.d - Df) * laps;
      const t1 = (path / vfF + 0.3) * n;
      row.seq.push({ op: `Helical interpolate, ${laps} laps`, tool: `Ø${Df.toFixed(0)} end mill`, min: t1 });
      const t2 = ((Math.PI * (h.d - Df) * 2) / (vfF * 0.5) + 0.25) * n;
      row.seq.push({ op: `Bore finish, 2 passes`, tool: "Adjustable boring head", min: t2 });
      boreOps.push({ d: h.d, n, min: t1 + t2, regs: row.regs, laps });
      row.totalMin = t1 + t2;
      schedule.push(row);
      continue;
    }

    /* spot drill when position matters or the hole is large */
    const needSpot = tol.v <= 0.10 || h.d > 12;
    if (needSpot) { spotN += n; row.seq.push({ op: "Spot drill", tool: "Ø10 90° spotting drill", min: n * (0.055 + HOLE_OVERHEAD * 0.6) }); }

    /* pilot the big ones in anything tougher than aluminium */
    let pilotMin = 0;
    if (h.d > 13 && mat.k >= 2.0) {
      const dp = h.d * 0.55, np = rpm(dp, 0.6), fnp = Math.min(0.38, 0.025 * dp) * (mat.k <= 2.5 ? 0.75 : 0.5);
      const pk = peckPlan(h.depth / dp);
      pilotMin = n * (((h.depth + 0.3 * dp) / (np * fnp * pk.feedScale)) * pk.fac + HOLE_OVERHEAD);
      row.seq.push({ op: `Pilot drill Ø${dp.toFixed(1)}`, tool: `Ø${dp.toFixed(1)} carbide`, min: pilotMin });
    }

    /* the drill itself */
    const nD = rpm(h.d, 0.6);
    const fn = Math.min(0.38, 0.025 * h.d) * (mat.k <= 1.2 ? 1 : mat.k <= 2.5 ? 0.75 : mat.k <= 4 ? 0.55 : 0.4);
    const vfD = nD * fn * h.peck.feedScale;
    const cutLen = h.depth + 0.3 * h.d + (h.through ? 0.3 * h.d : 0);
    const drillMin = n * ((cutLen / vfD) * h.peck.fac + HOLE_OVERHEAD);
    row.seq.push({
      op: `Drill${h.peck.mode === "straight" ? "" : ", " + h.peck.mode}`,
      tool: `Ø${h.d.toFixed(2)} carbide${h.ld > 5 ? ", through coolant" : ""}`,
      min: drillMin, rpm: nD, vf: vfD,
    });
    drillOps.push({ d: h.d, n, depth: h.depth, ld: h.ld, peck: h.peck.mode, min: drillMin + pilotMin, regs: row.regs, rpm: nD, vf: vfD });

    /* ream only true fit holes — never a tapped or clearance hole */
    if (h.purpose === "fit" && tol.v <= 0.05 && h.ld <= 8) {
      const passes = tol.v <= 0.025 ? 2 : 1;
      if (h.d <= 13) {
        const vfRe = rpm(h.d, 0.3) * 0.25;
        const t = n * ((h.depth / vfRe) * passes + 0.18 + HOLE_OVERHEAD);
        row.seq.push({ op: `Ream to ${tol.label}`, tool: `Ø${h.d.toFixed(2)} H7 chucking reamer`, min: t });
        reamOps.push({ d: h.d, n, min: t, regs: row.regs, passes, kind: "ream" });
      } else {
        const vfBo = rpm(h.d, 0.25) * 0.12;
        const t = n * ((h.depth / vfBo) * (passes + 1) + 0.45 + HOLE_OVERHEAD);
        row.seq.push({ op: `Single-point bore to ${tol.label}`, tool: `Ø${h.d.toFixed(2)} boring bar`, min: t });
        reamOps.push({ d: h.d, n, min: t, regs: row.regs, passes, kind: "bore" });
      }
    }

    /* tap */
    if (h.purpose === "tapped" && tapOn && h.tap) {
      const nT = Math.min(1200 / Math.pow(mat.k, 0.6), rpm(h.d, 0.25));
      const vfT = nT * h.tap.pitch;
      const td = h.through ? h.depth : Math.max(0, h.depth - 2 * h.tap.pitch);
      const t = n * ((2 * td) / vfT + 0.14 + HOLE_OVERHEAD);
      row.seq.push({ op: `Tap ${h.tap.name}`, tool: `${h.tap.name} spiral ${h.through ? "point" : "flute"}, rigid`, min: t });
      tapOps.push({ name: h.tap.name, d: h.d, n, min: t, regs: row.regs, rpm: nT, pitch: h.tap.pitch });
      tapEntries += n;
    }
    row.totalMin = row.seq.reduce((s, x) => s + x.min, 0);
    schedule.push(row);
  }

  if (spotN) ops.push({
    id: "spot", label: `Spot drill ×${spotN}`, tool: "Ø10 90° spotting drill",
    detail: tol.v <= 0.10 ? `${tol.label} needs a true start` : "large diameters need a start",
    qty: `${spotN} starts`, min: spotN * (0.055 + HOLE_OVERHEAD * 0.6),
  });
  for (const d of drillOps.sort((a, b) => b.d - a.d)) ops.push({
    id: `drill${d.d.toFixed(2)}_${d.depth.toFixed(0)}`,
    label: `Drill Ø${d.d.toFixed(2)} ×${d.n}`,
    tool: `Ø${d.d.toFixed(2)} carbide${d.ld > 5 ? ", through coolant" : ""}`,
    detail: `${d.depth.toFixed(1)} mm · ${d.ld.toFixed(1)}×D · ${d.peck} · ${d.rpm.toFixed(0)} rpm · ${d.vf.toFixed(0)} mm/min`,
    qty: `${d.n} hole${d.n === 1 ? "" : "s"}`, min: d.min, regs: d.regs,
  });
  for (const b of boreOps) ops.push({
    id: `bore${b.d.toFixed(1)}`, label: `Interpolate and bore Ø${b.d.toFixed(2)} ×${b.n}`,
    tool: `Ø${Df.toFixed(0)} end mill, then boring head`,
    detail: `${b.laps} helical laps, then 2 finish passes`, qty: `${b.n} bore${b.n === 1 ? "" : "s"}`,
    min: b.min, regs: b.regs,
  });
  for (const kind of ["ream", "bore"]) {
    const g = reamOps.filter((r) => r.kind === kind);
    if (!g.length) continue;
    ops.push({
      id: kind === "ream" ? "ream" : "boreFinish",
      label: `${kind === "ream" ? "Ream" : "Single-point bore"} ×${g.reduce((s, r) => s + r.n, 0)}`,
      tool: kind === "ream" ? "H7 chucking reamers" : "Boring bar, fine feed",
      detail: `${tol.label} on diameter · ${g[0].passes + (kind === "bore" ? 1 : 0)} passes · fit holes only`,
      qty: g.map((r) => `${r.n}×Ø${r.d.toFixed(2)}`).join(" · "),
      min: g.reduce((s, r) => s + r.min, 0), regs: g.flatMap((r) => r.regs),
    });
  }
  const tapAgg = new Map();
  for (const t of tapOps) {
    if (!tapAgg.has(t.name)) tapAgg.set(t.name, { ...t });
    else { const e = tapAgg.get(t.name); e.n += t.n; e.min += t.min; e.regs = [...e.regs, ...t.regs]; }
  }
  for (const t of tapAgg.values()) ops.push({
    id: `tap${t.name}`, label: `Tap ${t.name} ×${t.n}`,
    tool: `${t.name} spiral tap, rigid tapping`,
    detail: `${t.pitch.toFixed(2)} mm pitch · ${t.rpm.toFixed(0)} rpm · ${(t.rpm * t.pitch).toFixed(0)} mm/min`,
    qty: `${t.n} thread${t.n === 1 ? "" : "s"}`, min: t.min, regs: t.regs,
  });

  /* ---- 7. counterbores ---- */
  const cbores = fx.pockets.filter((p) => p.cbore);
  if (cbores.length) ops.push({
    id: "cbore", label: `Counterbore ×${cbores.length}`, tool: `Ø${(Df * 0.6).toFixed(1)} end mill, helical`,
    detail: `mean depth ${(cbores.reduce((s, p) => s + p.depth, 0) / cbores.length).toFixed(1)} mm`,
    qty: `${cbores.length} seats`, min: cbores.length * (0.22 + 0.02 * cbores[0].depth) * kfac,
    regs: cbores.flatMap((p) => p.regs),
  });

  /* ---- 8. chamfers, plus a lead-in on every tapped hole ---- */
  const chamLen = fx.chamferLen + tapEntries * Math.PI * 8;
  if (chamLen > 1) {
    const Dc = 8, vfC = rpm(Dc) * mat.fz * 4 * 0.7;
    ops.push({
      id: "chamfer", label: "Chamfer edges and thread lead-ins", tool: `Ø${Dc} 90° chamfer mill`,
      detail: `${vfC.toFixed(0)} mm/min · ${fx.chamfers.length} band${fx.chamfers.length === 1 ? "" : "s"}${tapEntries ? ` · ${tapEntries} lead-ins` : ""}`,
      qty: `${chamLen.toFixed(0)} mm of edge`,
      min: (chamLen / vfC) * 1.6 * kfac, regs: fx.chamfers.map((c) => c.r),
    });
  }

  /* ---- 9. ball-nose contouring ---- */
  if (contourArea > 40) {
    const Db = Math.max(4, Df * 0.5);
    const vfB = rpm(Db) * mat.fz * 0.5 * 2;
    ops.push({
      id: "contour", label: "Contour radii and blends", tool: `Ø${Db.toFixed(0)} ball nose`,
      detail: `${(Db * 0.08).toFixed(2)} mm stepover · scallop ${(Db * 0.08 * Db * 0.08 / (2 * Db) * 1000).toFixed(1)} µm`,
      qty: `${(contourArea / 100).toFixed(0)} cm², ${fx.fillets.length} radii`,
      min: (contourArea / (Db * 0.08 * vfB)) * tol.fin * 0.6 * kfac,
      regs: [...fx.fillets.flatMap((f) => f.regs), ...fx.curved.map((f) => f.r)],
    });
  }

  /* ---- 10. non-cutting: tool changes and rapids ---- */
  const toolChanges = ops.length;
  const moves = fx.holes.length * 2 + fx.pockets.length * 2 + toolChanges * 2 + 12;
  const rapidMin = moves * (0.35 * fx.diag / 25000 + 0.022);
  ops.push({
    id: "index", label: "Tool changes and rapids", tool: `${toolChanges} tools · ${fx.setups} setup${fx.setups === 1 ? "" : "s"}`,
    detail: `13 s per change · ${moves} positioning moves · ${(0.35 * fx.diag).toFixed(0)} mm mean travel`,
    qty: `${toolChanges} changes`, min: toolChanges * 0.22 + rapidMin,
  });

  /* ---- 11. deburr ---- */
  ops.push({
    id: "deburr", label: "Deburr and edge break", tool: "Bench: files, scotchbrite, air",
    detail: "off the machine, runs parallel to the next part",
    qty: `${fx.holes.length} hole edges`,
    min: 1.6 + an.area / 20000 + fx.holes.length * 0.14 + fx.chamferLen / 900,
  });

  const cycle = ops.reduce((s, o) => s + o.min, 0);
  const holeMin = schedule.reduce((s, r) => s + r.totalMin, 0);
  const maxMin = Math.max(...ops.map((o) => o.min));
  ops.forEach((o) => (o.heat = Math.min(1, Math.pow(o.min / maxMin, 0.6))));

  const regHeat = new Float32Array(an.regions.length).fill(0.06);
  ops.forEach((o) => { if (o.regs) o.regs.forEach((r) => { regHeat[r] = Math.max(regHeat[r], o.heat); }); });

  /* setup time scales with size and how many faces have to be presented */
  const kg = (fx.stockVol / 1e6) * mat.rho;
  const setupMin = fx.setups * (14 + Math.min(30, kg * 1.6) + (tol.v <= 0.05 ? 8 : 0));
  const toolCost = (removeVol / 1e6) * mat.tc + fx.holes.length * 0.35 * mat.k + tapEntries * 0.9 * mat.k;

  const faceMin = fc.total;
  return { ops, cycle, holeMin, schedule, regHeat, setups: fx.setups, toolChanges,
           st, stockVol, faceMin, facing: fc,
           removeVol, openVol, pocketVol, kfac, setupMin, toolCost, cornerLimited,
           Dr, Df, reachPct: rfF * 100, dirCount: [...fx.dirs.values()].filter((v) => v >= 2).length };
}

/* ---------------- cost for one part ---------------- */
function cost(pl, an, fx, mat, tol, rate) {
  const stockKg = (pl.stockVol / 1e6) * mat.rho;
  const lines = [
    { id: "stock", label: "Stock", heat: 0.18,
      sub: `${mat.name} · ${pl.st.stock.map((v) => v.toFixed(1)).join(" × ")} mm · ${stockKg.toFixed(2)} kg`,
      v: stockKg * mat.kg * 1.18 },
    { id: "machine", label: "Machine time", heat: 0.95,
      sub: `${fmtMin(pl.cycle)} cycle at $${rate}/hr`,
      v: (pl.cycle / 60) * rate },
    { id: "setup", label: "Setup and fixturing", heat: 0.55,
      sub: `${pl.setups} setup${pl.setups === 1 ? "" : "s"} · ${fmtMin(pl.setupMin)} · ${pl.toolChanges} tools loaded`,
      v: (pl.setupMin / 60) * rate },
    { id: "program", label: "CAM programming", heat: 0.4,
      sub: `${pl.toolChanges} operations posted and proved out`,
      v: 145 + pl.toolChanges * 11 },
    { id: "inspect", label: "Inspection", heat: tol.v <= 0.05 ? 0.9 : 0.3,
      sub: `${tol.gauge} · ${tol.insp} min`,
      v: (tol.insp / 60) * rate * 0.8 },
    { id: "tooling", label: "Cutter wear", heat: 0.5,
      sub: `${(pl.removeVol / 1e6).toFixed(2)} L removed at $${mat.tc}/L · ${fx.holes.length} holes`,
      v: pl.toolCost },
    { id: "finish", label: "Deburr and finishing", heat: 0.2,
      sub: `${fx.chamferLen.toFixed(0)} mm of edge, clean and pack`,
      v: 1.8 + fx.chamferLen / 900 + pl.setups * 0.9 },
  ];
  const subtotal = lines.reduce((s, l) => s + l.v, 0);
  const risk = tol.scrap * (1 + (mat.k - 1) * 0.2);
  return {
    lines, subtotal, risk, total: subtotal * (1 + risk), stockKg,
    hours: { cycle: pl.cycle, setup: pl.setupMin, total: pl.cycle + pl.setupMin },
    lead: Math.ceil(2 + pl.setups * 1.1 + pl.cycle / 120),
  };
}

const fmtMin = (m) => {
  let t = Math.round(m * 60);
  const h = Math.floor(t / 3600); t -= h * 3600;
  const mm = Math.floor(t / 60), ss = t - mm * 60;
  return h > 0 ? `${h}:${String(mm).padStart(2, "0")}:${String(ss).padStart(2, "0")}`
               : `${mm}:${String(ss).padStart(2, "0")}`;
};
const money = (n) => "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* cost ramp — the only color in the interface */
const RAMP = [[34, 214, 176], [255, 197, 61], [255, 59, 92]];
function ramp(t) {
  t = Math.max(0, Math.min(1, t));
  const [a, b, i] = t < 0.5 ? [RAMP[0], RAMP[1], t * 2] : [RAMP[1], RAMP[2], (t - 0.5) * 2];
  return [0, 1, 2].map((k) => a[k] + (b[k] - a[k]) * i);
}
const rgba = (c, a) => `rgba(${c[0] | 0},${c[1] | 0},${c[2] | 0},${a})`;

/* ---------------------------------------------------------------------------
   Optional: fold in ground truth from a SolidWorks macro.
   Anything present in the manifest replaces what was inferred from the mesh.
   --------------------------------------------------------------------------- */
export function applyManifest(an, fx, manifest = {}) {
  const out = { an: { ...an }, fx: { ...fx }, overrides: [] };
  if (manifest.volume > 0) { out.an.vol = manifest.volume; out.overrides.push("volume"); }
  if (manifest.area > 0) { out.an.area = manifest.area; out.overrides.push("area"); }
  if (Array.isArray(manifest.holes) && manifest.holes.length) {
    /* trust the Hole Wizard: exact diameter, depth, thread and blind flag */
    out.fx.holes = manifest.holes.map((h, i) => {
      const near = fx.holes.find((q) => Math.abs(q.d - h.d) < 0.25) || fx.holes[0] || {};
      const depth = h.depth ?? near.depth ?? 0;
      const d = h.d;
      return {
        ...near, d, depth, ld: depth / d, blind: !!h.blind, through: !h.blind,
        entry: near.entry ?? "hi", axis: h.axis ?? near.axis ?? [0, 0, 1],
        akey: near.akey ?? "0,0,1", center: h.center ?? near.center ?? [0, 0, 0],
        regs: near.regs ?? [], peck: peckPlan(depth / d),
        tap: h.thread ? { name: h.thread, pitch: h.pitch ?? 1 } : null,
        purpose: h.thread ? "tapped" : d > 20 ? "bore" : h.clearance ? "clearance" : "fit",
        fromCAD: true,
      };
    });
    out.overrides.push(`${manifest.holes.length} holes from the feature tree`);
  }
  if (manifest.minInternalR != null) { out.fx.minInternalR = manifest.minInternalR; out.overrides.push("corner radius"); }
  if (manifest.setups > 0) { out.fx.setups = manifest.setups; out.overrides.push("setups"); }
  return out;
}

export function materialByName(name = "") {
  const n = name.toLowerCase().replace(/[^a-z0-9]/g, "");
  return MATERIALS.find((m) => n.includes(m.name.toLowerCase().replace(/[^a-z0-9]/g, "")))
      || MATERIALS.find((m) => n.includes(m.id))
      || MATERIALS[0];
}

export {
  parseSTL, analyze, extract, makeStock, facingOps, plan, cost,
  peckPlan, tapFor, insideSolid, fmtMin, money, ramp, rgba,
  MATERIALS, TOLS, STOCK_MODES, PLATE_MM, PLATE_IN, RPM_MAX, DUTY, ALLOW, TAPS,
};
