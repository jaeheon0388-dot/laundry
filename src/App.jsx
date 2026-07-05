import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import { doc, setDoc, onSnapshot } from "firebase/firestore";

/* ───────── 세탁 서비스 관리 앱 ─────────
   기능: 일일 작업 · 정산(명세서) · 기사 관리 · 고객사 관리 · 운영 메모
   데이터는 Firebase Firestore에 저장되어 모든 사용자가 실시간으로 공유합니다.
*/

const dataRef = doc(db, "app", "data"); // 전체 데이터가 저장되는 Firestore 문서
const todayStr = () => new Date().toISOString().slice(0, 10);
const monthOf = (d) => d.slice(0, 7);
const fmt = (n) => (n || 0).toLocaleString("ko-KR");

const uid = () => Math.random().toString(36).slice(2, 9);

const seedData = () => {
  const c1 = uid(), c2 = uid(), c3 = uid(), c4 = uid();
  const d1 = uid(), d2 = uid();
  const m = monthOf(todayStr());
  return {
    customers: [
      { id: c1, name: "연세사우나", region: "김포", billingType: "sheet", items: [{ name: "수건", price: 75 }, { name: "옷", price: 400 }], monthlyFee: 0, startMonth: m, mapMemo: "정문 옆 계단 아래 수거함", driverId: d1 },
      { id: c2, name: "김포보석사우나", region: "김포", billingType: "sheet", items: [{ name: "수건", price: 80 }, { name: "옷", price: 400 }], monthlyFee: 0, startMonth: m, mapMemo: "", driverId: d1 },
      { id: c3, name: "아쿠아사우나", region: "양천", billingType: "bag", items: [{ name: "포대", price: 15000 }], monthlyFee: 0, startMonth: m, mapMemo: "지하 1층 카운터에 전달", driverId: d2 },
      { id: c4, name: "청학골프클럽", region: "화성", billingType: "flat", items: [], monthlyFee: 800000, startMonth: m, mapMemo: "프로샵 뒤편 수거장", driverId: d2 },
    ],
    drivers: [
      { id: d1, name: "PV5 검정색", phone: "" },
      { id: d2, name: "ST1 로만", phone: "" },
    ],
    counts: {},        // { [date]: { [customerId]: { [itemName]: count } } }
    delivery: {},      // { [date]: { [customerId]: true } }
    deliveryBags: {},  // { [date]: { [customerId]: { pickup, deliver } } } 기사 수거/배송 포대 기록
    payments: {},      // { [YYYY-MM]: { [customerId]: 입금액 } }
    memos: [],         // [{id, author, date, text}]
  };
};

// 결제 방식 라벨
const BILLING_LABEL = { sheet: "장당", bag: "포대당", flat: "월정액" };

// 다음 달 계산 (YYYY-MM)
const nextMonth = (m) => {
  const [y, mo] = m.split("-").map(Number);
  const d = new Date(y, mo, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
};

// 특정 월의 고객사 청구액 계산
const monthCharge = (data, cust, month) => {
  if (cust.billingType === "flat") {
    return (cust.startMonth || "0000-00") <= month ? Number(cust.monthlyFee) || 0 : 0;
  }
  let sum = 0;
  for (const [d, byCust] of Object.entries(data.counts)) {
    if (monthOf(d) !== month) continue;
    const items = byCust[cust.id];
    if (!items) continue;
    for (const [iname, cnt] of Object.entries(items)) {
      const price = cust.items.find((i) => i.name === iname)?.price || 0;
      sum += (Number(cnt) || 0) * price;
    }
  }
  return sum;
};

// 특정 월 이전까지의 미수금 이월액 계산 (청구 누적 - 입금 누적)
const carryOver = (data, cust, month) => {
  const months = new Set();
  for (const d of Object.keys(data.counts)) months.add(monthOf(d));
  for (const m of Object.keys(data.payments || {})) months.add(m);
  if (cust.billingType === "flat" && cust.startMonth) {
    let m = cust.startMonth;
    let guard = 0;
    while (m < month && guard < 240) { months.add(m); m = nextMonth(m); guard++; }
  }
  let bal = 0;
  for (const m of [...months].filter((x) => x < month)) {
    bal += monthCharge(data, cust, m) - (Number(data.payments?.[m]?.[cust.id]) || 0);
  }
  return bal;
};

/* ───────── 공용 스타일 ───────── */
const S = {
  page: { fontFamily: "'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif", background: "#eef4f9", minHeight: "100vh", paddingBottom: 90, color: "#1b2a3a" },
  header: { background: "linear-gradient(135deg,#1867c0,#33a3e0)", color: "#fff", padding: "20px 18px 16px", borderRadius: "0 0 22px 22px" },
  card: { background: "#fff", borderRadius: 16, padding: 16, margin: "12px 14px", boxShadow: "0 2px 8px rgba(20,60,110,.07)" },
  btn: { border: "none", borderRadius: 10, padding: "10px 14px", fontWeight: 700, cursor: "pointer", fontSize: 14 },
  btnBlue: { background: "linear-gradient(90deg,#2196f3,#26c6da)", color: "#fff" },
  btnGhost: { background: "#eef4f9", color: "#1867c0" },
  btnDanger: { background: "#ffecef", color: "#e0355a" },
  input: { width: "100%", boxSizing: "border-box", border: "1.5px solid #d7e4ef", borderRadius: 10, padding: "10px 12px", fontSize: 15, outline: "none", background: "#fbfdff" },
  tag: { display: "inline-block", background: "#e3f2fd", color: "#1867c0", borderRadius: 8, padding: "2px 8px", fontSize: 12, fontWeight: 700 },
  subTabs: { display: "flex", gap: 6, overflowX: "auto", padding: "10px 14px 0" },
  subTab: (act) => ({ flexShrink: 0, padding: "9px 14px", borderRadius: 12, fontSize: 13.5, fontWeight: 700, border: "none", cursor: "pointer", background: act ? "#1867c0" : "#fff", color: act ? "#fff" : "#5b7186", boxShadow: act ? "0 3px 8px rgba(24,103,192,.3)" : "0 1px 4px rgba(20,60,110,.06)" }),
};

/* ───────── 메인 앱 ───────── */
export default function App() {
  const [data, setData] = useState(null);
  const [tab, setTab] = useState("counts"); // counts | drivers | customers | memos
  const saveTimer = useRef(null);

  // Firestore에서 불러오기 + 실시간 동기화 (다른 사람이 수정하면 즉시 반영)
  useEffect(() => {
    const normalize = (d) => ({
      ...d,
      payments: d.payments || {},
      deliveryBags: d.deliveryBags || {},
      customers: (d.customers || []).map((c) => ({
        billingType: "sheet", monthlyFee: 0, startMonth: monthOf(todayStr()),
        ...c,
      })),
    });
    const unsub = onSnapshot(
      dataRef,
      (snap) => {
        if (snap.metadata.hasPendingWrites) return; // 내가 방금 쓴 건 이미 화면에 반영됨
        if (snap.exists()) {
          setData(normalize(snap.data()));
        } else {
          const seed = normalize(seedData());
          setDoc(dataRef, seed).catch(() => {});
          setData(seed);
        }
      },
      (err) => {
        console.error("Firestore 연결 오류:", err);
        alert("데이터베이스 연결에 실패했습니다. Firestore 보안 규칙을 확인해주세요.");
      }
    );
    return unsub;
  }, []);

  // 저장 (디바운스 후 Firestore에 기록)
  const update = (fn) => {
    setData((prev) => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(() => {
        setDoc(dataRef, next).catch((e) => console.error("저장 실패:", e));
      }, 400);
      return next;
    });
  };

  if (!data) return <div style={{ ...S.page, display: "flex", alignItems: "center", justifyContent: "center" }}>불러오는 중…</div>;

  return (
    <div style={S.page}>
      <div style={S.header}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 38, height: 38, borderRadius: "50%", background: "rgba(255,255,255,.2)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 13 }}>세탁</div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 18 }}>세탁 서비스 관리</div>
            <div style={{ fontSize: 12.5, opacity: 0.85 }}>린넨·유니폼 렌탈 통합 관리</div>
          </div>
        </div>
      </div>

      {tab === "counts" && <CountsScreen data={data} update={update} />}
      {tab === "settle" && <SettlementScreen data={data} update={update} />}
      {tab === "drivers" && <DriversScreen data={data} update={update} />}
      {tab === "customers" && <CustomersScreen data={data} update={update} />}
      {tab === "memos" && <MemosScreen data={data} update={update} />}

      {/* 하단 네비게이션 */}
      <nav style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "#fff", borderTop: "1px solid #e2ecf4", display: "flex", zIndex: 50 }}>
        {[
          ["counts", "📋", "작업"],
          ["settle", "💰", "정산"],
          ["drivers", "🚚", "기사"],
          ["customers", "🏪", "고객사"],
          ["memos", "📝", "메모"],
        ].map(([k, ic, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{ flex: 1, border: "none", background: "none", padding: "10px 0 12px", cursor: "pointer", color: tab === k ? "#1867c0" : "#93a7b8", fontWeight: tab === k ? 800 : 600, fontSize: 12 }}>
            <div style={{ fontSize: 20 }}>{ic}</div>
            {label}
          </button>
        ))}
      </nav>
    </div>
  );
}

/* ───────── 1. 일일 장당제 ───────── */
function CountsScreen({ data, update }) {
  const [date, setDate] = useState(todayStr());
  const month = monthOf(date);
  const dayCounts = data.counts[date] || {};

  // 합계 계산
  let dayTotal = 0, dayRevenue = 0, monthTotal = 0;
  for (const [d, byCust] of Object.entries(data.counts)) {
    for (const [cid, items] of Object.entries(byCust)) {
      const cust = data.customers.find((c) => c.id === cid);
      for (const [iname, cnt] of Object.entries(items)) {
        const n = Number(cnt) || 0;
        if (monthOf(d) === month) monthTotal += n;
        if (d === date) {
          dayTotal += n;
          const price = cust?.items.find((i) => i.name === iname)?.price || 0;
          dayRevenue += n * price;
        }
      }
    }
  }

  const setCount = (cid, iname, val) => {
    update((prev) => {
      const counts = { ...prev.counts };
      const day = { ...(counts[date] || {}) };
      const cust = { ...(day[cid] || {}) };
      if (val === "" || Number(val) === 0) delete cust[iname];
      else cust[iname] = Number(val);
      if (Object.keys(cust).length) day[cid] = cust; else delete day[cid];
      if (Object.keys(day).length) counts[date] = day; else delete counts[date];
      return { ...prev, counts };
    });
  };

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 2 }}>일일 작업 조회</div>
        <div style={{ fontSize: 12, color: "#93a7b8", marginBottom: 8 }}>입력하는 즉시 자동 저장되고 정산 탭에 바로 반영됩니다</div>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={S.input} />
        <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
          <Stat label={`${Number(date.slice(8))}일 총 장수`} value={`${fmt(dayTotal)} 장`} color="#1867c0" />
          <Stat label="당일 매출" value={`₩ ${fmt(dayRevenue)}`} color="#0e9f6e" />
          <Stat label={`${Number(month.slice(5))}월 누적`} value={`${fmt(monthTotal)} 장`} color="#7c5cd6" />
        </div>
      </div>

      {data.customers.map((c) => {
        const cc = dayCounts[c.id] || {};
        const custDayTotal = Object.values(cc).reduce((a, b) => a + Number(b || 0), 0);
        // 월 누적 (고객사별)
        let custMonth = 0;
        for (const [d, byCust] of Object.entries(data.counts)) {
          if (monthOf(d) !== month) continue;
          for (const v of Object.values(byCust[c.id] || {})) custMonth += Number(v) || 0;
        }
        const unit = c.billingType === "bag" ? "포대" : "장";

        // 월정액 고객사: 청구는 월정액 고정, 포대 수는 기록용으로 입력
        if (c.billingType === "flat") {
          const bagCnt = cc["포대"] ?? "";
          return (
            <div key={c.id} style={S.card}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontWeight: 800, fontSize: 15.5 }}>{c.name}</div>
                <span style={{ ...S.tag, background: "#f3ecfd", color: "#7c5cd6" }}>월정액</span>
              </div>
              <div style={{ fontSize: 12.5, color: "#7d94a8", margin: "4px 0 10px" }}>
                월 ₩{fmt(c.monthlyFee)} 고정 청구 · 오늘 {fmt(Number(bagCnt) || 0)}포대 · {Number(month.slice(5))}월 누적 {fmt(custMonth)}포대
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderTop: "1px solid #f0f5fa" }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>포대</div>
                  <div style={{ fontSize: 12, color: "#93a7b8" }}>기록용 (청구액에는 영향 없음)</div>
                </div>
                <button style={{ ...S.btn, ...S.btnGhost, padding: "6px 12px" }} onClick={() => setCount(c.id, "포대", Math.max(0, (Number(bagCnt) || 0) - 1))}>−1</button>
                <input
                  type="number" min="0" value={bagCnt}
                  placeholder="0"
                  onChange={(e) => setCount(c.id, "포대", e.target.value)}
                  style={{ ...S.input, width: 80, textAlign: "center", padding: "8px 6px" }}
                />
                <button style={{ ...S.btn, ...S.btnGhost, padding: "6px 12px" }} onClick={() => setCount(c.id, "포대", (Number(bagCnt) || 0) + 1)}>+1</button>
              </div>
            </div>
          );
        }

        return (
          <div key={c.id} style={S.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontWeight: 800, fontSize: 15.5 }}>{c.name}</div>
              <div style={{ display: "flex", gap: 6 }}>
                <span style={{ ...S.tag, background: c.billingType === "bag" ? "#fff3e0" : "#e3f2fd", color: c.billingType === "bag" ? "#c77800" : "#1867c0" }}>{BILLING_LABEL[c.billingType] || "장당"}</span>
                <span style={S.tag}>{c.region || "지역 미지정"}</span>
              </div>
            </div>
            <div style={{ fontSize: 12.5, color: "#7d94a8", margin: "4px 0 10px" }}>
              오늘 {fmt(custDayTotal)}{unit} · {Number(month.slice(5))}월 누적 {fmt(custMonth)}{unit}
            </div>
            {c.items.map((it) => (
              <div key={it.name} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderTop: "1px solid #f0f5fa" }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>{it.name}</div>
                  <div style={{ fontSize: 12, color: "#93a7b8" }}>₩{fmt(it.price)} / {unit}</div>
                </div>
                <button style={{ ...S.btn, ...S.btnGhost, padding: "6px 12px" }} onClick={() => setCount(c.id, it.name, Math.max(0, (Number(cc[it.name]) || 0) - (c.billingType === "bag" ? 1 : 10)))}>−{c.billingType === "bag" ? 1 : 10}</button>
                <input
                  type="number" min="0" value={cc[it.name] ?? ""}
                  placeholder="0"
                  onChange={(e) => setCount(c.id, it.name, e.target.value)}
                  style={{ ...S.input, width: 80, textAlign: "center", padding: "8px 6px" }}
                />
                <button style={{ ...S.btn, ...S.btnGhost, padding: "6px 12px" }} onClick={() => setCount(c.id, it.name, (Number(cc[it.name]) || 0) + (c.billingType === "bag" ? 1 : 10))}>+{c.billingType === "bag" ? 1 : 10}</button>
              </div>
            ))}
          </div>
        );
      })}
      {data.customers.length === 0 && <Empty text="등록된 고객사가 없습니다. '고객사' 탭에서 먼저 추가하세요." />}
    </div>
  );
}

function Stat({ label, value, color }) {
  return (
    <div style={{ flex: 1, background: "#f4f9fd", borderRadius: 12, padding: "10px 8px", textAlign: "center" }}>
      <div style={{ fontSize: 11.5, color: "#7d94a8", fontWeight: 600 }}>{label}</div>
      <div style={{ fontWeight: 800, fontSize: 15, color, marginTop: 2 }}>{value}</div>
    </div>
  );
}

/* ───────── 정산 ───────── */
function SettlementScreen({ data, update }) {
  const [month, setMonth] = useState(monthOf(todayStr()));
  const [openId, setOpenId] = useState(null); // 일별 내역 펼친 고객사
  const [stmtId, setStmtId] = useState(null); // 명세서 모달 대상 고객사
  const [copied, setCopied] = useState(false);
  const pays = data.payments?.[month] || {};

  // 고객사의 일별 명세 행 만들기 (날짜·품목별)
  const buildStmtRows = (c) => {
    const rows = [];
    const dates = Object.keys(data.counts)
      .filter((d) => monthOf(d) === month && data.counts[d]?.[c.id])
      .sort();
    for (const d of dates) {
      for (const [iname, cnt] of Object.entries(data.counts[d][c.id])) {
        const n = Number(cnt) || 0;
        if (!n) continue;
        const unit = iname === "포대" ? "포대" : "장";
        const price = c.items.find((i) => i.name === iname)?.price || 0;
        rows.push({
          day: `${Number(d.slice(8))}일`, iname, n, unit,
          price: c.billingType === "flat" ? null : price,
          amt: c.billingType === "flat" ? null : n * price,
        });
      }
    }
    return rows;
  };

  // 명세서 전체 데이터 (표 형태 2차원 배열)
  const buildStmtAoa = (c) => {
    const charge = monthCharge(data, c, month);
    const carry = carryOver(data, c, month);
    const paid = Number(pays[c.id]) || 0;
    const total = charge + carry;
    const rows = buildStmtRows(c);
    const aoa = [
      [`${c.name} 정산 명세서`],
      [`청구월: ${month}`, "", `결제방식: ${BILLING_LABEL[c.billingType] || "장당"}`],
      [],
      ["날짜", "품목", "수량", "단가", "금액"],
      ...rows.map((r) => [r.day, r.iname, `${fmt(r.n)}${r.unit}`, r.price == null ? "-" : r.price, r.amt == null ? "-" : r.amt]),
      [],
    ];
    if (c.billingType === "flat") aoa.push(["", "", "", "월정액", Number(c.monthlyFee) || 0]);
    aoa.push(["", "", "", "당월 청구액", charge]);
    if (carry !== 0) aoa.push(["", "", "", "전월 이월", carry]);
    aoa.push(["", "", "", "받을 금액 합계", total]);
    aoa.push(["", "", "", "입금액", paid]);
    aoa.push(["", "", "", "남은 금액", total - paid]);
    return aoa;
  };

  // 엑셀(.xlsx) 다운로드
  const downloadXlsx = (c) => {
    const aoa = buildStmtAoa(c);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "명세서");
    try {
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${c.name}_${month}_명세서.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch {
      alert("이 환경에서는 다운로드가 차단되어 있습니다. '복사하기'로 복사한 뒤 엑셀에 붙여넣어 주세요.");
    }
  };

  // 텍스트 복사 (엑셀에 붙여넣으면 셀로 나뉘는 탭 구분 형식)
  const copyStmt = async (c) => {
    const text = buildStmtAoa(c)
      .map((row) => row.map((v) => (typeof v === "number" ? fmt(v) : v)).join("\t"))
      .join("\n");
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };

  const setPaid = (cid, val) => {
    update((prev) => {
      const payments = { ...(prev.payments || {}) };
      const m = { ...(payments[month] || {}) };
      if (val === "" || Number(val) === 0) delete m[cid];
      else m[cid] = Number(val);
      if (Object.keys(m).length) payments[month] = m; else delete payments[month];
      return { ...prev, payments };
    });
  };

  const rows = data.customers.map((c) => {
    const charge = monthCharge(data, c, month);
    const carry = carryOver(data, c, month);
    const paid = Number(pays[c.id]) || 0;
    const balance = charge + carry - paid;
    return { c, charge, carry, paid, balance };
  });

  const totCharge = rows.reduce((a, r) => a + r.charge, 0);
  const totCarry = rows.reduce((a, r) => a + r.carry, 0);
  const totPaid = rows.reduce((a, r) => a + r.paid, 0);
  const totBal = rows.reduce((a, r) => a + r.balance, 0);

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 8 }}>월별 정산</div>
        <input type="month" value={month} onChange={(e) => setMonth(e.target.value)} style={S.input} />
        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          <Stat label="당월 청구" value={`₩${fmt(totCharge)}`} color="#1867c0" />
          <Stat label="이월 미수" value={`₩${fmt(totCarry)}`} color="#c77800" />
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <Stat label="입금 합계" value={`₩${fmt(totPaid)}`} color="#0e9f6e" />
          <Stat label="남은 미수금" value={`₩${fmt(totBal)}`} color={totBal > 0 ? "#e0355a" : "#0e9f6e"} />
        </div>
      </div>

      {rows.map(({ c, charge, carry, paid, balance }) => {
        const total = charge + carry;
        const status = total <= 0 && paid === 0 ? "none" : balance <= 0 ? "done" : paid > 0 ? "partial" : "unpaid";
        const badge = {
          done: { text: "완납", bg: "#e3f8ee", fg: "#0e9f6e" },
          partial: { text: "부분 입금", bg: "#fff3e0", fg: "#c77800" },
          unpaid: { text: "미수", bg: "#ffecef", fg: "#e0355a" },
          none: { text: "청구 없음", bg: "#eef2f6", fg: "#93a7b8" },
        }[status];

        // 일별 작업 내역 (작업 탭 입력과 자동 연동)
        const daily = Object.keys(data.counts)
          .filter((d) => monthOf(d) === month && data.counts[d]?.[c.id])
          .sort()
          .map((d) => {
            const items = data.counts[d][c.id];
            let amt = 0;
            const parts = [];
            for (const [iname, cnt] of Object.entries(items)) {
              const n = Number(cnt) || 0;
              if (!n) continue;
              const unit = iname === "포대" ? "포대" : "장";
              parts.push(`${iname} ${fmt(n)}${unit === "포대" ? "포대" : "장"}`);
              const price = c.items.find((i) => i.name === iname)?.price || 0;
              amt += n * price;
            }
            return { d, parts, amt };
          })
          .filter((r) => r.parts.length);
        const isOpen = openId === c.id;

        return (
          <div key={c.id} style={S.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontWeight: 800, fontSize: 15.5 }}>{c.name}</div>
              <div style={{ display: "flex", gap: 6 }}>
                <span style={{ ...S.tag, background: "#f3ecfd", color: "#7c5cd6" }}>{BILLING_LABEL[c.billingType] || "장당"}</span>
                <span style={{ ...S.tag, background: badge.bg, color: badge.fg }}>{badge.text}</span>
              </div>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginTop: 10, color: "#5b7186" }}>
              <span>당월 청구액</span><b style={{ color: "#1b2a3a" }}>₩{fmt(charge)}</b>
            </div>
            {carry !== 0 && (
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginTop: 4, color: "#5b7186" }}>
                <span>전월 이월 {carry > 0 ? "미수금" : "(과입금)"}</span>
                <b style={{ color: carry > 0 ? "#c77800" : "#0e9f6e" }}>₩{fmt(carry)}</b>
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginTop: 4, paddingTop: 6, borderTop: "1px dashed #e2ecf4", color: "#5b7186" }}>
              <span>받을 금액 합계</span><b style={{ color: "#1b2a3a" }}>₩{fmt(total)}</b>
            </div>

            {/* 일별 내역 · 명세서 */}
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <button
                style={{ ...S.btn, ...S.btnGhost, flex: 1, fontSize: 13 }}
                onClick={() => setOpenId(isOpen ? null : c.id)}
              >
                {isOpen ? "일별 내역 닫기 ▲" : `일별 내역 (${daily.length}일) ▼`}
              </button>
              <button
                style={{ ...S.btn, ...S.btnGhost, flex: 1, fontSize: 13 }}
                onClick={() => setStmtId(c.id)}
              >
                📄 명세서
              </button>
            </div>
            {isOpen && (
              <div style={{ marginTop: 8, background: "#f7fafd", borderRadius: 12, padding: "4px 12px" }}>
                {daily.length === 0 && <div style={{ textAlign: "center", color: "#93a7b8", fontSize: 13, padding: "12px 0" }}>이번 달 작업 기록이 없습니다.</div>}
                {daily.map(({ d, parts, amt }) => (
                  <div key={d} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #eaf1f7", fontSize: 13 }}>
                    <span style={{ fontWeight: 800, color: "#1867c0", width: 38 }}>{Number(d.slice(8))}일</span>
                    <span style={{ flex: 1, color: "#5b7186" }}>{parts.join(" · ")}</span>
                    <b style={{ color: "#1b2a3a" }}>{c.billingType === "flat" ? "기록" : `₩${fmt(amt)}`}</b>
                  </div>
                ))}
                {daily.length > 0 && c.billingType !== "flat" && (
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", fontSize: 13, fontWeight: 800 }}>
                    <span style={{ color: "#5b7186" }}>합계</span>
                    <span>₩{fmt(daily.reduce((a, r) => a + r.amt, 0))}</span>
                  </div>
                )}
                {c.billingType === "flat" && daily.length > 0 && (
                  <div style={{ fontSize: 12, color: "#93a7b8", padding: "8px 0" }}>월정액 고객사로 청구액은 ₩{fmt(c.monthlyFee)} 고정, 포대 수는 기록용입니다.</div>
                )}
              </div>
            )}

            <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 10 }}>
              <input
                type="number" min="0" placeholder="입금액 입력"
                value={pays[c.id] ?? ""}
                onChange={(e) => setPaid(c.id, e.target.value)}
                style={{ ...S.input, flex: 1 }}
              />
              <button style={{ ...S.btn, ...S.btnBlue, whiteSpace: "nowrap" }} onClick={() => setPaid(c.id, total > 0 ? total : "")}>완납 처리</button>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, marginTop: 10, fontWeight: 800 }}>
              <span style={{ color: "#5b7186" }}>남은 금액</span>
              <span style={{ color: balance > 0 ? "#e0355a" : "#0e9f6e" }}>₩{fmt(balance)}</span>
            </div>
          </div>
        );
      })}
      {data.customers.length === 0 && <Empty text="등록된 고객사가 없습니다." />}

      {/* 명세서 모달 */}
      {stmtId && (() => {
        const c = data.customers.find((x) => x.id === stmtId);
        if (!c) return null;
        const aoa = buildStmtAoa(c);
        const cell = { border: "1px solid #cbd8e4", padding: "6px 8px", fontSize: 12.5, whiteSpace: "nowrap" };
        return (
          <div
            style={{ position: "fixed", inset: 0, background: "rgba(10,25,45,.55)", zIndex: 100, display: "flex", alignItems: "flex-end", justifyContent: "center" }}
            onClick={() => setStmtId(null)}
          >
            <div
              style={{ background: "#fff", borderRadius: "20px 20px 0 0", width: "100%", maxWidth: 560, maxHeight: "85vh", display: "flex", flexDirection: "column" }}
              onClick={(e) => e.stopPropagation()}
            >
              <div style={{ padding: "16px 18px 10px", borderBottom: "1px solid #eef2f6", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 16 }}>{c.name} 명세서</div>
                  <div style={{ fontSize: 12.5, color: "#7d94a8" }}>{month} · {BILLING_LABEL[c.billingType] || "장당"}</div>
                </div>
                <button style={{ ...S.btn, ...S.btnGhost, padding: "8px 12px" }} onClick={() => setStmtId(null)}>닫기 ✕</button>
              </div>

              {/* 엑셀 느낌 미리보기 */}
              <div style={{ overflow: "auto", padding: "12px 18px", flex: 1 }}>
                <table style={{ borderCollapse: "collapse", width: "100%" }}>
                  <tbody>
                    {aoa.map((row, ri) => {
                      const isHeader = ri === 3;
                      const isSummary = ri > 3 && row[3] && row[0] === "";
                      if (row.length === 0) return <tr key={ri}><td style={{ height: 8 }} colSpan={5}></td></tr>;
                      return (
                        <tr key={ri} style={{ background: isHeader ? "#e8f1fa" : isSummary ? "#f7fafd" : "#fff" }}>
                          {ri < 2 ? (
                            <td colSpan={5} style={{ ...cell, border: "none", fontWeight: ri === 0 ? 800 : 600, fontSize: ri === 0 ? 15 : 12.5, color: ri === 0 ? "#1b2a3a" : "#5b7186" }}>
                              {row.filter(Boolean).join("   ")}
                            </td>
                          ) : (
                            [0, 1, 2, 3, 4].map((ci) => (
                              <td key={ci} style={{ ...cell, fontWeight: isHeader || (isSummary && ci >= 3) ? 700 : 400, textAlign: ci >= 2 ? "right" : "left" }}>
                                {typeof row[ci] === "number" ? fmt(row[ci]) : (row[ci] ?? "")}
                              </td>
                            ))
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div style={{ padding: "12px 18px 20px", borderTop: "1px solid #eef2f6", display: "flex", gap: 8 }}>
                <button style={{ ...S.btn, ...S.btnBlue, flex: 1 }} onClick={() => downloadXlsx(c)}>⬇ 엑셀 다운로드</button>
                <button style={{ ...S.btn, ...S.btnGhost, flex: 1 }} onClick={() => copyStmt(c)}>{copied ? "복사됨 ✓" : "📋 복사하기"}</button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

/* ───────── 2. 기사 관리 ───────── */
function DriversScreen({ data, update }) {
  const [sub, setSub] = useState("list"); // list | assign | work | records
  return (
    <div>
      <div style={S.subTabs}>
        {[["list", "배송 기사님 관리"], ["assign", "고객사 배정"], ["work", "기사 실사용 화면"], ["records", "배송 기록"]].map(([k, l]) => (
          <button key={k} style={S.subTab(sub === k)} onClick={() => setSub(k)}>{l}</button>
        ))}
      </div>
      {sub === "list" && <DriverList data={data} update={update} />}
      {sub === "assign" && <DriverAssign data={data} update={update} />}
      {sub === "work" && <DriverWork data={data} update={update} />}
      {sub === "records" && <DeliveryRecords data={data} />}
    </div>
  );
}

function DriverList({ data, update }) {
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [editId, setEditId] = useState(null);

  const save = () => {
    if (!name.trim()) return;
    update((prev) => {
      if (editId) {
        return { ...prev, drivers: prev.drivers.map((d) => (d.id === editId ? { ...d, name: name.trim(), phone } : d)) };
      }
      return { ...prev, drivers: [...prev.drivers, { id: uid(), name: name.trim(), phone }] };
    });
    setName(""); setPhone(""); setEditId(null);
  };

  const remove = (id) => {
    if (!window.confirm("이 기사를 삭제할까요? 배정된 고객사는 미배정 상태가 됩니다.")) return;
    update((prev) => ({
      ...prev,
      drivers: prev.drivers.filter((d) => d.id !== id),
      customers: prev.customers.map((c) => (c.driverId === id ? { ...c, driverId: null } : c)),
    }));
  };

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, marginBottom: 10 }}>{editId ? "기사 정보 수정" : "기사 추가"}</div>
        <input style={S.input} placeholder="기사명 (예: PV5 검정색 한성기)" value={name} onChange={(e) => setName(e.target.value)} />
        <input style={{ ...S.input, marginTop: 8 }} placeholder="연락처 (선택)" value={phone} onChange={(e) => setPhone(e.target.value)} />
        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
          <button style={{ ...S.btn, ...S.btnBlue, flex: 1 }} onClick={save}>{editId ? "수정 저장" : "+ 기사 추가"}</button>
          {editId && <button style={{ ...S.btn, ...S.btnGhost }} onClick={() => { setEditId(null); setName(""); setPhone(""); }}>취소</button>}
        </div>
      </div>

      {data.drivers.map((d) => {
        const cnt = data.customers.filter((c) => c.driverId === d.id).length;
        return (
          <div key={d.id} style={S.card}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <Avatar text={d.name[0]} />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 800 }}>{d.name}</div>
                <div style={{ fontSize: 12.5, color: "#7d94a8" }}>배정 고객사 {cnt}곳{d.phone ? ` · ${d.phone}` : ""}</div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button style={{ ...S.btn, ...S.btnGhost }} onClick={() => { setEditId(d.id); setName(d.name); setPhone(d.phone || ""); window.scrollTo(0, 0); }}>✏️ 수정</button>
              <button style={{ ...S.btn, ...S.btnDanger }} onClick={() => remove(d.id)}>삭제</button>
            </div>
          </div>
        );
      })}
      {data.drivers.length === 0 && <Empty text="등록된 기사가 없습니다." />}
    </div>
  );
}

function DriverAssign({ data, update }) {
  const [openId, setOpenId] = useState(null);
  const toggle = (custId, driverId) => {
    update((prev) => ({
      ...prev,
      customers: prev.customers.map((c) => (c.id === custId ? { ...c, driverId: c.driverId === driverId ? null : driverId } : c)),
    }));
  };
  return (
    <div>
      <div style={{ ...S.card, paddingBottom: 8 }}>
        <div style={{ fontWeight: 800 }}>고객사 배정</div>
        <div style={{ fontSize: 12.5, color: "#7d94a8", marginTop: 4, marginBottom: 8 }}>기사 카드를 누르면 해당 기사의 고객사 배정 화면이 열립니다.</div>
      </div>
      {data.drivers.map((d) => {
        const assigned = data.customers.filter((c) => c.driverId === d.id);
        const open = openId === d.id;
        return (
          <div key={d.id} style={S.card}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }} onClick={() => setOpenId(open ? null : d.id)}>
              <Avatar text={d.name[0]} />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 800 }}>{d.name}</div>
                <div style={{ fontSize: 12.5, color: "#7d94a8" }}>배정 고객사 {assigned.length}곳 · 눌러서 배정 관리</div>
              </div>
              <div style={{ color: "#93a7b8", fontSize: 18, transform: open ? "rotate(90deg)" : "none", transition: "transform .2s" }}>›</div>
            </div>
            {open && (
              <div style={{ marginTop: 12, borderTop: "1px solid #f0f5fa", paddingTop: 8 }}>
                {data.customers.map((c) => {
                  const mine = c.driverId === d.id;
                  const other = c.driverId && !mine ? data.drivers.find((x) => x.id === c.driverId)?.name : null;
                  return (
                    <label key={c.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: "1px solid #f6fafd", cursor: "pointer" }}>
                      <input type="checkbox" checked={mine} onChange={() => toggle(c.id, d.id)} style={{ width: 18, height: 18 }} />
                      <div style={{ flex: 1 }}>
                        <div style={{ fontWeight: 700, fontSize: 14 }}>{c.name}</div>
                        {other && <div style={{ fontSize: 11.5, color: "#e0355a" }}>현재 {other} 담당 → 체크하면 이 기사로 이동</div>}
                      </div>
                      <span style={S.tag}>{c.region || "-"}</span>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function DriverWork({ data, update }) {
  const [driverId, setDriverId] = useState(data.drivers[0]?.id || null);
  const [date, setDate] = useState(todayStr());
  const stops = data.customers.filter((c) => c.driverId === driverId);
  const dayDone = data.delivery[date] || {};
  const dayBags = data.deliveryBags?.[date] || {};
  const done = stops.filter((c) => dayDone[c.id]).length;

  // ── 드래그앤드롭 순서 변경 ──
  const [dragId, setDragId] = useState(null);
  const dragRef = useRef(null);      // 현재 드래그 중인 고객사 id
  const stopsRef = useRef(stops);    // 최신 순서 참조
  const itemRefs = useRef({});       // 카드 DOM 참조
  stopsRef.current = stops;

  const toggleDone = (cid) => {
    update((prev) => {
      const delivery = { ...prev.delivery };
      const day = { ...(delivery[date] || {}) };
      if (day[cid]) delete day[cid]; else day[cid] = true;
      if (Object.keys(day).length) delivery[date] = day; else delete delivery[date];
      return { ...prev, delivery };
    });
  };

  // 수거/배송 포대 수 기록 (자동 저장)
  const setBags = (cid, key, val) => {
    update((prev) => {
      const deliveryBags = { ...(prev.deliveryBags || {}) };
      const day = { ...(deliveryBags[date] || {}) };
      const rec = { ...(day[cid] || {}) };
      if (val === "" || Number(val) === 0) delete rec[key]; else rec[key] = Number(val);
      if (Object.keys(rec).length) day[cid] = rec; else delete day[cid];
      if (Object.keys(day).length) deliveryBags[date] = day; else delete deliveryBags[date];
      return { ...prev, deliveryBags };
    });
  };

  // 배정된 고객사끼리 순서 교체 (전체 customers 배열 내 위치 유지)
  const reorder = (fromId, toId) => {
    update((prev) => {
      const assignedIdx = prev.customers
        .map((c, i) => (c.driverId === driverId ? i : -1))
        .filter((i) => i !== -1);
      const ordered = assignedIdx.map((i) => prev.customers[i]);
      const from = ordered.findIndex((c) => c.id === fromId);
      const to = ordered.findIndex((c) => c.id === toId);
      if (from === -1 || to === -1 || from === to) return prev;
      const [moved] = ordered.splice(from, 1);
      ordered.splice(to, 0, moved);
      const arr = [...prev.customers];
      assignedIdx.forEach((pos, k) => { arr[pos] = ordered[k]; });
      return { ...prev, customers: arr };
    });
  };

  const startPress = (e, cid) => {
    if (e.target.closest("button,input")) return; // 버튼·입력칸 터치는 드래그 제외
    const startY = e.clientY;
    let activated = false;

    const activate = () => {
      activated = true;
      dragRef.current = cid;
      setDragId(cid);
      if (navigator.vibrate) navigator.vibrate(30); // 드래그 시작 진동 피드백
    };
    const timer = setTimeout(activate, 500); // 0.5초 길게 누르면 드래그 시작

    const blockScroll = (ev) => { if (dragRef.current != null) ev.preventDefault(); };

    const onMove = (ev) => {
      const y = ev.clientY;
      if (y == null) return;
      if (!activated) {
        // 0.5초 전에 10px 이상 움직이면 스크롤로 판단하고 취소
        if (Math.abs(y - startY) > 10) cleanup();
        return;
      }
      for (const c of stopsRef.current) {
        if (c.id === dragRef.current) continue;
        const el = itemRefs.current[c.id];
        if (!el) continue;
        const r = el.getBoundingClientRect();
        if (y >= r.top && y <= r.bottom) {
          reorder(dragRef.current, c.id);
          break;
        }
      }
    };
    const cleanup = () => {
      clearTimeout(timer);
      dragRef.current = null;
      setDragId(null);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", cleanup);
      window.removeEventListener("pointercancel", cleanup);
      window.removeEventListener("touchmove", blockScroll);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", cleanup);
    window.addEventListener("pointercancel", cleanup);
    window.addEventListener("touchmove", blockScroll, { passive: false }); // 드래그 중 화면 스크롤 방지
  };

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, marginBottom: 8 }}>기사 실사용 화면</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
          {data.drivers.map((d) => (
            <button key={d.id} style={S.subTab(driverId === d.id)} onClick={() => setDriverId(d.id)}>{d.name}</button>
          ))}
        </div>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={S.input} />
        <div style={{ display: "flex", gap: 10, marginTop: 10 }}>
          <Stat label="완료" value={done} color="#0e9f6e" />
          <Stat label="미완료" value={stops.length - done} color="#e0355a" />
        </div>
      </div>

      {/* 기사 화면 (다크 카드) */}
      <div style={{ ...S.card, background: "linear-gradient(160deg,#132437,#1c3a57)", color: "#eaf3fb" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ fontWeight: 800, letterSpacing: 1, fontSize: 13, opacity: 0.7 }}>DELIVERY</div>
          <span style={{ background: "#2b9cf2", borderRadius: 8, padding: "3px 10px", fontSize: 12, fontWeight: 700 }}>
            {Number(date.slice(5, 7))}월 {Number(date.slice(8))}일
          </span>
        </div>
        <div style={{ fontWeight: 800, fontSize: 18, marginBottom: 4 }}>{data.drivers.find((d) => d.id === driverId)?.name || "기사 선택"}</div>
        <div style={{ fontSize: 12, opacity: 0.55, marginBottom: 12 }}>카드를 0.5초 꾹 누른 뒤 위아래로 끌면 배송 순서가 바뀝니다</div>

        {stops.map((c, i) => {
          const isDone = !!dayDone[c.id];
          const dragging = dragId === c.id;
          return (
            <div
              key={c.id}
              ref={(el) => { itemRefs.current[c.id] = el; }}
              onPointerDown={(e) => startPress(e, c.id)}
              onContextMenu={(e) => e.preventDefault()}
              style={{
                background: dragging ? "rgba(43,156,242,.22)" : "rgba(255,255,255,.06)",
                border: dragging ? "1.5px solid #2b9cf2" : "1px solid rgba(255,255,255,.1)",
                borderRadius: 14, padding: 14, marginBottom: 12,
                transform: dragging ? "scale(1.02)" : "none",
                boxShadow: dragging ? "0 8px 20px rgba(0,0,0,.35)" : "none",
                transition: "background .15s, transform .15s",
                cursor: dragging ? "grabbing" : "grab",
                userSelect: "none", WebkitUserSelect: "none", WebkitTouchCallout: "none",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ background: "#2b9cf2", borderRadius: 8, padding: "2px 8px", fontSize: 12, fontWeight: 800 }}>{i + 1}번</span>
                <div style={{ fontWeight: 800, fontSize: 15, flex: 1 }}>{c.name}</div>
                <span style={{ fontSize: 11.5, background: isDone ? "rgba(14,159,110,.25)" : "rgba(224,53,90,.25)", color: isDone ? "#5ee6b0" : "#ff9db4", borderRadius: 8, padding: "3px 8px", fontWeight: 700 }}>
                  {isDone ? "완료" : "미완료"}
                </span>
              </div>
              {c.mapMemo && (
                <div style={{ marginTop: 8, fontSize: 12.5, lineHeight: 1.5, opacity: 0.85 }}>
                  <span style={{ opacity: 0.6 }}>지도 메모 · </span>{c.mapMemo}
                </div>
              )}
              {/* 수거/배송 포대 입력 */}
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                {[["pickup", "수거 포대"], ["deliver", "배송 포대"]].map(([k, l]) => (
                  <div key={k} style={{ flex: 1 }}>
                    <div style={{ fontSize: 11.5, opacity: 0.65, marginBottom: 4, fontWeight: 700 }}>{l}</div>
                    <input
                      type="number" min="0" placeholder="0"
                      value={(dayBags[c.id] || {})[k] ?? ""}
                      onChange={(e) => setBags(c.id, k, e.target.value)}
                      style={{
                        width: "100%", boxSizing: "border-box", borderRadius: 10, padding: "9px 10px",
                        fontSize: 15, textAlign: "center", outline: "none",
                        background: "rgba(255,255,255,.1)", border: "1px solid rgba(255,255,255,.2)", color: "#fff",
                      }}
                    />
                  </div>
                ))}
              </div>
              <button
                onClick={() => toggleDone(c.id)}
                style={{ ...S.btn, width: "100%", marginTop: 10, background: isDone ? "rgba(255,255,255,.12)" : "linear-gradient(90deg,#2196f3,#26c6da)", color: "#fff" }}
              >
                {isDone ? "완료 취소" : "배송 완료"}
              </button>
            </div>
          );
        })}
        {stops.length === 0 && <div style={{ textAlign: "center", opacity: 0.6, padding: "20px 0" }}>배정된 고객사가 없습니다.</div>}
      </div>
    </div>
  );
}

function DeliveryRecords({ data }) {
  const [date, setDate] = useState(todayStr());
  const day = data.delivery[date] || {};
  const doneList = data.customers.filter((c) => day[c.id]);

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, marginBottom: 8 }}>배송 기록</div>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={S.input} />
        <div style={{ marginTop: 10 }}><span style={S.tag}>총 {doneList.length}건 완료</span></div>
      </div>
      {data.drivers.map((d) => {
        const mine = doneList.filter((c) => c.driverId === d.id);
        return (
          <div key={d.id} style={S.card}>
            <div style={{ fontWeight: 800, marginBottom: 6 }}>{d.name} <span style={{ color: "#7d94a8", fontWeight: 600, fontSize: 13 }}>· {mine.length}건</span></div>
            {mine.length === 0 ? (
              <div style={{ border: "1.5px dashed #d7e4ef", borderRadius: 12, textAlign: "center", padding: "18px 0", color: "#93a7b8", fontSize: 13.5 }}>배송 기록 없음</div>
            ) : (
              mine.map((c) => {
                const bags = (data.deliveryBags?.[date] || {})[c.id] || {};
                return (
                  <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #f0f5fa", fontSize: 14 }}>
                    <div>
                      <div style={{ fontWeight: 700 }}>{c.name}</div>
                      {(bags.pickup || bags.deliver) && (
                        <div style={{ fontSize: 12, color: "#7d94a8", marginTop: 2 }}>
                          {bags.pickup ? `수거 ${fmt(bags.pickup)}포대` : ""}{bags.pickup && bags.deliver ? " · " : ""}{bags.deliver ? `배송 ${fmt(bags.deliver)}포대` : ""}
                        </div>
                      )}
                    </div>
                    <span style={{ color: "#0e9f6e", fontWeight: 700 }}>완료 ✓</span>
                  </div>
                );
              })
            )}
          </div>
        );
      })}
    </div>
  );
}

/* ───────── 3. 고객사 관리 ───────── */
function CustomersScreen({ data, update }) {
  const blank = { name: "", region: "", mapMemo: "", billingType: "sheet", items: [{ name: "수건", price: "" }], bagPrice: "", monthlyFee: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [showForm, setShowForm] = useState(false);

  const save = () => {
    if (!form.name.trim()) return;
    let items = [];
    if (form.billingType === "sheet") {
      items = form.items.filter((i) => i.name.trim()).map((i) => ({ name: i.name.trim(), price: Number(i.price) || 0 }));
    } else if (form.billingType === "bag") {
      items = [{ name: "포대", price: Number(form.bagPrice) || 0 }];
    }
    const payload = {
      name: form.name.trim(), region: form.region, mapMemo: form.mapMemo,
      billingType: form.billingType, items,
      monthlyFee: form.billingType === "flat" ? Number(form.monthlyFee) || 0 : 0,
    };
    update((prev) => {
      if (editId) {
        return { ...prev, customers: prev.customers.map((c) => (c.id === editId ? { ...c, ...payload } : c)) };
      }
      return { ...prev, customers: [...prev.customers, { id: uid(), ...payload, startMonth: monthOf(todayStr()), driverId: null }] };
    });
    setForm(blank); setEditId(null); setShowForm(false);
  };

  const remove = (id) => {
    if (!window.confirm("고객사를 삭제할까요? 관련 장당제·정산 기록도 화면에서 제외됩니다.")) return;
    update((prev) => ({ ...prev, customers: prev.customers.filter((c) => c.id !== id) }));
  };

  const setItem = (idx, key, val) => {
    setForm((f) => ({ ...f, items: f.items.map((it, i) => (i === idx ? { ...it, [key]: val } : it)) }));
  };

  const startEdit = (c) => {
    setEditId(c.id);
    setForm({
      name: c.name, region: c.region || "", mapMemo: c.mapMemo || "",
      billingType: c.billingType || "sheet",
      items: c.billingType === "sheet" ? c.items.map((i) => ({ ...i })) : [{ name: "수건", price: "" }],
      bagPrice: c.billingType === "bag" ? (c.items[0]?.price ?? "") : "",
      monthlyFee: c.billingType === "flat" ? (c.monthlyFee ?? "") : "",
    });
    setShowForm(true);
    window.scrollTo(0, 0);
  };

  return (
    <div>
      <div style={S.card}>
        <button style={{ ...S.btn, ...S.btnBlue, width: "100%" }} onClick={() => { setShowForm((v) => !v); setEditId(null); setForm(blank); }}>
          {showForm ? "닫기" : "+ 고객사 추가"}
        </button>
        {showForm && (
          <div style={{ marginTop: 12 }}>
            <input style={S.input} placeholder="고객사명 (예: 연세사우나)" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            <input style={{ ...S.input, marginTop: 8 }} placeholder="지역 (예: 김포)" value={form.region} onChange={(e) => setForm({ ...form, region: e.target.value })} />
            <textarea style={{ ...S.input, marginTop: 8, minHeight: 60 }} placeholder="지도 메모 (수거·배송 위치 안내)" value={form.mapMemo} onChange={(e) => setForm({ ...form, mapMemo: e.target.value })} />

            <div style={{ fontWeight: 700, fontSize: 13.5, margin: "12px 0 6px", color: "#5b7186" }}>결제 방식</div>
            <div style={{ display: "flex", gap: 6 }}>
              {[["sheet", "장당"], ["bag", "포대당"], ["flat", "월정액"]].map(([k, l]) => (
                <button key={k} style={{ ...S.subTab(form.billingType === k), flex: 1, textAlign: "center" }} onClick={() => setForm({ ...form, billingType: k })}>{l}</button>
              ))}
            </div>

            {form.billingType === "sheet" && (
              <div>
                <div style={{ fontWeight: 700, fontSize: 13.5, margin: "12px 0 6px", color: "#5b7186" }}>품목 · 장당 단가</div>
                {form.items.map((it, i) => (
                  <div key={i} style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                    <input style={{ ...S.input, flex: 1 }} placeholder="품목명" value={it.name} onChange={(e) => setItem(i, "name", e.target.value)} />
                    <input style={{ ...S.input, width: 100 }} type="number" placeholder="단가" value={it.price} onChange={(e) => setItem(i, "price", e.target.value)} />
                    <button style={{ ...S.btn, ...S.btnDanger, padding: "6px 10px" }} onClick={() => setForm((f) => ({ ...f, items: f.items.filter((_, x) => x !== i) }))}>✕</button>
                  </div>
                ))}
                <button style={{ ...S.btn, ...S.btnGhost, width: "100%" }} onClick={() => setForm((f) => ({ ...f, items: [...f.items, { name: "", price: "" }] }))}>+ 품목 추가</button>
              </div>
            )}
            {form.billingType === "bag" && (
              <div>
                <div style={{ fontWeight: 700, fontSize: 13.5, margin: "12px 0 6px", color: "#5b7186" }}>포대당 단가</div>
                <input style={S.input} type="number" placeholder="포대당 금액 (예: 15000)" value={form.bagPrice} onChange={(e) => setForm({ ...form, bagPrice: e.target.value })} />
                <div style={{ fontSize: 12, color: "#93a7b8", marginTop: 6 }}>장당제 탭에서 매일 포대 수를 입력하면 포대 수 × 단가로 청구됩니다.</div>
              </div>
            )}
            {form.billingType === "flat" && (
              <div>
                <div style={{ fontWeight: 700, fontSize: 13.5, margin: "12px 0 6px", color: "#5b7186" }}>월정액 금액</div>
                <input style={S.input} type="number" placeholder="매월 청구 금액 (예: 800000)" value={form.monthlyFee} onChange={(e) => setForm({ ...form, monthlyFee: e.target.value })} />
                <div style={{ fontSize: 12, color: "#93a7b8", marginTop: 6 }}>일일 입력 없이 매월 이 금액이 자동 청구됩니다.</div>
              </div>
            )}

            <button style={{ ...S.btn, ...S.btnBlue, width: "100%", marginTop: 12 }} onClick={save}>{editId ? "수정 저장" : "고객사 등록"}</button>
          </div>
        )}
      </div>

      {data.customers.map((c) => (
        <div key={c.id} style={S.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontWeight: 800, fontSize: 15.5 }}>{c.name}</div>
            <div style={{ display: "flex", gap: 6 }}>
              <span style={{ ...S.tag, background: "#f3ecfd", color: "#7c5cd6" }}>{BILLING_LABEL[c.billingType] || "장당"}</span>
              <span style={S.tag}>{c.region || "-"}</span>
            </div>
          </div>
          <div style={{ fontSize: 12.5, color: "#7d94a8", margin: "4px 0" }}>
            담당: {data.drivers.find((d) => d.id === c.driverId)?.name || "미배정"}
          </div>
          <div style={{ fontSize: 13, color: "#5b7186" }}>
            {c.billingType === "flat"
              ? `월정액 ₩${fmt(c.monthlyFee)} / 월`
              : c.items.map((i) => `${i.name} ₩${fmt(i.price)}`).join(" · ")}
          </div>
          {c.mapMemo && <div style={{ fontSize: 12.5, color: "#93a7b8", marginTop: 6, background: "#f7fafd", borderRadius: 8, padding: "8px 10px" }}>📍 {c.mapMemo}</div>}
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...S.btn, ...S.btnGhost }} onClick={() => startEdit(c)}>✏️ 수정</button>
            <button style={{ ...S.btn, ...S.btnDanger }} onClick={() => remove(c.id)}>삭제</button>
          </div>
        </div>
      ))}
      {data.customers.length === 0 && <Empty text="등록된 고객사가 없습니다. 위에서 추가하세요." />}
    </div>
  );
}

/* ───────── 4. 운영 메모 ───────── */
function MemosScreen({ data, update }) {
  const [text, setText] = useState("");
  const [editId, setEditId] = useState(null);

  const save = () => {
    if (!text.trim()) return;
    update((prev) => {
      if (editId) {
        return { ...prev, memos: prev.memos.map((m) => (m.id === editId ? { ...m, text: text.trim() } : m)) };
      }
      return { ...prev, memos: [{ id: uid(), author: "관리자", date: todayStr(), text: text.trim() }, ...prev.memos] };
    });
    setText(""); setEditId(null);
  };

  return (
    <div>
      <div style={S.card}>
        <div style={{ fontWeight: 800, marginBottom: 4 }}>운영 메모</div>
        <div style={{ fontSize: 12.5, color: "#7d94a8", marginBottom: 10 }}>관리자끼리 공유할 내용을 입력하세요. 예: 은평 루트 오전 배송 우선 확인</div>
        <textarea style={{ ...S.input, minHeight: 90 }} value={text} onChange={(e) => setText(e.target.value)} placeholder="공유할 메모를 입력하세요" />
        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
          <button style={{ ...S.btn, ...S.btnBlue, flex: 1 }} onClick={save}>{editId ? "수정 저장" : "+ 메모 추가"}</button>
          {editId && <button style={{ ...S.btn, ...S.btnGhost }} onClick={() => { setEditId(null); setText(""); }}>취소</button>}
        </div>
      </div>

      {data.memos.map((m) => (
        <div key={m.id} style={{ ...S.card, background: "#fffcf2", border: "1px solid #f3e8c8" }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: "#a08d55", fontWeight: 700 }}>
            <span>📌 {m.author}</span><span>{m.date}</span>
          </div>
          <div style={{ marginTop: 8, fontSize: 14.5, lineHeight: 1.55, whiteSpace: "pre-wrap" }}>{m.text}</div>
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...S.btn, ...S.btnGhost }} onClick={() => { setEditId(m.id); setText(m.text); window.scrollTo(0, 0); }}>✏️ 수정</button>
            <button style={{ ...S.btn, ...S.btnDanger }} onClick={() => update((prev) => ({ ...prev, memos: prev.memos.filter((x) => x.id !== m.id) }))}>삭제</button>
          </div>
        </div>
      ))}
      {data.memos.length === 0 && <Empty text="아직 메모가 없습니다." />}
    </div>
  );
}

/* ───────── 공용 컴포넌트 ───────── */
function Avatar({ text }) {
  return (
    <div style={{ width: 40, height: 40, borderRadius: 12, background: "#e3f2fd", color: "#1867c0", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 15, flexShrink: 0 }}>
      {text}
    </div>
  );
}
function Empty({ text }) {
  return <div style={{ ...S.card, textAlign: "center", color: "#93a7b8", padding: "26px 16px", fontSize: 14 }}>{text}</div>;
}
