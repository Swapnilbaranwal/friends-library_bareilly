import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

type SeatStatus = 'vacant' | 'booked' | 'due_soon' | 'overdue';

type SeatRecord = {
  seatNo: number;
  studentName: string;
  joiningDate?: string;
  feeDueDate?: string;
  paidThroughMonth?: string;
  active: boolean;
  phone?: string;
  notes?: string;
};

type SeatViewModel = SeatRecord & {
  status: SeatStatus;
  statusLabel: string;
  dueInDays?: number;
};

const TOTAL_SEATS = 50;
const DUE_SOON_WINDOW = 3;
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const SEAT_LAYOUT_ROWS: Array<{ left: number[]; right: number[] }> = [
  { left: [50, 49, 48, 47, 46], right: [41, 42, 43, 44, 45] },
  { left: [40, 39, 38, 37, 36], right: [31, 32, 33, 34, 35] },
  { left: [30, 29, 28, 27, 26], right: [21, 22, 23, 24, 25] },
  { left: [20, 19, 18, 17, 16], right: [11, 12, 13, 14, 15] },
  { left: [10, 9, 8, 7, 6], right: [1, 2, 3, 4, 5] },
];

const emptySeat = (seatNo: number): SeatRecord => ({
  seatNo,
  studentName: '',
  active: false,
});

const demoData: SeatRecord[] = [
  {
    seatNo: 1,
    studentName: 'Aarav Sharma',
    joiningDate: '2026-01-10',
    feeDueDate: '2026-06-02',
    active: true,
    phone: '9876543210',
    notes: 'Morning batch',
  },
  {
    seatNo: 2,
    studentName: 'Meera Patel',
    joiningDate: '2026-02-01',
    feeDueDate: '2026-05-30',
    active: true,
    phone: '9123456780',
    notes: 'Overdue fee follow-up',
  },
  {
    seatNo: 3,
    studentName: 'Kabir Singh',
    joiningDate: '2026-04-18',
    feeDueDate: '2026-06-03',
    active: true,
    phone: '9988776655',
  },
  {
    seatNo: 7,
    studentName: 'Riya Mehta',
    joiningDate: '2026-03-21',
    feeDueDate: '2026-05-31',
    active: true,
  },
  {
    seatNo: 14,
    studentName: 'Yash Verma',
    joiningDate: '2026-04-05',
    feeDueDate: '2026-06-10',
    active: true,
  },
  {
    seatNo: 26,
    studentName: 'Sara Khan',
    joiningDate: '2026-04-09',
    feeDueDate: '2026-06-01',
    active: true,
    notes: 'Evening batch',
  },
  {
    seatNo: 31,
    studentName: 'Vihaan Joshi',
    joiningDate: '2026-01-25',
    feeDueDate: '2026-05-27',
    active: true,
  },
  {
    seatNo: 45,
    studentName: 'Ananya Rao',
    joiningDate: '2026-03-14',
    feeDueDate: '2026-06-12',
    active: true,
  },
];

const headerAliases: Record<string, string[]> = {
  seatNo: ['seat', 'seat no', 'seat number', 'seat_no', 'seatno'],
  studentName: ['name', 'student', 'student name', 'student_name', 'studentname'],
  joiningDate: ['joining date', 'join date', 'date of joining', 'joining_date'],
  feeDueDate: [
    'due date',
    'fee due date',
    'fees due date',
    'fee_due_date',
    'duedate',
    'next fee date',
    'fees paid',
    'fee paid',
    'fees paid date',
  ],
  active: ['active', 'status', 'student active', 'is active'],
  phone: ['phone', 'mobile', 'contact', 'contact number'],
  notes: ['notes', 'remark', 'remarks', 'comment'],
};

const normalize = (value: unknown) =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');

const toIsoDate = (year: number, month: number, day: number) => {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    Number.isNaN(date.getTime()) ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return undefined;
  }
  return date.toISOString().slice(0, 10);
};

const parseDDMMYY = (raw: string) => {
  const digits = raw.replace(/\D/g, '');
  if (digits.length !== 6) return undefined;

  const day = Number(digits.slice(0, 2));
  const month = Number(digits.slice(2, 4));
  const shortYear = Number(digits.slice(4, 6));
  const fullYear = shortYear >= 70 ? 1900 + shortYear : 2000 + shortYear;
  return toIsoDate(fullYear, month, day);
};

const parseDate = (value: unknown) => {
  if (!value) return undefined;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    // Joining dates can be entered as DDMMYY in numeric form, e.g. 290526.
    if (value >= 100000 && value <= 999999) {
      return parseDDMMYY(String(value));
    }

    const date = XLSX.SSF.parse_date_code(value);
    if (!date) return undefined;
    return toIsoDate(date.y, date.m, date.d);
  }

  const asString = String(value).trim();
  const ddmmyyDate = parseDDMMYY(asString);
  if (ddmmyyDate) return ddmmyyDate;

  const parsed = new Date(asString);
  return Number.isNaN(parsed.getTime()) ? asString : parsed.toISOString().slice(0, 10);
};

const getHeaderKey = (header: string) => {
  const normalized = normalize(header);
  return (Object.entries(headerAliases).find(([, aliases]) =>
    aliases.some((alias) => normalized === alias),
  )?.[0] ?? normalized) as keyof SeatRecord | string;
};

const parseBoolean = (value: unknown) => {
  const normalized = normalize(value);
  return ['yes', 'true', 'active', 'booked', '1', 'y'].includes(normalized);
};

const parseMonthHeader = (header: string) => {
  const match = header.trim().match(/^([A-Za-z]{3})-(\d{4})$/);
  if (!match) return undefined;

  const monthIndex = MONTH_NAMES.findIndex(
    (month) => month.toLowerCase() === match[1].toLowerCase(),
  );
  if (monthIndex < 0) return undefined;

  return Number(match[2]) * 12 + monthIndex;
};

const isPaidCell = (value: unknown) => {
  const normalized = normalize(value);
  return ['paid', 'yes', 'true', '1', 'done'].includes(normalized);
};

const getMonthIndexFromIsoDate = (isoDate: string) => {
  const date = new Date(isoDate);
  if (Number.isNaN(date.getTime())) return undefined;
  return date.getUTCFullYear() * 12 + date.getUTCMonth();
};

const formatMonthIndex = (monthIndex: number) => {
  const year = Math.floor(monthIndex / 12);
  const month = monthIndex % 12;
  return `${MONTH_NAMES[month]}-${year}`;
};

const getDayFromIsoDate = (isoDate: string) => {
  const date = new Date(isoDate);
  if (Number.isNaN(date.getTime())) return undefined;
  return date.getUTCDate();
};

const getCycleDateForMonth = (monthIndex: number, cycleDay: number) => {
  const year = Math.floor(monthIndex / 12);
  const monthZeroBased = monthIndex % 12;
  const daysInMonth = new Date(Date.UTC(year, monthZeroBased + 1, 0)).getUTCDate();
  const day = Math.min(cycleDay, daysInMonth);
  return toIsoDate(year, monthZeroBased + 1, day);
};

const addMonthsByCycle = (isoDate: string, monthsToAdd: number) => {
  const monthIndex = getMonthIndexFromIsoDate(isoDate);
  const cycleDay = getDayFromIsoDate(isoDate);
  if (typeof monthIndex !== 'number' || typeof cycleDay !== 'number') return undefined;
  return getCycleDateForMonth(monthIndex + monthsToAdd, cycleDay);
};

const getDaysUntil = (dateString?: string) => {
  if (!dateString) return undefined;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dueDate = new Date(dateString);
  dueDate.setHours(0, 0, 0, 0);
  const delta = Math.round((dueDate.getTime() - today.getTime()) / 86_400_000);
  return delta;
};

const buildSeats = (records: SeatRecord[]) => {
  const bySeat = new Map<number, SeatRecord>();

  records.forEach((record) => {
    if (record.seatNo >= 1 && record.seatNo <= TOTAL_SEATS) {
      bySeat.set(record.seatNo, record);
    }
  });

  return Array.from({ length: TOTAL_SEATS }, (_, index) => {
    const seatNo = index + 1;
    return bySeat.get(seatNo) ?? emptySeat(seatNo);
  });
};

const classifySeat = (seat: SeatRecord): SeatViewModel => {
  const isOccupied = seat.active;
  const dueInDays = getDaysUntil(seat.feeDueDate);

  if (!isOccupied) {
    return {
      ...seat,
      status: 'vacant',
      statusLabel: 'Vacant',
    };
  }

  if (typeof dueInDays === 'number' && dueInDays <= 0) {
    return {
      ...seat,
      status: 'overdue',
      statusLabel: 'Fee overdue',
      dueInDays,
    };
  }

  if (typeof dueInDays === 'number' && dueInDays <= DUE_SOON_WINDOW) {
    return {
      ...seat,
      status: 'due_soon',
      statusLabel: 'Fee due soon',
      dueInDays,
    };
  }

  return {
    ...seat,
    status: 'booked',
    statusLabel: 'Booked',
    dueInDays,
  };
};

const parseExcelBuffer = (buffer: ArrayBuffer) => {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: '',
    raw: true,
  });

  return rows.map((row) => {
    const mapped: Partial<SeatRecord> = {};
    const paidMonthIndices: number[] = [];

    Object.entries(row).forEach(([header, value]) => {
      const key = getHeaderKey(header);

      if (key === 'seatNo') {
        mapped.seatNo = Number(value);
      } else if (key === 'studentName') {
        mapped.studentName = String(value ?? '').trim();
      } else if (key === 'joiningDate') {
        mapped.joiningDate = parseDate(value);
      } else if (key === 'feeDueDate') {
        mapped.feeDueDate = parseDate(value);
      } else if (key === 'active') {
        mapped.active = parseBoolean(value);
      } else if (key === 'phone') {
        mapped.phone = String(value ?? '').trim();
      } else if (key === 'notes') {
        mapped.notes = String(value ?? '').trim();
      } else {
        const monthIndex = parseMonthHeader(header);
        if (typeof monthIndex === 'number' && isPaidCell(value)) {
          paidMonthIndices.push(monthIndex);
        }
      }
    });

    const joiningDate = mapped.joiningDate;
    const joiningMonthIndex = joiningDate ? getMonthIndexFromIsoDate(joiningDate) : undefined;
    const cycleDay = joiningDate ? getDayFromIsoDate(joiningDate) : undefined;

    const paidAfterJoining =
      typeof joiningMonthIndex === 'number'
        ? paidMonthIndices.filter((monthIndex) => monthIndex >= joiningMonthIndex)
        : paidMonthIndices;

    const latestPaidMonth = paidAfterJoining.length ? Math.max(...paidAfterJoining) : undefined;

    let computedFeeDueDate = mapped.feeDueDate;
    if (typeof latestPaidMonth === 'number' && typeof cycleDay === 'number') {
      computedFeeDueDate = getCycleDateForMonth(latestPaidMonth + 1, cycleDay) ?? mapped.feeDueDate;
    } else if (joiningDate) {
      computedFeeDueDate = addMonthsByCycle(joiningDate, 1) ?? mapped.feeDueDate;
    }

    const paidThroughMonth =
      typeof latestPaidMonth === 'number' ? formatMonthIndex(latestPaidMonth) : undefined;

    return {
      seatNo: mapped.seatNo ?? 0,
      studentName: mapped.studentName ?? '',
      joiningDate,
      feeDueDate: computedFeeDueDate,
      paidThroughMonth,
      active: mapped.active ?? false,
      phone: mapped.phone,
      notes: mapped.notes,
    } satisfies SeatRecord;
  });
};

const parseExcelFile = async (file: File) => {
  const buffer = await file.arrayBuffer();
  return parseExcelBuffer(buffer);
};

const formatRelative = (days?: number) => {
  if (typeof days !== 'number') return 'No due date';
  if (days === 0) return 'Due today';
  if (days > 0) return `${days} day${days === 1 ? '' : 's'} left`;
  const overdue = Math.abs(days);
  return `${overdue} day${overdue === 1 ? '' : 's'} overdue`;
};

export default function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [records, setRecords] = useState<SeatRecord[]>([]);
  const [selectedSeat, setSelectedSeat] = useState<SeatViewModel | null>(null);
  const [isDrawerOpen, setIsDrawerOpen] = useState(false);
  const [lastUpdateLabel, setLastUpdateLabel] = useState('No file loaded');

  const seats = useMemo(() => buildSeats(records).map(classifySeat), [records]);
  const seatsByNo = useMemo(() => new Map(seats.map((seat) => [seat.seatNo, seat])), [seats]);

  const counts = useMemo(() => {
    return seats.reduce(
      (accumulator, seat) => {
        accumulator[seat.status] += 1;
        return accumulator;
      },
      { vacant: 0, booked: 0, due_soon: 0, overdue: 0 },
    );
  }, [seats]);

  const activeStudentCount = useMemo(
    () => records.filter((record) => record.active).length,
    [records],
  );

  const activeWithoutValidSeatCount = useMemo(
    () =>
      records.filter(
        (record) => record.active && !(record.seatNo >= 1 && record.seatNo <= TOTAL_SEATS),
      ).length,
    [records],
  );

  const showSeat = (seat: SeatViewModel) => {
    setSelectedSeat(seat);
    setIsDrawerOpen(true);
  };

  useEffect(() => {
    const loadBundledSheet = async () => {
      try {
        for (const path of ['/Students_Final.xlsx', '/Students.xlsx']) {
          const response = await fetch(path, { cache: 'no-store' });
          if (!response.ok) continue;

          const parsed = parseExcelBuffer(await response.arrayBuffer());
          if (!parsed.length) continue;

          setRecords(parsed);
          setLastUpdateLabel(`Auto-loaded ${path.slice(1)}`);
          break;
        }
      } catch {
        // Ignore auto-load errors and let user upload manually.
      }
    };

    void loadBundledSheet();
  }, []);

  const handleUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const parsed = await parseExcelFile(file);
    setRecords(parsed);
    setSelectedSeat(null);
    setIsDrawerOpen(false);
    setLastUpdateLabel(`Updated from ${file.name}`);
    event.target.value = '';
  };

  const resetDemo = () => {
    setRecords(demoData);
    setSelectedSeat(null);
    setIsDrawerOpen(false);
    setLastUpdateLabel('Sample data loaded');
  };

  return (
    <main className="shell">
      <section className="top-summary">
        <section className="stats-row" aria-label="Seat summary">
          <StatCard label="Total seats" value={TOTAL_SEATS} tone="neutral" />
          <StatCard label="Active students" value={activeStudentCount} tone="green" />
          <StatCard label="Booked" value={counts.booked} tone="green" />
          <StatCard label="Due soon" value={counts.due_soon} tone="yellow" />
          <StatCard label="Overdue" value={counts.overdue} tone="red" />
        </section>

        {activeWithoutValidSeatCount > 0 ? (
          <section className="seat-warning" aria-live="polite">
            {activeWithoutValidSeatCount} active student
            {activeWithoutValidSeatCount === 1 ? '' : 's'}{' '}
            {activeWithoutValidSeatCount === 1 ? 'has' : 'have'} no valid seat number (1-50), so{' '}
            {activeWithoutValidSeatCount === 1 ? 'this row is' : 'these rows are'} not shown on the
            seat map.
          </section>
        ) : null}
      </section>

      <section className="hero">
        <div className="hero__copy">
          <p className="eyebrow">FriendsDigital Library</p>
          <h1>Seat allotment that updates from Excel in real time.</h1>
          <p className="lead">
            Track 50 seats, fee timing, and student active status with a polished dashboard built
            for laptops, tablets, and mobile screens.
          </p>
        </div>

        <aside className={`details-drawer details-drawer--hero ${isDrawerOpen ? 'is-open' : ''}`} aria-live="polite">
          {selectedSeat ? (
            <>
              <div className={`details-drawer__badge details-drawer__badge--${selectedSeat.status}`}>
                {selectedSeat.statusLabel}
              </div>
              <h3>Seat {selectedSeat.seatNo}</h3>
              <dl className="details-list">
                <DetailRow label="Student" value={selectedSeat.studentName || 'Vacant'} />
                <DetailRow label="Joining date" value={selectedSeat.joiningDate || 'Not set'} />
                <DetailRow label="Paid through" value={selectedSeat.paidThroughMonth || 'Not marked'} />
                <DetailRow label="Fee due" value={selectedSeat.feeDueDate || 'Not set'} />
                <DetailRow label="Due status" value={formatRelative(selectedSeat.dueInDays)} />
                <DetailRow label="Active" value={selectedSeat.active ? 'Yes' : 'No'} />
                <DetailRow label="Phone" value={selectedSeat.phone || 'Not set'} />
                <DetailRow label="Notes" value={selectedSeat.notes || 'No notes'} />
              </dl>
              <button className="button button--ghost details-drawer__close" onClick={() => setIsDrawerOpen(false)} type="button">
                Close
              </button>
            </>
          ) : (
            <div className="details-empty">
              <p className="section-label">Seat details</p>
              <h3>Hover or tap a seat</h3>
              <p>The seat card opens here with student details and fee timing.</p>
            </div>
          )}
        </aside>
      </section>

      <section className="library-stage">
        <div className="stage__header">
          <div>
            <p className="section-label">Seat layout</p>
            <h2>50-seat reading hall</h2>
          </div>
          <div className="legend">
            <LegendItem color="green" label="Booked" />
            <LegendItem color="yellow" label="Due soon" />
            <LegendItem color="red" label="Overdue" />
            <LegendItem color="grey" label="Vacant" />
          </div>
        </div>

        <div className="room">
          <div className="room__glow room__glow--left" />
          <div className="room__glow room__glow--right" />
          <div className="room__rails" />
          <div className="room__desk">Library desk</div>

          <div className="seat-rows" aria-label="Seat rows">
            {SEAT_LAYOUT_ROWS.map((row) => (
              <div
                key={`row-${row.left[0]}-${row.right[0]}`}
                className="seat-row"
              >
                {row.left.map((seatNo) => {
                  const seat = seatsByNo.get(seatNo);
                  return seat ? <SeatTile key={seatNo} seat={seat} onShow={showSeat} /> : null;
                })}
                <div className="seat-row__aisle" aria-hidden="true" />
                {row.right.map((seatNo) => {
                  const seat = seatsByNo.get(seatNo);
                  return seat ? <SeatTile key={seatNo} seat={seat} onShow={showSeat} /> : null;
                })}
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="bottom-actions" aria-label="Actions">
        <label className="button button--primary">
          Upload Excel
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} />
        </label>
        <button className="button button--ghost" onClick={resetDemo} type="button">
          Reset sample data
        </button>
        <div className="status-pill">{lastUpdateLabel}</div>
      </section>
    </main>
  );
}

function SeatTile({ seat, onShow }: { seat: SeatViewModel; onShow: (seat: SeatViewModel) => void }) {
  return (
    <button
      className={`seat seat--${seat.status}`}
      type="button"
      onMouseEnter={() => onShow(seat)}
      onFocus={() => onShow(seat)}
      onClick={() => onShow(seat)}
      aria-label={`Seat ${seat.seatNo}, ${seat.statusLabel}`}
    >
      <span className="seat__number">{seat.seatNo}</span>
      <span className="seat__indicator" />
    </button>
  );
}

function StatCard({ label, value, tone }: { label: string; value: number; tone: 'neutral' | 'green' | 'yellow' | 'red' }) {
  return (
    <article className={`stat-card stat-card--${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function LegendItem({ color, label }: { color: 'green' | 'yellow' | 'red' | 'grey'; label: string }) {
  return (
    <div className="legend-item">
      <span className={`legend-item__swatch legend-item__swatch--${color}`} />
      <span>{label}</span>
    </div>
  );
}

function DetailRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="detail-row">
      <dt>{label}</dt>
      <dd>{value}</dd>
    </div>
  );
}