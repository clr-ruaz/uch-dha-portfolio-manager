import * as React from "react";
import styles from "./Dashboard.module.scss";
import { ExpirationFilter, expirationBucket, expirationDateLabel, matchesExpiration, storedCalendarDay } from "./hapExpirationDates";

export interface HapExpirationRecord {
  id: number;
  resident: string;
  property: string;
  unit: string;
  start: unknown;
  end: unknown;
  documentUrl?: string;
}

interface Props {
  records: HapExpirationRecord[];
  today: number;
  loading: boolean;
  error: string;
  filterKey: string;
  onOpen: (id: number) => void;
  filter: ExpirationFilter;
  onFilterChange: (filter: ExpirationFilter) => void;
}

const labels: { [key in ExpirationFilter]: string } = {
  upcoming: "All upcoming (1–90 days)", expired: "Expired", today: "Expiring today",
  "1-30": "1–30 days", "31-60": "31–60 days", "61-90": "61–90 days",
};
const buckets: ExpirationFilter[] = ["1-30", "31-60", "61-90"];
const otherFilters: ExpirationFilter[] = ["upcoming", "today", "expired"];
type SortKey = "resident" | "property" | "unit" | "start" | "end";

export default function HapExpirations(props: Props): React.ReactElement {
  const { filter, onFilterChange } = props;
  const [page, setPage] = React.useState(1);
  const [pageSize, setPageSize] = React.useState(10);
  const [sort, setSort] = React.useState<{ key: SortKey; desc: boolean }>({ key: "end", desc: false });
  React.useEffect(() => { setPage(1); }, [props.filterKey, filter, pageSize, sort]);
  const records = props.records.map((record) => {
    const end = storedCalendarDay(record.end);
    const days = end === undefined ? undefined : end - props.today;
    return { ...record, days, bucket: expirationBucket(days) };
  });
  const unknown = records.filter((record) => record.days === undefined).length;
  const count = (key: ExpirationFilter): number => records.filter((record) => matchesExpiration(record.bucket, key)).length;
  const filtered = records.filter((record) => matchesExpiration(record.bucket, filter)).sort((a, b) => {
    let difference: number;
    if (sort.key === "start" || sort.key === "end") {
      const left = storedCalendarDay(a[sort.key]);
      const right = storedCalendarDay(b[sort.key]);
      if (left === undefined && right !== undefined) return 1;
      if (right === undefined && left !== undefined) return -1;
      difference = (left || 0) - (right || 0);
    } else difference = a[sort.key].localeCompare(b[sort.key], undefined, { numeric: true });
    return difference * (sort.desc ? -1 : 1) || a.resident.localeCompare(b.resident) || a.id - b.id;
  });
  const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const currentPage = Math.min(page, pages);
  const rows = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const unavailable = props.loading || Boolean(props.error);
  const filterButton = (key: ExpirationFilter): React.ReactElement => (
    <button key={key} type="button" aria-pressed={filter === key} disabled={unavailable}
      onClick={() => onFilterChange(key)}>
      <span>{labels[key]}</span><strong>{unavailable ? "—" : count(key)}</strong>
    </button>
  );
  const headings: Array<{ key: SortKey; label: string }> = [
    { key: "resident", label: "Resident" }, { key: "property", label: "Property" },
    { key: "unit", label: "Unit" }, { key: "start", label: "HAP start" }, { key: "end", label: "HAP end" },
  ];
  return <section className={styles.tableCard} aria-busy={props.loading}>
    <div className={`${styles.tableHeader} ${styles.expirationHeader}`}>
      <div><h2>HAP Expirations <span>({unavailable ? "—" : filtered.length})</span></h2><p>Recorded HAP end dates · As of {expirationDateLabel(new Date(props.today * 86400000).toISOString())} (Dallas time)</p></div>
      <div className={`${styles.tableActions} ${styles.expirationFilters}`} role="group" aria-label="Expiration views">{buckets.concat(otherFilters).map(filterButton)}</div>
    </div>
    <p className={styles.expirationNote} role="status">{unavailable ? "" : `${unknown} records have missing or invalid HAP end dates and are excluded from these buckets.`}</p>
    {props.loading ? <p className={styles.empty} role="status">Loading HAP expirations…</p> : props.error ?
      <p className={styles.error} role="alert">Unable to load HAP expirations. {props.error} Use Refresh Data to retry.</p> : <>
      <div className={styles.scroll}>
        <table className={styles.expirationTable} aria-label={`HAP Expirations — ${labels[filter]}`}>
          <thead><tr>{headings.map((heading) => <th key={heading.key} scope="col" className={heading.key === "resident" ? styles.stickyColumn : undefined}
            aria-sort={sort.key === heading.key ? (sort.desc ? "descending" : "ascending") : "none"}>
            <button type="button" onClick={() => setSort({ key: heading.key, desc: sort.key === heading.key && !sort.desc })}>
              {heading.label}{sort.key === heading.key ? (sort.desc ? " ↓" : " ↑") : " ↕"}
            </button>
          </th>)}<th scope="col"><span>Days remaining</span></th><th scope="col"><span>Expiration bucket</span></th><th scope="col"><span>Contract</span></th></tr></thead>
          <tbody>{rows.map((record) => <tr key={record.id} onClick={() => props.onOpen(record.id)}>
            <td className={styles.stickyColumn}><button type="button" className={styles.expirationResident} onClick={(event) => { event.stopPropagation(); props.onOpen(record.id); }} aria-label={`Open resident ${record.resident}`}><strong>{record.resident}</strong></button><small>{record.unit || "Unit"} · {record.property || "Property"}</small></td>
            <td>{record.property || "—"}</td><td>{record.unit || "—"}</td>
            <td>{expirationDateLabel(record.start)}</td><td>{expirationDateLabel(record.end)}</td>
            <td>{record.days! < 0 ? `${Math.abs(record.days!)} days overdue` : record.days}</td>
            <td>{record.bucket ? labels[record.bucket] : "—"}</td>
            <td>{record.documentUrl ? <a href={record.documentUrl} target="_blank" rel="noreferrer" data-interception="off" onClick={(event) => event.stopPropagation()}>View contract documents</a> : "Not uploaded"}</td>
          </tr>)}{!rows.length && <tr><td colSpan={8} className={styles.empty}>No HAP contracts match this expiration view and the selected filters.</td></tr>}</tbody>
        </table>
      </div>
      <footer><div className={styles.pageSummary}>
        <span>Showing {filtered.length ? (currentPage - 1) * pageSize + 1 : 0}–{Math.min(currentPage * pageSize, filtered.length)} of {filtered.length} records</span>
        <label className={styles.pageSize}>Rows per page <select value={pageSize} onChange={(event) => setPageSize(Number(event.target.value))}>{[10, 20, 50, 100].map((size) => <option key={size} value={size}>{size}</option>)}</select></label>
      </div><div>
        <button type="button" disabled={currentPage === 1} onClick={() => setPage(1)}>First</button>
        <button type="button" disabled={currentPage === 1} onClick={() => setPage(currentPage - 1)}>Previous</button>
        <b>Page {currentPage} of {pages}</b>
        <button type="button" disabled={currentPage === pages} onClick={() => setPage(currentPage + 1)}>Next</button>
        <button type="button" disabled={currentPage === pages} onClick={() => setPage(pages)}>Last</button>
      </div></footer>
    </>}
  </section>;
}
