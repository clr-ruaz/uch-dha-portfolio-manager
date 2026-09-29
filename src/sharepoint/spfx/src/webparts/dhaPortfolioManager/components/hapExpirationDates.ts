const dayMilliseconds = 86400000;

export type ExpirationBucket = "expired" | "today" | "1-30" | "31-60" | "61-90";
export type ExpirationFilter = ExpirationBucket | "upcoming";

// SharePoint date fields represent a calendar date, not an instant to shift to local time.
export function storedCalendarDay(value: unknown): number | undefined {
  if (typeof value !== "string") return undefined;
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:$|T)/.exec(value);
  if (!match || (value.length > 10 && !isFinite(Date.parse(value)))) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const stamp = Date.UTC(year, month - 1, day);
  const parsed = new Date(stamp);
  return parsed.getUTCFullYear() === year && parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day ? stamp / dayMilliseconds : undefined;
}

export function dallasToday(now: Date = new Date()): number {
  const date = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Chicago", year: "numeric", month: "2-digit", day: "2-digit",
  }).format(now).split("/");
  return Date.UTC(Number(date[2]), Number(date[0]) - 1, Number(date[1])) / dayMilliseconds;
}

export function expirationBucket(days: number | undefined): ExpirationBucket | undefined {
  if (days === undefined || !isFinite(days)) return undefined;
  if (days < 0) return "expired";
  if (days === 0) return "today";
  if (days <= 30) return "1-30";
  if (days <= 60) return "31-60";
  if (days <= 90) return "61-90";
  return undefined;
}

export function matchesExpiration(bucket: ExpirationBucket | undefined, filter: ExpirationFilter): boolean {
  return filter === "upcoming"
    ? bucket === "1-30" || bucket === "31-60" || bucket === "61-90"
    : bucket === filter;
}

export function expirationDateLabel(value: unknown): string {
  const day = storedCalendarDay(value);
  return day === undefined ? "—" : new Date(day * dayMilliseconds).toLocaleDateString("en-US", { timeZone: "UTC" });
}
