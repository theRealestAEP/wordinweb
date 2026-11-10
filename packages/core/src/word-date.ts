/**
 * Word writes w:date as the author's local wall-clock time with a "Z" suffix,
 * although the time is not UTC: at 11:01 in California it writes
 * "2026-11-10T11:01:00Z". Word keeps the true UTC time in a separate dateUtc
 * attribute (w16du:dateUtc on a revision, w16cex:dateUtc in
 * commentsExtensible.xml). Word reads w:date the same way, so a UTC time in
 * w:date shows in Word shifted by the author's UTC offset.
 */

/** A w:date value for `at`, written as Word writes it: local time, whole seconds, "Z". */
export function wordDate(at: Date = new Date()): string {
  const two = (n: number) => String(n).padStart(2, "0");
  return `${at.getFullYear()}-${two(at.getMonth() + 1)}-${two(at.getDate())}T${two(at.getHours())}:${two(at.getMinutes())}:${two(at.getSeconds())}Z`;
}

/** The moment a w:date names: dateUtc when the file has it, else w:date read as local time. */
export function readWordDate(date: string | undefined, dateUtc?: string): Date | null {
  if (dateUtc) {
    const utc = new Date(dateUtc);
    if (!isNaN(utc.getTime())) return utc;
  }
  const m = date?.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return null;
  const local = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] ?? 0));
  return isNaN(local.getTime()) ? null : local;
}
