/** Bounded per-kind formatter cache; independent of UI locale. */
const limit = 128;
const numbers = new Map<string, Intl.NumberFormat>();
const dates = new Map<string, Intl.DateTimeFormat>();
function cached<T>(cache: Map<string, T>, key: string, create: () => T): T {
  const existing = cache.get(key);
  if (existing !== undefined) {
    cache.delete(key); cache.set(key, existing); return existing;
  }
  const formatter = create();
  if (cache.size >= limit) {
    const oldest = cache.keys().next().value;
    if (oldest !== undefined) cache.delete(oldest);
  }
  cache.set(key, formatter); return formatter;
}
export function numberFormatter(locale: string, options: Intl.NumberFormatOptions = {}): Intl.NumberFormat {
  return cached(numbers, locale + ':' + JSON.stringify(options), () => new Intl.NumberFormat(locale, options));
}
export function dateFormatter(locale: string, options: Intl.DateTimeFormatOptions = {}): Intl.DateTimeFormat {
  return cached(dates, locale + ':' + JSON.stringify(options), () => new Intl.DateTimeFormat(locale, options));
}
