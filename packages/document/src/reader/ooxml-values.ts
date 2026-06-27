export function parseEnumValue<const T extends string>(
  value: string | undefined,
  allowed: ReadonlySet<T>,
): T | undefined {
  return value !== undefined && allowed.has(value as T) ? (value as T) : undefined;
}

export function parseEnumValueWithDefault<const T extends string>(
  value: string | undefined,
  allowed: ReadonlySet<T>,
  fallback: T,
): T {
  return parseEnumValue(value, allowed) ?? fallback;
}
