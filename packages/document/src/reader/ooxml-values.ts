import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";

export function parseEnumValue<const T extends string>(
  value: string | undefined,
  allowed: ReadonlySet<T>,
): T | undefined {
  return value !== undefined && allowed.has(unsafeTypeAssertion<T>(value))
    ? unsafeTypeAssertion<T>(value)
    : undefined;
}

export function parseEnumValueWithDefault<const T extends string>(
  value: string | undefined,
  allowed: ReadonlySet<T>,
  fallback: T,
): T {
  return parseEnumValue(value, allowed) ?? fallback;
}
