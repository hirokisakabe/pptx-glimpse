export function unsafeOoxmlBoundaryAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

export function unsafeBrandAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

export function unsafeFixtureAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

function assertUnsafeBoundary<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Document OOXML, brand, and fixture helpers intentionally centralize unchecked narrowing here.
  return value as T;
}
