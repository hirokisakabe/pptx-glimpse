export function unsafeXmlBoundaryAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

export function unsafeFixtureAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

export function unsafeAdapterBoundaryAssertion<T>(value: unknown): T {
  return assertUnsafeBoundary<T>(value);
}

function assertUnsafeBoundary<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Core XML, fixture, and adapter helpers intentionally centralize unchecked narrowing here.
  return value as T;
}
