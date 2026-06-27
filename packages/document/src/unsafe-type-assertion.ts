export function unsafeOoxmlBoundaryAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Document OOXML reader boundaries narrow parsed XML output behind this named helper.
  return value as T;
}

export function unsafeBrandAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Document branded value constructors narrow primitive values behind this named helper.
  return value as T;
}

export function unsafeFixtureAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Document tests narrow fixture values behind this test-only helper.
  return value as T;
}
