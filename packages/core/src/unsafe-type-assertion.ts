export function unsafeXmlBoundaryAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Core XML parser boundaries narrow fast-xml-parser output behind this named helper.
  return value as T;
}

export function unsafeFixtureAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Core tests narrow fixture values behind this test-only helper.
  return value as T;
}

export function unsafeBrandAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Core branded value constructors narrow primitive values behind this named helper.
  return value as T;
}
