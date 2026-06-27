export function unsafeExternalInteropAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Renderer external-library interop narrows untyped library values behind this named helper.
  return value as T;
}

export function unsafeBrandAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Renderer branded value constructors narrow primitive values behind this named helper.
  return value as T;
}

export function unsafeFixtureAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Renderer tests narrow fixture values behind this test-only helper.
  return value as T;
}
