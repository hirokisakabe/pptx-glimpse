export function unsafeTypeAssertion<T>(value: unknown): T {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- Renderer adapter and external-library boundaries intentionally centralize unchecked narrowing here.
  return value as T;
}
