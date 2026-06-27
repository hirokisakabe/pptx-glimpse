export function assertNever(
  value: never,
  message = "Unexpected discriminated union member",
): never {
  throw new Error(`${message}: ${JSON.stringify(value)}`);
}
