declare module 'node:test' {
  export default function test(name: string, fn: () => void): void;
}

declare module 'node:assert/strict' {
  interface StrictAssert {
    equal(actual: unknown, expected: unknown, message?: string): void;
    deepEqual(actual: unknown, expected: unknown, message?: string): void;
    throws(fn: () => unknown, expected?: RegExp, message?: string): void;
  }

  const assert: StrictAssert;
  export default assert;
}
